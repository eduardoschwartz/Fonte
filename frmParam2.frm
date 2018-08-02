VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmParam2 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4560
   ClientLeft      =   1680
   ClientTop       =   3300
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6195
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4440
      TabIndex        =   49
      ToolTipText     =   "Sair da Tela"
      Top             =   4140
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParam2.frx":0000
      PICN            =   "frmParam2.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   3375
      TabIndex        =   50
      ToolTipText     =   "Gravar os Dados"
      Top             =   4155
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmParam2.frx":008A
      PICN            =   "frmParam2.frx":00A6
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
      Left            =   4440
      TabIndex        =   45
      ToolTipText     =   "Cancelar Edição"
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmParam2.frx":044B
      PICN            =   "frmParam2.frx":0467
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   240
      TabIndex        =   46
      ToolTipText     =   "Novo Registro"
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmParam2.frx":05C1
      PICN            =   "frmParam2.frx":05DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1290
      TabIndex        =   47
      ToolTipText     =   "Editar Registro"
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmParam2.frx":0737
      PICN            =   "frmParam2.frx":0753
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
      Left            =   2340
      TabIndex        =   48
      ToolTipText     =   "Excluir Registro"
      Top             =   4140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmParam2.frx":08AD
      PICN            =   "frmParam2.frx":08C9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00EEEEEE&
      Height          =   990
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   3030
      Width           =   6120
      Begin VB.TextBox txtFator 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4470
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Width           =   1005
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   1005
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   5
         Top             =   570
         Width           =   4005
      End
      Begin VB.Label lblFator 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator........:"
         Height          =   195
         Left            =   3630
         TabIndex        =   7
         Top             =   285
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição...........:"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Width           =   1275
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00EEEEEE&
      Height          =   990
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   3030
      Width           =   6090
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1470
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtMin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   17
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtFator2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   15
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtMax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         MaxLength       =   50
         TabIndex        =   19
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator..................:"
         Height          =   195
         Left            =   3030
         TabIndex        =   20
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequência.........:"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "m² Mínimo.........:"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   16
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "m² Máximo.........:"
         Height          =   195
         Index           =   5
         Left            =   3030
         TabIndex        =   14
         Top             =   600
         Width           =   1245
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00EEEEEE&
      Height          =   990
      Index           =   2
      Left            =   0
      TabIndex        =   24
      Top             =   3030
      Width           =   6090
      Begin VB.TextBox txtValorM2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4380
         MaxLength       =   50
         TabIndex        =   27
         Top             =   570
         Width           =   1455
      End
      Begin VB.ComboBox cmbAgrup 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   210
         Width           =   1095
      End
      Begin VB.ComboBox cmbMoeda 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do m²........:"
         Height          =   195
         Index           =   6
         Left            =   3090
         TabIndex        =   31
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Terrenos (valor / m²)"
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
         Left            =   3060
         TabIndex        =   30
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Moeda..:"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   29
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agrupamento.....:"
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   28
         Top             =   300
         Width           =   1245
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00EEEEEE&
      Height          =   990
      Index           =   3
      Left            =   0
      TabIndex        =   38
      Top             =   3030
      Width           =   6090
      Begin VB.TextBox txtFatorCateg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4410
         MaxLength       =   50
         TabIndex        =   41
         Top             =   540
         Width           =   1455
      End
      Begin VB.ComboBox cmbMoeda3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   510
         Width           =   1455
      End
      Begin VB.TextBox txtCateg 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   255
         Left            =   1470
         MaxLength       =   50
         TabIndex        =   39
         Top             =   180
         Width           =   2505
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Categoria..:"
         Height          =   195
         Index           =   10
         Left            =   3120
         TabIndex        =   44
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Moeda..:"
         Height          =   195
         Index           =   15
         Left            =   180
         TabIndex        =   43
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria...........:"
         Height          =   225
         Index           =   16
         Left            =   180
         TabIndex        =   42
         Top             =   210
         Width           =   1245
      End
   End
   Begin VB.Frame FraCat 
      BackColor       =   &H00EEEEEE&
      Height          =   525
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   6105
      Begin VB.ComboBox cmbUso 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmParam2.frx":096B
         Left            =   630
         List            =   "frmParam2.frx":096D
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   150
         Width           =   1755
      End
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmParam2.frx":096F
         Left            =   2820
         List            =   "frmParam2.frx":0971
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   150
         Width           =   1845
      End
      Begin VB.ComboBox cmbAno2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmParam2.frx":0973
         Left            =   4830
         List            =   "frmParam2.frx":0975
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   150
         Width           =   855
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Uso:"
         Height          =   195
         Left            =   270
         TabIndex        =   37
         Top             =   210
         Width           =   315
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   2430
         TabIndex        =   36
         Top             =   210
         Width           =   375
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdAno 
      Height          =   2505
      Left            =   0
      TabIndex        =   23
      Top             =   510
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^Seq     |>m² Mínimo            |>m² Máximo             |>Fator          "
   End
   Begin MSFlexGridLib.MSFlexGrid grdProf 
      Height          =   2505
      Left            =   0
      TabIndex        =   11
      Top             =   510
      Visible         =   0   'False
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^Seq     |>m² Mínimo            |>m² Máximo             |>Fator          "
   End
   Begin VB.Frame fraAno 
      BackColor       =   &H00EEEEEE&
      Height          =   525
      Left            =   0
      TabIndex        =   8
      Top             =   -30
      Width           =   6105
      Begin VB.ComboBox cmbDist 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3210
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   150
         Width           =   2745
      End
      Begin VB.ComboBox cmbAno 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmParam2.frx":0977
         Left            =   960
         List            =   "frmParam2.frx":0979
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   150
         Width           =   1335
      End
      Begin VB.Label lblDist 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito.:"
         Height          =   195
         Left            =   2580
         TabIndex        =   22
         Top             =   210
         Width           =   555
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano........:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   210
         Width           =   795
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   2505
      Left            =   0
      TabIndex        =   0
      Top             =   510
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   4419
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Código       |<Descricão                                                     |>Fator             "
   End
End
Attribute VB_Name = "frmParam2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim Sql As String

Dim Evento As String, bExec As Boolean
Dim sRet As String
'Dim evCnPE As Integer, evCnSI As Integer, evCnTO As Integer
'Dim evCnDI As Integer, evCnFB As Integer, evCnPR As Integer
'Dim evCnGL As Integer, evCnPG As Integer, evCnFC As Integer
'
'Dim evNewPE As Integer, evNewSI As Integer, evNewFC As Integer
'Dim evNewTO As Integer, evNewDI As Integer, evNewFB As Integer
'Dim evNewPR As Integer, evNewGL As Integer, evNewPG As Integer
'
'Dim evEditPE As Integer, evEditSI As Integer, evEditFC As Integer
'Dim evEditTO As Integer, evEditDI As Integer, evEditFB As Integer
'Dim evEditPr As Integer, evEditGL As Integer, evEditPG As Integer
'
'Dim evDelPE As Integer, evDelSI As Integer, evDelFC As Integer
'Dim evDelTO As Integer, evDelDI As Integer, evDelFB As Integer
'Dim evDelPR As Integer, evDelGL As Integer, evDelPG As Integer
'
'Dim bEvCnPE As Boolean, bEvCnSI As Boolean, bEvCnFC As Boolean
'Dim bEvCnTO As Boolean, bEvCnDI As Boolean, bEvCnFB As Boolean
'Dim bEvCnPR As Boolean, bEvCnGL As Boolean, bEvCnPG As Boolean
'
'Dim bEvNewPE As Boolean, bEvNewSI As Boolean, bEvNewFC As Boolean
'Dim bEvNewTO As Boolean, bEvNewDI As Boolean, bEvNewFB As Boolean
'Dim bEvNewPR As Boolean, bEvNewGL As Boolean, bEvNewPG As Boolean
'
'Dim bEvEditPE As Boolean, bEvEditSI As Boolean, bEvEditFC As Boolean
'Dim bEvEditTO As Boolean, bEvEditDI As Boolean, bEvEditFB As Boolean
'Dim bEvEditPR As Boolean, bEvEditGL As Boolean, bEvEditPG As Boolean
'
'Dim bEvDelPE As Boolean, bEvDelSI As Boolean, bEvDelFC As Boolean
'Dim bEvDelTO As Boolean, bEvDelDI As Boolean, bEvDelFB As Boolean
'Dim bEvDelPR As Boolean, bEvDelGL As Boolean, bEvDelPG As Boolean

Private Sub cmbAno_Click()
On Error Resume Next
If cmbAno.ListIndex = -1 Or Not bExec Then Exit Sub
CarregaLista

End Sub

Private Sub cmbAno2_Click()
On Error Resume Next
If cmbAno2.ListIndex = -1 Or Not bExec Then Exit Sub
CarregaLista

End Sub

Private Sub cmbDist_Click()
If cmbDist.ListIndex = -1 Or Not bExec Then Exit Sub
CarregaLista
End Sub

Private Sub cmbMoeda_Click()
Dim nLin As Integer, nCol As Integer
nLin = grdAno.Row
nCol = cmbMoeda.ItemData(cmbMoeda.ListIndex)
txtValorM2.Text = FormatNumber(grdAno.TextMatrix(nLin, nCol), 2)
End Sub

Private Sub cmbMoeda3_Click()
Dim nLin As Integer, nCol As Integer
On Error Resume Next
nLin = grdAno.Row
nCol = cmbMoeda3.ItemData(cmbMoeda3.ListIndex)
txtFatorCateg.Text = FormatNumber(grdAno.TextMatrix(nLin, nCol), 2)
txtFatorCateg.SetFocus
End Sub

Private Sub cmbTipo_Click()
On Error Resume Next
If cmbTipo.ListIndex = -1 Or Not bExec Then Exit Sub
CarregaLista

End Sub

Private Sub cmbUso_Click()
On Error Resume Next
If cmbUso.ListIndex = -1 Or Not bExec Then Exit Sub
CarregaLista

End Sub

Private Sub cmdAlterar_Click()
    
Select Case sParamForm
      Case "DIST", "SITU", "TOPO", "PEDO"
            If txtCod.Text = "" Then
               MsgBox "Não existem Registros.", vbCritical, "Atenção"
               Exit Sub
            End If
            sOldDesc = txtDesc.Text
      Case "FGLE", "FPRO"
            If txtSeq.Text = "" Then
               MsgBox "Não existem Registros.", vbCritical, "Atenção"
               Exit Sub
            End If
      Case "FCAT"
            cmbTipo.Enabled = False
            cmbUso.Enabled = False
            cmbAno2.Enabled = False
            If txtCateg.Text = "" Then
               MsgBox "Não existem Registros.", vbCritical, "Atenção"
               Exit Sub
            End If
End Select

Eventos "INCLUIR"
Evento = "Alterar"

End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""

End Sub

Private Sub cmdExcluir_Click()

Select Case sParamForm
      Case "DIST", "SITU", "TOPO", "PEDO"
            If txtCod.Text = "" Then
               MsgBox "Não existem Registros.", vbCritical, "Atenção"
               Exit Sub
            End If
            sOldDesc = txtDesc.Text
      Case "FGLE", "FPRO"
            If txtSeq.Text = "" Then
               MsgBox "Não existem Registros.", vbCritical, "Atenção"
               Exit Sub
            End If
End Select
    
    If MsgBox("Excluir este Registro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
        Select Case sParamForm
             Case "PEDO"
                 Sql = "DELETE FROM FATORPEDOLOGIA WHERE  CODPEDOLOGIA=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
                 Sql = "DELETE FROM PEDOLOGIA WHERE  CODPEDOLOGIA=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
             Case "SITU"
                 Sql = "DELETE FROM FATORSITUACAO WHERE  CODSITUACAO=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
                 Sql = "DELETE FROM SITUACAO WHERE CODSITUACAO=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
             Case "TOPO"
                 Sql = "DELETE FROM FATORTOPOGRAFIA WHERE  CODTOPOG=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
                 Sql = "DELETE FROM TOPOGRAFIA WHERE CODTOPOGRAFIA=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
             Case "DIST"
                 Sql = "DELETE FROM FATORDISTRITO WHERE  CODDISTRITO=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
                 Sql = "DELETE FROM DISTRITO WHERE CODDISTRITO=" & txtCod.Text
                 cn.Execute Sql, rdExecDirect
             Case "FPRO"
                 Sql = "DELETE FROM FATORPROFUN WHERE  CODPROFUN=" & txtSeq.Text
                 cn.Execute Sql, rdExecDirect
                 Sql = "DELETE FROM PROFUNDIDADE WHERE  CODDISTRITO=" & cmbDist.ItemData(cmbDist.ListIndex) & " AND CODPROFUN=" & txtSeq.Text
                 cn.Execute Sql, rdExecDirect
             Case "FGLE"
                 Sql = "DELETE FROM FATORGLEBA WHERE  CODGLEBA=" & txtSeq.Text
                 cn.Execute Sql, rdExecDirect
                 Sql = "DELETE FROM GLEBA WHERE CODGLEBA=" & txtSeq.Text
                 cn.Execute Sql, rdExecDirect
        End Select
       
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtDesc.Text
       Limpa
       CarregaLista
       Le
    End If

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer
Dim nOldMin As Double, nOldMax As Double
Dim nNewMin As Double, nNewMax As Double

If Trim$(txtFator.Text) = "" Then txtFator.Text = 0

Select Case sParamForm
      Case "DIST", "SITU", "TOPO", "PEDO"
            If txtDesc.Text = "" Then
               MsgBox "Favor digitar a Descrição.", vbExclamation, "Atenção"
               txtDesc.SetFocus
               Exit Sub
            End If
      Case "FGLE", "FPRO"
            If Val(txtMin.Text) = 0 Then
               MsgBox "Favor digitar o Valor Mínimo.", vbExclamation, "Atenção"
               txtMin.SetFocus
               Exit Sub
            End If
            If Val(txtMax.Text) = 0 Then
               If grdProf.Rows >= 2 Then
                    If Val(grdProf.TextMatrix(grdProf.Rows - 1, 2)) = 0 And Evento = "Novo" Then
                          MsgBox "Apenas o Último Item pode conter um Valor Máximo igual a Zero." & vbCrLf & vbCrLf & "Se quiser manter este valor como Zero atualize antes o valor máximo do ultimo item cadastrado para um valor diferente de Zero.", vbExclamation, "Atenção"
                          txtMax.SetFocus
                         Exit Sub
                    End If
                End If
            End If
            If Val(txtMax.Text) > 0 Then
               If grdProf.Rows >= 2 Then
                    If Val(grdProf.TextMatrix(grdProf.Rows - 1, 2)) = 0 And Evento = "Novo" Then
                          MsgBox "Não pode haver um valor máximo maior que zero se já existir um valor máximo igual a zero." & vbCrLf & vbCrLf & "Se quiser manter este valor como maior que zero atualize antes o valor=0 para um valor diferente de Zero.", vbExclamation, "Atenção"
                          txtMax.SetFocus
                         Exit Sub
                    End If
                End If
            End If
            If Val(txtMin.Text) >= Val(txtMax.Text) And Val(txtMax.Text) <> 0 Then
               MsgBox "O Valor Máximo deve ser  maior que o valor mínimo.", vbExclamation, "Atenção"
               txtMin.SetFocus
               Exit Sub
            End If
            If Evento = "Novo" Then
                  If txtMin.Text = "" Then txtMin.Text = 0
                  If txtMax.Text = "" Then txtMax.Text = 0
                   nNewMin = txtMin.Text
                   nNewMax = txtMax.Text
                   For x = 1 To grdProf.Rows - 1
                         nOldMin = grdProf.TextMatrix(x, 1)
                         nOldMax = grdProf.TextMatrix(x, 2)
                         If nNewMax > 0 Then
                               If (nNewMin >= nOldMin And nNewMin <= nOldMax) Or (nNewMax >= nOldMin And nNewMax <= nOldMax) Then
                                    MsgBox "Este Intervalo ja coincide com outro intervalo cadastrado. Verifique !!!" & vbCrLf & vbCrLf & "Intervalo em Conflito: (" & nOldMin & " - " & nOldMax & ")", vbExclamation, "Atenção"
                                    txtMin.SetFocus
                                    Exit Sub
                               End If
                         Else
                               If (nNewMin >= nOldMin And nNewMin <= nOldMax) Then
                                    MsgBox "Este Intervalo ja coincide com outro intervalo cadastrado. Verifique !!!" & vbCrLf & vbCrLf & "Intervalo em Conflito: (" & nOldMin & " - " & nOldMax & ")", vbExclamation, "Atenção"
                                    txtMin.SetFocus
                                    Exit Sub
                               End If
                         End If
                   Next
             End If
End Select


Grava
Eventos "INICIAR"

End Sub


Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
    
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()

Ocupado

Select Case sParamForm
       Case "PEDO"
               Me.Caption = "Fator Pedologia"
       Case "SITU"
               Me.Caption = "Fator Situação"
       Case "TOPO"
               Me.Caption = "Fator Topografia"
       Case "DIST"
               Me.Caption = "Fator Distrito"
       Case "FBEN"
               Me.Caption = "Fator Benfeitoria"
       Case "FGLE"
               Me.Caption = "Fator Gleba"
       Case "FPRO"
               Me.Caption = "Fator Profundidade"
       Case "PGEN"
               Me.Caption = "Planta Genérica de Valores"
       Case "FCAT"
               Me.Caption = "Fator Categoria"
End Select

Select Case sParamForm
        Case "PEDO", "SITU", "TOPO", "DIST", "FBEN"
                grdMain.Visible = True
                Fra(0).Visible = True
                grdProf.Visible = False
                Fra(1).Visible = False
                lblDist.Visible = False
                cmbDist.Visible = False
                grdAno.Visible = False
                Fra(2).Visible = False
                Fra(3).Visible = False
                FraCat.Visible = False
        Case "FGLE", "FPRO"
                grdMain.Visible = False
                Fra(0).Visible = False
                grdProf.Visible = True
                Fra(1).Visible = True
                If sParamForm = "FPRO" Then
                    lblDist.Visible = True
                    cmbDist.Visible = True
                Else
                    lblDist.Visible = False
                    cmbDist.Visible = False
                End If
                grdAno.Visible = False
                Fra(2).Visible = False
                Fra(3).Visible = False
                FraCat.Visible = False
        Case "PGEN"
                grdMain.Visible = False
                Fra(0).Visible = False
                grdProf.Visible = False
                Fra(1).Visible = False
                lblDist.Visible = False
                cmbDist.Visible = False
                grdAno.Visible = True
                Fra(2).Visible = True
                Fra(3).Visible = False
                FraCat.Visible = False
                Sql = "SELECT CODMOEDA,DESCMOEDA  FROM MOEDA"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                grdAno.Cols = RdoAux.RowCount + 1
                For x = 0 To grdAno.Cols - 1
                      grdAno.ColWidth(x) = 900
                      grdAno.ColAlignment(x) = flexAlignCenterCenter
                Next
                grdAno.TextMatrix(0, 0) = "Agrup."
                x = 1
                cmbMoeda.Clear
                Do Until RdoAux.EOF
                     grdAno.TextMatrix(0, x) = RdoAux!DESCMOEDA
                     cmbMoeda.AddItem RdoAux!DESCMOEDA
                     cmbMoeda.ItemData(cmbMoeda.NewIndex) = RdoAux!CODMOEDA
                     RdoAux.MoveNext
                     x = x + 1
                Loop
                For x = 1 To 7
                   cmbAgrup.Clear
                   grdAno.TextMatrix(x, 0) = x
                   cmbAgrup.AddItem x
                Next
        Case "FCAT"
                grdMain.Visible = False
                Fra(0).Visible = False
                grdProf.Visible = False
                Fra(1).Visible = False
                lblDist.Visible = False
                cmbDist.Visible = False
                grdAno.Visible = True
                Fra(2).Visible = False
                Fra(3).Visible = True
                FraCat.Visible = True
End Select

Centraliza Me
sRet = RetEventUserForm(Me.Name)
grdMain.Rows = 1

bExec = False
For x = 1997 To Format(Year(Now) + 1, "0000")
    cmbAno.AddItem x
    cmbAno2.AddItem x
Next

Sql = "SELECT CODDISTRITO,DESCDISTRITO FROM DISTRITO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
       Do Until .EOF
            cmbDist.AddItem !DescDistrito
            cmbDist.ItemData(cmbDist.NewIndex) = !CODDISTRITO
           .MoveNext
       Loop
     .Close
End With

Sql = "SELECT CODUSOCONSTR,DESCUSOCONSTR FROM USOCONSTR WHERE CODUSOCONSTR<>999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
       Do Until .EOF
            cmbUso.AddItem !DESCUSOCONSTR
            cmbUso.ItemData(cmbUso.NewIndex) = !CODUSOCONSTR
           .MoveNext
       Loop
     .Close
End With

Sql = "SELECT CODTIPOCONSTR,DESCTIPOCONSTR FROM TIPOCONSTR WHERE CODTIPOCONSTR<>999"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
       Do Until .EOF
            cmbTipo.AddItem !DESCTIPOCONSTR
            cmbTipo.ItemData(cmbTipo.NewIndex) = !CODTIPOCONSTR
           .MoveNext
       Loop
     .Close
End With

cmbUso.ListIndex = 0
cmbTipo.ListIndex = 0
cmbAno.ListIndex = cmbAno.ListCount - 1
cmbAno2.ListIndex = cmbAno2.ListCount - 1
cmbDist.ListIndex = 0

bExec = True
CarregaLista
Le
If grdMain.Rows > 1 Then grdMain_Click
Eventos "INICIAR"

End Sub

Private Sub grdAno_Click()
Limpa
Le
End Sub

Private Sub grdAno_SelChange()
grdAno_Click
End Sub

Private Sub grdMain_Click()
Limpa
If grdMain.Row > 0 Then
     txtCod.Text = grdMain.TextMatrix(grdMain.Row, 0)
     txtDesc.Text = grdMain.TextMatrix(grdMain.Row, 1)
     txtFator.Text = grdMain.TextMatrix(grdMain.Row, 2)
 End If
End Sub

Private Sub grdMain_RowColChange()
grdMain_Click
End Sub

Private Sub grdProf_RowColChange()
Limpa
If grdProf.Row > 0 Then
     txtSeq.Text = grdProf.TextMatrix(grdProf.Row, 0)
     txtMin.Text = grdProf.TextMatrix(grdProf.Row, 1)
     txtMax.Text = grdProf.TextMatrix(grdProf.Row, 2)
     txtFator2.Text = grdProf.TextMatrix(grdProf.Row, 3)
 End If
End Sub

Private Sub Grava()
Dim nLin As Integer, nCol As Integer
Dim x As Integer
Dim MaxCod As Integer
Dim qd As New rdoQuery

On Error Resume Next
RdoAux.Close
On Error GoTo fim
Set qd.ActiveConnection = cn

Select Case sParamForm
     Case "PEDO"
        Sql = "SELECT MAX(CODPEDOLOGIA) AS MAXIMO FROM PEDOLOGIA WHERE CODPEDOLOGIA<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT PEDOLOGIA (CODPEDOLOGIA,DescPedologia) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.Text) & "')"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORPEDOLOGIA (ANOPEDOLOGIA,CODPEDOLOGIA,FATORPEDOLOGIA) VALUES("
            Sql = Sql & cmbAno.Text & "," & MaxCod & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE PEDOLOGIA SET DescPedologia='" & Mask(txtDesc.Text) & "' WHERE "
            Sql = Sql & "CODPEDOLOGIA=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & cmbAno.Text & " AND "
            Sql = Sql & "CODPEDOLOGIA=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORPEDOLOGIA (ANOPEDOLOGIA,CODPEDOLOGIA,FATORPEDOLOGIA) VALUES("
            Sql = Sql & cmbAno.Text & "," & Val(txtCod.Text) & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        End If
     Case "SITU"
        Sql = "SELECT MAX(CODSITUACAO) AS MAXIMO FROM SITUACAO WHERE CODSITUACAO<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT SITUACAO (CODSITUACAO,DescSITUACAO) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.Text) & "')"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORSITUACAO (ANOSITUACAO,CODSITUACAO,FATORSITUACAO) VALUES("
            Sql = Sql & cmbAno.Text & "," & MaxCod & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE SITUACAO SET DescSITUACAO='" & Mask(txtDesc.Text) & "' WHERE "
            Sql = Sql & "CODSITUACAO=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM FATORSITUACAO WHERE ANOSITUACAO=" & cmbAno.Text & " AND "
            Sql = Sql & "CODSITUACAO=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORSITUACAO (ANOSITUACAO,CODSITUACAO,FATORSITUACAO) VALUES("
            Sql = Sql & cmbAno.Text & "," & Val(txtCod.Text) & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        End If
     Case "TOPO"
        Sql = "SELECT MAX(CODTOPOGRAFIA) AS MAXIMO FROM TOPOGRAFIA WHERE CODTOPOGRAFIA<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT TOPOGRAFIA (CODTOPOGRAFIA,DescTOPOGRAFIA) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.Text) & "')"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORTOPOGRAFIA (ANOTOPOG,CODTOPOGRAFIA,FATORTOPOG) VALUES("
            Sql = Sql & cmbAno.Text & "," & MaxCod & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE TOPOGRAFIA SET DESCTOPOGRAFIA='" & Mask(txtDesc.Text) & "' WHERE "
            Sql = Sql & "CODTOPOGRAFIA=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & cmbAno.Text & " AND "
            Sql = Sql & "CODTOPOGRAFIA=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORTOPOG (ANOTOPOG,CODTOPOGRAFIA,FATORTOPOG) VALUES("
            Sql = Sql & cmbAno.Text & "," & Val(txtCod.Text) & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        End If
     Case "DIST"
        Sql = "SELECT MAX(CODDISTRITO) AS MAXIMO FROM DISTRITO WHERE CODDISTRITO<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT DISTRITO (CODDISTRITO,DescDISTRITO) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.Text) & "')"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORDISTRITO (ANODISTRITO,CODDISTRITO,FATORDISTRITO) VALUES("
            Sql = Sql & cmbAno.Text & "," & MaxCod & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE DISTRITO SET DESCDISTRITO='" & Mask(txtDesc.Text) & "' WHERE "
            Sql = Sql & "CODDISTRITO=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM FATORDISTRITO WHERE ANODISTRITO=" & cmbAno.Text & " AND "
            Sql = Sql & "CODDISTRITO=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORDISTRITO (ANODISTRITO,CODDISTRITO,FATORDISTRITO) VALUES("
            Sql = Sql & cmbAno.Text & "," & Val(txtCod.Text) & "," & Virg2Ponto(txtFator.Text) & ")"
            cn.Execute Sql, rdExecDirect
        End If
     Case "FBEN"
        Sql = "Delete From FATORBENFEITORIA Where ANOFATORBENF =cmbAno.text AND "
        Sql = Sql & " CODBENFEITORIA = Val(txtCod.text)"
        cn.Execute Sql, rdExecDirect
        Sql = Sql & "Insert FATORBENFEITORIA (ANOFATORBENF,CODBENFEITORIA,CODTIPOPROP,FATORBENFEITORIA) values("
        Sql = Sql & cmbAno.Text & "," & Left$(grdMain.TextMatrix(grdMain.Row, 0), 1) & "," & Left$(grdMain.TextMatrix(grdMain.Row, 0), 1) & "," & Virg2Ponto(txtFator.Text) & ")"
        cn.Execute Sql, rdExecDirect
     Case "FGLE"
        Sql = "SELECT MAX(CODGLEBA) AS MAXIMO FROM GLEBA"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Trim$(txtMin.Text) = "" Then txtMin.Text = 0
        If Trim$(txtMax.Text) = "" Then txtMax.Text = 0
        If Trim$(txtFator2.Text) = "" Then txtFator2.Text = 0
        If Evento = "Novo" Then
            Sql = "INSERT GLEBA (CODGLEBA,MINGLEBA,MAXGLEBA) VALUES("
            Sql = Sql & MaxCod & "," & Virg2Ponto(txtMin.Text) & "," & Virg2Ponto(txtMax.Text) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORGLEBA (ANOGLEBA,CODGLEBA,FATORGLEBA) VALUES("
            Sql = Sql & cmbAno.Text & "," & MaxCod & "," & Virg2Ponto(txtFator2.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE GLEBA SET MINGLEBA=" & Virg2Ponto(txtMin.Text) & ",MAXGLEBA=" & Virg2Ponto(txtMax.Text) & "' WHERE "
            Sql = Sql & "CODGLEBA=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM FATORGLEBA WHERE ANOGLEBA=" & cmbAno.Text & " AND "
            Sql = Sql & "CODGLEBA=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORGLEBA (ANOGLEBA,CODGLEBA,FATORGLEBA) VALUES("
            Sql = Sql & cmbAno.Text & "," & Val(txtCod.Text) & "," & Virg2Ponto(txtFator2.Text) & ")"
            cn.Execute Sql, rdExecDirect
        End If
     Case "FPRO"
        Sql = "SELECT MAX(CODPROFUN) AS MAXIMO FROM PROFUNDIDADE WHERE CODDISTRITO=" & cmbDist.ItemData(cmbDist.ListIndex)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Trim$(txtMin.Text) = "" Then txtMin.Text = 0
        If Trim$(txtMax.Text) = "" Then txtMax.Text = 0
        If Trim$(txtFator2.Text) = "" Then txtFator2.Text = 0
        If Evento = "Novo" Then
            Sql = "INSERT PROFUNDIDADE (CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN) VALUES("
            Sql = Sql & cmbDist.ItemData(cmbDist.ListIndex) & "," & MaxCod & "," & Virg2Ponto(txtMin.Text) & "," & Virg2Ponto(txtMax.Text) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORPROFUN (CODDISTRITO,CODPROFUN,FATORPROFUN) VALUES("
            Sql = Sql & cmbDist.ItemData(cmbDist.ListIndex) & "," & MaxCod & "," & Virg2Ponto(txtFator2.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE PROFUNDIDADE SET MINPROFUN=" & Virg2Ponto(txtMin.Text) & ",MAXPROFUN=" & Virg2Ponto(txtMax.Text) & "' WHERE "
            Sql = Sql & "CODPROFUN=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM FATORPROFUN WHERE ANOPROFUN=" & cmbAno.Text & " AND "
            Sql = Sql & "CODPROFUN=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT FATORPROFUN (ANOPROFUN,CODPROFUN,FATORPROFUN) VALUES("
            Sql = Sql & cmbAno.Text & "," & Val(txtCod.Text) & "," & Virg2Ponto(txtFator2.Text) & ")"
            cn.Execute Sql, rdExecDirect
        End If
     Case "PGEN"
        If Evento = "Novo" Then
            Sql = "INSERT TERRENO (CODAGRUPAMENTO,ANOFATOR,CODMOEDA,VALORTERRENO) VALUES("
            Sql = Sql & cmbAgrup.Text & "," & cmbAno.Text & "," & cmbMoeda.ItemData(cmbMoeda.ListIndex) & "," & Virg2Ponto(txtValorM2.Text) & ")"
            cn.Execute Sql, rdExecDirect
        Else
            Sql = "UPDATE TERRENO SET VALORTERRENO=" & Virg2Ponto(txtValorM2.Text) & " WHERE CODAGRUPAMENTO="
            Sql = Sql & cmbAgrup.Text & " AND ANOFATOR=" & cmbAno.Text & " AND CODMOEDA=" & cmbMoeda.ItemData(cmbMoeda.ListIndex)
            cn.Execute Sql, rdExecDirect
        End If
     Case "FCAT"
        If Evento = "Novo" Then
            Sql = "INSERT FATORCATEG (CODUSO,CODTIPO,CODCATEG,ANOCATEG,CODMOEDA,FATORCATEG) VALUES("
            Sql = Sql & cmbUso.ItemData(cmbUso.ListIndex) & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & ","
            Sql = Sql & grdAno.Row & "," & cmbAno2.Text & "," & cmbMoeda3.ItemData(cmbMoeda3.ListIndex) & ","
            Sql = Sql & Virg2Ponto(txtFatorCateg.Text) & ")"
        Else
            Sql = "UPDATE FATORCATEG SET FATORCATEG=" & Virg2Ponto(txtFatorCateg.Text) & " WHERE CODUSO=" & cmbUso.ItemData(cmbUso.ListIndex) & " AND "
            Sql = Sql & "CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND CODCATEG=" & grdAno.Row & " AND "
            Sql = Sql & "ANOCATEG=" & cmbAno2.Text & " AND CODMOEDA=" & cmbMoeda3.ItemData(cmbMoeda3.ListIndex)
        End If
        cn.Execute Sql, rdExecDirect
End Select

Select Case sParamForm
      Case "DIST", "SITU", "TOPO", "PEDO"
            If Evento = "Novo" Then
                 grdMain.AddItem MaxCod & Chr(9) & txtDesc.Text & Chr(9) & Format(txtFator.Text, "#0.0000")
            Else
                grdMain.TextMatrix(grdMain.Row, 2) = FormatNumber(txtFator.Text, 4)
            End If
      Case "FGLE", "FPRO"
            If Evento = "Novo" Then
                 grdProf.AddItem MaxCod & Chr(9) & FormatNumber(txtMin.Text, 2) & Chr(9) & FormatNumber(txtMax.Text, 2) & Chr(9) & FormatNumber(txtFator2.Text, 4)
            Else
                grdProf.TextMatrix(grdProf.Row, 1) = FormatNumber(txtMin.Text, 2)
                grdProf.TextMatrix(grdProf.Row, 2) = FormatNumber(txtMax.Text, 2)
                grdProf.TextMatrix(grdProf.Row, 3) = FormatNumber(txtFator2.Text, 4)
            End If
      Case "PGEN"
            If Evento <> "Novo" Then
                 grdAno.TextMatrix(grdAno.Row, cmbMoeda.ItemData(cmbMoeda.ListIndex)) = FormatNumber(txtValorM2.Text, 2)
            End If
      Case "FCAT"
            If Evento <> "Novo" Then
                 nLin = grdAno.Row
                 nCol = cmbMoeda3.ItemData(cmbMoeda3.ListIndex)
                 grdAno.TextMatrix(nLin, nCol) = FormatNumber(txtFatorCateg.Text, 2)
                 grdAno.ColSel = grdAno.Cols - 1
                cmbTipo.Enabled = True
                cmbUso.Enabled = True
                cmbAno2.Enabled = True
            End If
End Select

Evento = ""
      
Exit Sub

fim:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next
Resume Next
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Enabled = False
   cmdAlterar.Visible = True
   cmdExcluir.Enabled = False
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmbTipo.Enabled = True
   cmbUso.Enabled = True
   cmbAno2.Enabled = True
   For Each Ct In frmParam2
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdMain.Enabled = True
   grdProf.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Enabled = False
   cmdAlterar.Visible = False
   cmdExcluir.Enabled = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmParam2
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   txtCateg.BackColor = Kde
   txtCateg.Enabled = False
   txtCod.BackColor = Kde
   txtCod.Locked = True
   txtDesc.BackColor = Kde
'   txtDesc.Locked = True
   txtSeq.BackColor = Kde
   txtSeq.Locked = True
   grdMain.Enabled = False
   grdProf.Enabled = False
   If txtFator.Visible = True Then
        txtFator.SetFocus
   ElseIf txtFator2.Visible = True Then
        txtFator2.SetFocus
   ElseIf txtValorM2.Visible = True Then
        txtValorM2.SetFocus
   Else
       txtFatorCateg.SetFocus
   End If
End If

'FormHagana sParamForm

End Sub

Private Sub Le()

Select Case sParamForm
      Case "DIST", "SITU", "TOPO", "PEDO"
            If grdMain.Row = 0 Then Exit Sub
            txtCod.Text = grdMain.TextMatrix(grdMain.Row, 0)
            txtDesc.Text = grdMain.TextMatrix(grdMain.Row, 1)
            txtFator.Text = grdMain.TextMatrix(grdMain.Row, 2)
      Case "FGLE", "FPRO"
            If grdProf.Row > 0 Then Exit Sub
            txtSeq.Text = grdProf.TextMatrix(grdProf.Row, 0)
            txtMin.Text = grdProf.TextMatrix(grdProf.Row, 1)
            txtMax.Text = grdProf.TextMatrix(grdProf.Row, 2)
            txtFator2.Text = grdProf.TextMatrix(grdProf.Row, 3)
      Case "PGEN"
            cmbAgrup.ListIndex = grdAno.Row - 1
            cmbMoeda.ListIndex = 0
            Sql = "SELECT CODMOEDA,VALORTERRENO FROM TERRENO WHERE CODAGRUPAMENTO=" & cmbAgrup.Text & " AND ANOFATOR=" & cmbAno.Text & " AND CODMOEDA=" & cmbMoeda.ItemData(cmbMoeda.ListIndex)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
            If RdoAux.RowCount > 0 Then
                txtValorM2.Text = FormatNumber(RdoAux!VALORTERRENO, 2)
            End If
      Case "FCAT"
            txtCateg.Text = grdAno.TextMatrix(grdAno.Row, 0)
            cmbMoeda3.ListIndex = 0
            Sql = "Select CODCATEG,CODMOEDA,FATORCATEG FROM FATORCATEG WHERE CODUSO=" & cmbUso.ItemData(cmbUso.ListIndex)
            Sql = Sql & " AND CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND ANOCATEG=" & cmbAno2.Text & " AND "
            Sql = Sql & "CODCATEG=" & grdAno.Row
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
            If RdoAux.RowCount > 0 Then
                txtFatorCateg.Text = FormatNumber(RdoAux!FATORCATEG, 2)
            Else
               txtFatorCateg.Text = 0
            End If
End Select


End Sub

Private Sub Limpa()

txtCod.Text = ""
txtDesc.Text = ""
txtFator.Text = ""
txtSeq.Text = ""
txtMin.Text = ""
txtMax.Text = ""
txtFator2.Text = ""
txtValorM2.Text = ""

End Sub

Private Sub CarregaLista()
Dim nFator As Double

Select Case sParamForm
     Case "PEDO"
           Sql = "Select CODPEDOLOGIA,DESCPEDOLOGIA From PEDOLOGIA WHERE "
           Sql = Sql & "CODPEDOLOGIA<>999 ORDER BY DESCPEDOLOGIA"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdMain.Rows = 1
           With RdoAux
                Do Until .EOF
                      Sql = "SELECT FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE CODPEDOLOGIA=" & !CODPEDOLOGIA
                      Sql = Sql & " AND ANOPEDOLOGIA=" & cmbAno.Text
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      If RdoAux2.RowCount > 0 Then
                           nFator = RdoAux2!FATORPEDOLOGIA
                      Else
                          nFator = 0
                      End If
                      grdMain.AddItem !CODPEDOLOGIA & Chr(9) & !DescPedologia & Chr(9) & FormatNumber(nFator, 4)
                      RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "SITU"
           Sql = "Select CODSITUACAO,DESCSITUACAO From SITUACAO WHERE "
           Sql = Sql & "CODSITUACAO<>999 ORDER BY DESCSITUACAO"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdMain.Rows = 1
           With RdoAux
                Do Until .EOF
                      Sql = "SELECT FATORSITUACAO FROM FATORSITUACAO WHERE CODSITUACAO=" & !CODSITUACAO
                      Sql = Sql & " AND ANOSITUACAO=" & cmbAno.Text
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      If RdoAux2.RowCount > 0 Then
                           nFator = RdoAux2!FATORSITUACAO
                      Else
                          nFator = 0
                      End If
                      grdMain.AddItem !CODSITUACAO & Chr(9) & !DescSituacao & Chr(9) & FormatNumber(nFator, 4)
                      RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "TOPO"
           Sql = "Select CODTOPOGRAFIA,DESCTOPOGRAFIA From TOPOGRAFIA WHERE "
           Sql = Sql & "CODTOPOGRAFIA<>999 ORDER BY DESCTOPOGRAFIA"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdMain.Rows = 1
           With RdoAux
                Do Until .EOF
                      Sql = "SELECT FATORTOPOG FROM FATORTOPOGRAFIA WHERE CODTOPOG=" & !CODTOPOGRAFIA
                      Sql = Sql & " AND ANOTOPOG=" & cmbAno.Text
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      If RdoAux2.RowCount > 0 Then
                           nFator = RdoAux2!FATORTOPOG
                      Else
                          nFator = 0
                      End If
                      grdMain.AddItem !CODTOPOGRAFIA & Chr(9) & !DescTopografia & Chr(9) & FormatNumber(nFator, 4)
                      RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "DIST"
           Sql = "Select CODDISTRITO,DESCDISTRITO From DISTRITO WHERE "
           Sql = Sql & "CODDISTRITO<>999 ORDER BY DESCDISTRITO"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdMain.Rows = 1
           With RdoAux
                Do Until .EOF
                      Sql = "SELECT FATORDISTRITO FROM FATORDISTRITO WHERE CODDISTRITO=" & !CODDISTRITO
                      Sql = Sql & " AND ANODISTRITO=" & cmbAno.Text
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      If RdoAux2.RowCount > 0 Then
                           nFator = RdoAux2!FATORDISTRITO
                      Else
                          nFator = 0
                      End If
                      grdMain.AddItem !CODDISTRITO & Chr(9) & !DescDistrito & Chr(9) & FormatNumber(nFator, 4)
                      RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "FBEN"
           Sql = "Select CODBENFEITORIA,DESCBENFEITORIA From BENFEITORIA WHERE "
           Sql = Sql & "CODBENFEITORIA<>999 ORDER BY DESCBENFEITORIA"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdMain.Rows = 1
           With RdoAux
                Do Until .EOF
                    If !CODBENFEITORIA = 999 Then
                        Sql = "SELECT ANOFATORBENF,FATORBENFEITORIA FROM FATORBENFEITORIA "
                        Sql = Sql & "WHERE ANOFATORBENF=" & cmbAno.Text & " AND CODBENFEITORIA=" & .rdoColumns(0)
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                         If RdoAux2.RowCount > 0 Then
                               grdMain.AddItem "0 (Pred.)" & Chr(9) & "Terreno + Edificação" & Chr(9) & FormatNumber(RdoAux2!FATORBENFEITORIA, 4)
                         Else
                              grdMain.AddItem "0 (Pred.)" & Chr(9) & "Terreno + Edificação" & Chr(9) & "0,0000"
                         End If
                     Else
                        Sql = "SELECT ANOFATORBENF,FATORBENFEITORIA FROM FATORBENFEITORIA "
                        Sql = Sql & "WHERE ANOFATORBENF=" & cmbAno.Text & " AND CODBENFEITORIA=" & !CODBENFEITORIA
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                         If RdoAux2.RowCount > 0 Then
                               grdMain.AddItem .rdoColumns(0) & " (Terr.)" & Chr(9) & .rdoColumns(1) & Chr(9) & FormatNumber(RdoAux2!FATORBENFEITORIA, 4)
                         Else
                              grdMain.AddItem .rdoColumns(0) & " (Terr.)" & Chr(9) & .rdoColumns(1) & Chr(9) & "0,0000"
                         End If
                     End If
                     RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "FPRO"
           Sql = "Select CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE WHERE  CODDISTRITO=" & cmbDist.ItemData(cmbDist.ListIndex) & " ORDER BY MINPROFUN"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdProf.Rows = 1
           With RdoAux
                Do Until .EOF
                     Sql = "SELECT ANOPROFUN,FATORPROFUN FROM FATORPROFUN "
                     Sql = Sql & "WHERE ANOPROFUN=" & cmbAno.Text & " AND CODDISTRITO=" & cmbDist.ItemData(cmbDist.ListIndex) & " AND CODPROFUN=" & .rdoColumns(1)
                     Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                     If RdoAux2.RowCount > 0 Then
                          grdProf.AddItem Format(.rdoColumns(1), "00") & Chr(9) & FormatNumber(.rdoColumns(2), 2) & Chr(9) & FormatNumber(.rdoColumns(3), 2) & Chr(9) & FormatNumber(RdoAux2!FATORPROFUN, 4)
                     Else
                          grdProf.AddItem Format(.rdoColumns(1), "00") & Chr(9) & FormatNumber(.rdoColumns(2), 2) & Chr(9) & FormatNumber(.rdoColumns(3), 2) & Chr(9) & "0,0000"
                     End If
                     RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "FGLE"
           Sql = "Select CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY MINGLEBA"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           grdProf.Rows = 1
           With RdoAux
                Do Until .EOF
                       Sql = "SELECT ANOGLEBA,FATORGLEBA FROM FATORGLEBA "
                       Sql = Sql & "WHERE ANOGLEBA=" & cmbAno.Text & " AND CODGLEBA=" & .rdoColumns(0)
                       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                       If RdoAux2.RowCount > 0 Then
                            grdProf.AddItem Format(.rdoColumns(0), "00") & Chr(9) & FormatNumber(.rdoColumns(1), 2) & Chr(9) & FormatNumber(.rdoColumns(2), 2) & Chr(9) & FormatNumber(RdoAux2!FATORGLEBA, 4)
                       Else
                            grdProf.AddItem Format(.rdoColumns(0), "00") & Chr(9) & FormatNumber(.rdoColumns(1), 2) & Chr(9) & FormatNumber(.rdoColumns(2), 2) & Chr(9) & "0,00"
                       End If
                       RdoAux2.Close
                    .MoveNext
                Loop
               .Close
            End With
     Case "PGEN"
            Sql = "SELECT CODMOEDA,DESCMOEDA  FROM MOEDA"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
            grdAno.Cols = RdoAux.RowCount + 1
            For x = 0 To grdAno.Cols - 1
                  grdAno.ColWidth(x) = 900
                  grdAno.ColAlignment(x) = flexAlignCenterCenter
            Next
            grdAno.TextMatrix(0, 0) = "Agrup."
            x = 1
            cmbMoeda.Clear
            Do Until RdoAux.EOF
                 grdAno.TextMatrix(0, x) = RdoAux!DESCMOEDA
                 cmbMoeda.AddItem RdoAux!DESCMOEDA
                 cmbMoeda.ItemData(cmbMoeda.NewIndex) = RdoAux!CODMOEDA
                 RdoAux.MoveNext
                 x = x + 1
            Loop
            grdAno.Rows = 7
            For x = 1 To 7
                cmbAgrup.Clear
                grdAno.TextMatrix(x, 0) = x
                cmbAgrup.AddItem x
            Next
            Sql = "Select CODAGRUPAMENTO,ANOFATOR,CODMOEDA,VALORTERRENO FROM TERRENO WHERE ANOFATOR=" & cmbAno.Text
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdProf.Rows = 1
            With RdoAux
                For x = 1 To grdAno.Rows - 1
                      For Y = 1 To grdAno.Cols - 1
                             grdAno.TextMatrix(x, Y) = "0,00"
                      Next Y
                Next x
               Do Until .EOF
                     nLin = RdoAux!CODAGRUPAMENTO
                     nCol = RdoAux!CODMOEDA
                     grdAno.TextMatrix(nLin, nCol) = Format(RdoAux!VALORTERRENO, "#0.00")
                   .MoveNext
               Loop
              .Close
            End With
            If grdAno.Rows > 1 Then
                 grdAno.Row = 1
                 grdAno.ColSel = grdAno.Cols - 1
                 Le
           End If
      Case "FCAT"
         Sql = "SELECT CODCATEGCONSTR,DESCCATEGCONSTR  FROM CATEGCONSTR WHERE CODCATEGCONSTR<>999"
         Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
         grdAno.Rows = RdoAux.RowCount + 1
         x = 1
         With RdoAux
                Do Until .EOF
                    grdAno.TextMatrix(x, 0) = !DESCCATEGCONSTR
                    x = x + 1
                   .MoveNext
                Loop
         End With
         Sql = "SELECT CODMOEDA,DESCMOEDA  FROM MOEDA"
         Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
         grdAno.Cols = RdoAux.RowCount + 1
         For x = 1 To grdAno.Cols - 1
               grdAno.ColWidth(x) = 900
               grdAno.ColAlignment(x) = flexAlignCenterCenter
         Next
         grdAno.ColWidth(0) = 1600
         grdAno.ColAlignment(0) = flexAlignLeftCenter
         grdAno.TextMatrix(0, 0) = "Categoria"
         cmbMoeda3.Clear
         x = 1
         Do Until RdoAux.EOF
               grdAno.TextMatrix(0, x) = RdoAux!DESCMOEDA
              cmbMoeda3.AddItem RdoAux!DESCMOEDA
              cmbMoeda3.ItemData(cmbMoeda3.NewIndex) = RdoAux!CODMOEDA
              x = x + 1
              RdoAux.MoveNext
         Loop
         cmbMoeda3.ListIndex = 0
         Sql = "Select CODCATEG,CODMOEDA,FATORCATEG FROM FATORCATEG WHERE CODUSO=" & cmbUso.ItemData(cmbUso.ListIndex)
         Sql = Sql & " AND CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND ANOCATEG=" & cmbAno2.Text
         Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux
             For x = 1 To grdAno.Rows - 1
                   For Y = 1 To grdAno.Cols - 1
                          grdAno.TextMatrix(x, Y) = "0,00"
                   Next Y
             Next x
            Do Until .EOF
                  nLin = RdoAux!CODCATEG
                  nCol = RdoAux!CODMOEDA
                  grdAno.TextMatrix(nLin, nCol) = Format(RdoAux!FATORCATEG, "#0.00")
                .MoveNext
            Loop
           .Close
        End With
        If grdAno.Rows > 1 Then
           grdAno.Row = 1
           grdAno.ColSel = grdAno.Cols - 1
           Le
        End If
End Select
End Sub

Private Sub FormHagana(sTela As String)

If NomeDeLogin = "SCHWARTZ" Then Exit Sub
'Exit Sub
evNewPE = 21: evNewSI = 25: evNewTO = 33: evNewDI = 49: evNewFB = 65: evNewPR = 53: evNewGL = 61: evNewPG = 57: evNewFC = 73
evEditPE = 22: evEditSI = 26: evEditTO = 34: evEditDI = 50: evEditFB = 66: evEditPr = 54: evEditGL = 62: evEditPG = 58: evEditFC = 74
evDelPE = 23: evDelSI = 27: evDelTO = 35: evDelDI = 51: evDelFB = 67: evDelPR = 55: evDelGL = 63: evDelPG = 59: evDelFC = 75

bEvNewPE = False: bEvNewSI = False: bEvNewTO = False: bEvNewDI = False: bEvNewFB = False: bEvNewPR = False: bEvNewGL = False: bEvNewPG = False: bEvNewFC = False
bEvEditPE = False: bEvEditSI = False: bEvEditTO = False: bEvEditDI = False: bEvEditFB = False: bEvEditPR = False: bEvEditGL = False: bEvEditPG = False: bEvEditFC = False
bEvDelPE = False: bEvDelSI = False: bEvDelTO = False: bEvDelDI = False: bEvDelFB = False: bEvDelPR = False: bEvDelGL = False: bEvDelPG = False: bEvDelFC = False


If InStr(1, sRet, Format(evEditPE, "000"), vbBinaryCompare) > 0 Then bEvEditPE = True
If InStr(1, sRet, Format(evEditSI, "000"), vbBinaryCompare) > 0 Then bEvEditSI = True
If InStr(1, sRet, Format(evEditTO, "000"), vbBinaryCompare) > 0 Then bEvEditTO = True
If InStr(1, sRet, Format(evEditDI, "000"), vbBinaryCompare) > 0 Then bEvEditDI = True
If InStr(1, sRet, Format(evEditFB, "000"), vbBinaryCompare) > 0 Then bEvEditFB = True
If InStr(1, sRet, Format(evEditPr, "000"), vbBinaryCompare) > 0 Then bEvEditPR = True
If InStr(1, sRet, Format(evEditGL, "000"), vbBinaryCompare) > 0 Then bEvEditGL = True
If InStr(1, sRet, Format(evEditPG, "000"), vbBinaryCompare) > 0 Then bEvEditPG = True
If InStr(1, sRet, Format(evEditFC, "000"), vbBinaryCompare) > 0 Then bEvEditFC = True

Select Case sTela
          Case "PEDO"
                cmdNovo.Enabled = bEvNewPE
                cmdAlterar.Enabled = bEvEditPE
                cmdExcluir.Enabled = bEvDelPE
          Case "SITU"
                cmdNovo.Enabled = bEvNewSI
                cmdAlterar.Enabled = bEvEditSI
                cmdExcluir.Enabled = bEvDelSI
          Case "TOPO"
                cmdNovo.Enabled = bEvNewTO
                cmdAlterar.Enabled = bEvEditTO
                cmdExcluir.Enabled = bEvDelTO
          Case "DIST"
                cmdNovo.Enabled = bEvNewDI
                cmdAlterar.Enabled = bEvEditDI
                cmdExcluir.Enabled = bEvDelDI
          Case "FBEN"
                cmdNovo.Enabled = bEvNewFB
                cmdAlterar.Enabled = bEvEditFB
                cmdExcluir.Enabled = bEvDelFB
          Case "FPRO"
                cmdNovo.Enabled = bEvNewPR
                cmdAlterar.Enabled = bEvEditPR
                cmdExcluir.Enabled = bEvDelPR
          Case "FGLE"
                cmdNovo.Enabled = bEvNewGL
                cmdAlterar.Enabled = bEvEditGL
                cmdExcluir.Enabled = bEvDelGL
          Case "PGEN"
                cmdNovo.Enabled = bEvNewPG
                cmdAlterar.Enabled = bEvEditPG
                cmdExcluir.Enabled = bEvDelPG
          Case "FCAT"
'                cmdNovo.Enabled = bEvNewFC
'                cmdAlterar.Enabled = bEvEditFC
'                cmdExcluir.Enabled = bEvDelFC
End Select

End Sub

Private Sub txtFator_KeyPress(KeyAscii As Integer)
'Tweak txtFator, KeyAscii, DecimalPositive
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If
If KeyAscii = 44 And InStr(1, txtFator.Text, ",", vbTextCompare) <> 0 Then
   KeyAscii = 0
End If
End Sub

Private Sub txtFator2_KeyPress(KeyAscii As Integer)
'Tweak txtFator2, KeyAscii, DecimalPositive
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If
If KeyAscii = 44 And InStr(1, txtFator2.Text, ",", vbTextCompare) <> 0 Then
   KeyAscii = 0
End If
End Sub

Private Sub txtFatorCateg_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If
If KeyAscii = 44 And InStr(1, txtFatorCateg.Text, ",", vbTextCompare) <> 0 Then
   KeyAscii = 0
End If

End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If
If KeyAscii = 44 And InStr(1, txtMax.Text, ",", vbTextCompare) <> 0 Then
   KeyAscii = 0
End If
End Sub

Private Sub txtMin_KeyPress(KeyAscii As Integer)

If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If
If KeyAscii = 44 And InStr(1, txtMin.Text, ",", vbTextCompare) <> 0 Then
   KeyAscii = 0
End If

End Sub

