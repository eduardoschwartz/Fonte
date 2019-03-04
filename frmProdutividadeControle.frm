VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProdutividadeControle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtividade - Controle diário de tarefas"
   ClientHeight    =   6045
   ClientLeft      =   5895
   ClientTop       =   3120
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   7410
   Begin Tributacao.jcFrames pnlObs 
      Height          =   3705
      Left            =   495
      Top             =   1035
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
         TabIndex        =   30
         Top             =   450
         Width           =   6495
      End
      Begin prjChameleon.chameleonButton cmdGravarObs 
         Height          =   315
         Left            =   5400
         TabIndex        =   31
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
         MICON           =   "frmProdutividadeControle.frx":0000
         PICN            =   "frmProdutividadeControle.frx":001C
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
   Begin prjChameleon.chameleonButton cmdData 
      Height          =   345
      Left            =   2565
      TabIndex        =   29
      ToolTipText     =   "Alterar a data da tarefa"
      Top             =   5625
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   14
      TX              =   "&Data"
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
      MICON           =   "frmProdutividadeControle.frx":03C1
      PICN            =   "frmProdutividadeControle.frx":03DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5940
      MaxLength       =   15
      TabIndex        =   28
      Top             =   1170
      Width           =   1410
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   585
      Locked          =   -1  'True
      MaxLength       =   1000
      TabIndex        =   26
      Top             =   5220
      Width           =   6720
   End
   Begin prjChameleon.chameleonButton cmbObs 
      Height          =   345
      Left            =   1305
      TabIndex        =   25
      ToolTipText     =   "Observação do item selecionado"
      Top             =   5625
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
      MICON           =   "frmProdutividadeControle.frx":07BC
      PICN            =   "frmProdutividadeControle.frx":07D8
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
      Height          =   345
      Left            =   5130
      TabIndex        =   6
      ToolTipText     =   "Editar Registro"
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmProdutividadeControle.frx":0854
      PICN            =   "frmProdutividadeControle.frx":0870
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
      Height          =   345
      Left            =   6270
      TabIndex        =   7
      ToolTipText     =   "Excluir Registro"
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmProdutividadeControle.frx":09CA
      PICN            =   "frmProdutividadeControle.frx":09E6
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
      Height          =   345
      Left            =   5130
      TabIndex        =   20
      ToolTipText     =   "Gravar os Dados"
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmProdutividadeControle.frx":0A88
      PICN            =   "frmProdutividadeControle.frx":0AA4
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
      Height          =   345
      Left            =   6270
      TabIndex        =   21
      ToolTipText     =   "Cancelar Edição"
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmProdutividadeControle.frx":0E49
      PICN            =   "frmProdutividadeControle.frx":0E65
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
      Height          =   345
      Left            =   3990
      TabIndex        =   5
      ToolTipText     =   "Novo Registro"
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmProdutividadeControle.frx":0FBF
      PICN            =   "frmProdutividadeControle.frx":0FDB
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
      Left            =   90
      TabIndex        =   8
      ToolTipText     =   "Alterar a quantidade do item selecionado"
      Top             =   5625
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
      MICON           =   "frmProdutividadeControle.frx":1135
      PICN            =   "frmProdutividadeControle.frx":1151
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7395
      Begin VB.ComboBox cmbFiscal 
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   180
         Width           =   4305
      End
      Begin VB.TextBox txtEvento 
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "NORMAL"
         Top             =   600
         Width           =   1845
      End
      Begin VB.TextBox txtTarefas 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPontos 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6390
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   630
         Width           =   855
      End
      Begin MSComCtl2.DTPicker mskDataIni 
         Height          =   315
         Left            =   5940
         TabIndex        =   1
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         _Version        =   393216
         Format          =   172621825
         CurrentDate     =   40968
         MaxDate         =   43830
         MinDate         =   40544
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiscal....:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   240
         Width           =   705
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data..:"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   5400
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Evento..:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   660
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Tarefas..:"
         Height          =   225
         Index           =   2
         Left            =   2850
         TabIndex        =   14
         Top             =   660
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pontos no dia..:"
         Height          =   225
         Index           =   3
         Left            =   5100
         TabIndex        =   13
         Top             =   660
         Width           =   1215
      End
   End
   Begin prjChameleon.chameleonButton cmdT1 
      Height          =   285
      Left            =   90
      TabIndex        =   2
      ToolTipText     =   "Remove centro de custos"
      Top             =   1140
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   503
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
      MICON           =   "frmProdutividadeControle.frx":1230
      PICN            =   "frmProdutividadeControle.frx":124C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdT2 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      ToolTipText     =   "Adiciona centro de custos"
      Top             =   1140
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   503
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
      MICON           =   "frmProdutividadeControle.frx":13A6
      PICN            =   "frmProdutividadeControle.frx":13C2
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
      Height          =   3615
      Left            =   30
      TabIndex        =   4
      Top             =   1530
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6376
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Obs.:"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   90
      TabIndex        =   27
      Top             =   5265
      Width           =   390
   End
   Begin VB.Label lblSeq 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3240
      TabIndex        =   24
      Top             =   5685
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos na Tarefa..:"
      Height          =   225
      Index           =   5
      Left            =   2670
      TabIndex        =   23
      Top             =   1200
      Width           =   1365
   End
   Begin VB.Label lblPontoTarefa 
      BackStyle       =   0  'Transparent
      Caption         =   "000"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4170
      TabIndex        =   22
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Processo nº..:"
      Height          =   225
      Index           =   4
      Left            =   4875
      TabIndex        =   19
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label lblTarefa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tarefa 0 de 0"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   540
      TabIndex        =   18
      Top             =   1200
      Width           =   1245
   End
End
Attribute VB_Name = "frmProdutividadeControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type aDiaType
    nSeq As Integer
    sNumProc As String
    nAnoproc As Integer
    nNumproc As Long
End Type

Private Type aEventoType
    sNome As String
    nPontos As Double
End Type

Dim nPointer As Integer, nMaxPointer As Integer, bNovo As Boolean, aDia() As aDiaType, aEvento() As aEventoType
Dim bIsBoss As Boolean

Private Sub cmbFiscal_Click()
If cmbFiscal.ListIndex = -1 Then Exit Sub
bIsBoss = ProdIsBoss(cmbFiscal.ItemData(cmbFiscal.ListIndex))
CarregaDia
End Sub

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
cmdAlterar.Enabled = bEnable
cmdCancel.Enabled = bEnable
cmdData.Enabled = bEnable
cmdExcluir.Enabled = bEnable
cmdGravar.Enabled = bEnable
cmdNovo.Enabled = bEnable
cmdQtde.Enabled = bEnable
cmbObs.Enabled = bEnable
lvMain.Enabled = bEnable

End Sub


Private Sub cmdAlterar_Click()
Dim dData As Date, nMes As Integer, nAno As Integer
Dim Sql As String, RdoAux As rdoResultset

If cmbFiscal.ListIndex = -1 Then
    MsgBox "Selecione um fiscal.", vbExclamation, "Atenção"
    Exit Sub
End If
If nPointer = 0 Then
    MsgBox "Selecione a tarefa a ser alterada.", vbExclamation, "Atenção"
    Exit Sub
End If

dData = mskDataIni.value

If dData < CDate("01/04/2012") Then
    MsgBox "Não é possível alterar tarefas deste período.", vbCritical, "Atenção"
    Exit Sub
End If

nMes = Month(dData)
nAno = Year(dData)

Sql = "SELECT * FROM PRODUTIVIDADEENCERRA WHERE ANO=" & nAno & " AND MES=" & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este mês já foi encerrado e não pode mais ser alterado.", vbCritical, "Atenção"
    RdoAux.Close
    Exit Sub
End If
RdoAux.Close


bNovo = False
ControlBehaviour False
End Sub

Private Sub cmdCancel_Click()
ControlBehaviour True
mskDataIni_Change
End Sub

Private Sub cmdData_Click()
Dim n As Variant, sOld As String, nMes As Integer, nAno As Integer, nMesOld As Integer, nAnoOld As Integer, nMesAtual As Integer, nAnoAtual As Integer
Dim bValido As Boolean

bValido = False
If nPointer = 0 Then
    MsgBox "Selecione a tarefa a ser alterada.", vbExclamation, "Atenção"
Else
    sOld = Format(mskDataIni.value, "dd/mm/yyyy")
    n = InputBox("Digite a nova data para a tarefa", "Nova data", sOld)
    If Not IsDate(n) Then
        MsgBox "Data inválida", vbCritical, "Atenção"
    Else
        n = Format(n, "dd/mm/yyyy")
        nMesAtual = Month(Now)
        nAnoAtual = Year(Now)
        nMes = Month(n)
        nAno = Year(n)
        If nMes = 1 Then
            nMesOld = 12
            nAnoOld = nAno - 1
        Else
            nMesOld = nMes - 1
            nAnoOld = nAno
        End If
        
        If nMes = nMesAtual And nAno = nAnoAtual Then
            bValido = True
        Else
            If nMesAtual = 1 And nMesOld = 12 And nAnoOld = nAnoAtual - 1 Then
                bValido = True
            Else
                If nMesAtual > 1 And (nMesAtual - nMes) = 1 And nAno = nAnoAtual Then
                    bValido = True
                End If
            End If
            
        End If
        
        If bValido Then
            Sql = "update produtividadetarefa set data='" & Format(n, "mm/dd/yyyy") & "' where fiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex) & " and "
            Sql = Sql & "data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' and seq=" & Val(lblSeq.Caption)
            cn.Execute Sql, rdExecDirect
            mskDataIni.value = Format(n, "dd/mm/yyyy")
            mskDataIni_Change
            ControlBehaviour True
        Else
            MsgBox "A Data não pode ser alterada para a data digitada.", vbCritical, "Atenção"
        End If
    End If
End If

End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String

If cmbFiscal.ListIndex = -1 Then
    MsgBox "Selecione um fiscal.", vbExclamation, "Atenção"
    Exit Sub
End If
If nPointer = 0 Then
    MsgBox "Selecione a tarefa a ser excluida.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Deseja excluir esta tarefa?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "delete from produtividadetarefa where fiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex)
Sql = Sql & " and data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' and seq=" & aDia(nPointer).nSeq
cn.Execute Sql, rdExecDirect

CarregaDia

End Sub

Private Sub cmdGravar_Click()
If Val(lblPontoTarefa.Caption) = 0 Then
    MsgBox "Nenhum item selecionado.", vbCritical, "Atenção"
    Exit Sub
End If

If Not Valida() Then Exit Sub
Grava
ControlBehaviour True
CarregaDia

End Sub

Private Function Valida() As Boolean
Dim x As Integer, bApuracao As Boolean, bOutros As Boolean, sItem As String, sNumProcesso As String, bInterno As Boolean

'bApuracao = False
'bOutros = False
'bInterno = False
'
'For x = 1 To lvMain.ListItems.Count
'    If lvMain.ListItems(x).Checked Then
'        sItem = lvMain.ListItems(x).Text
'        If sItem = "15" Or sItem = "15.1" Or sItem = "16.1" Then
'            bApuracao = True
'        End If
'        If sItem <> "15" And sItem <> "15.1" And sItem <> "16.1" And sItem <> "19" Then
'            bOutros = True
'        End If
'        If sItem = "19" Then
'            bInterno = True
'        End If
'    End If
'Next
'
'If bApuracao And bOutros And Not ProdIsBossLogin Then
'    MsgBox "Apuração fiscal não pode conter outros ítens.", vbCritical, "Erro"
'    Valida = False
'    Exit Function
'End If

If Len(sNumProcesso) < 5 Then
    If bInterno And (bOutros Or bApuracao) Then
        MsgBox "Item 19 sem processo não pode conter outros ítens.", vbCritical, "Erro"
        Valida = False
        Exit Function
    Else
        Valida = True
        Exit Function
    End If
End If

sNumProcesso = txtNumProc.Text
sNumProcesso = Replace(sNumProcesso, "-", "")
If Len(sNumProcesso) < 5 Then
    sNumProcesso = "000000-0/0000"
End If
nNumproc = Val(Left$(sNumProcesso, InStr(1, sNumProcesso, "/", vbBinaryCompare) - 1))
nAnoproc = Val(Right$(sNumProcesso, 4))

If Not ProdIsBossLogin Then
    If Right$(nNumproc, 1) <> RetornaDVProcesso(Val(Left$(sNumProcesso, InStr(1, sNumProcesso, "/", vbBinaryCompare) - 2))) Then
        MsgBox "Número de Processo inválido", vbExclamation, "Atenção"
        Valida = False
        Exit Function
    End If
End If


Valida = True
End Function

Private Sub cmdGravarObs_Click()
pnlObs.Visible = False
txtObs.Text = txtObsTmp.Text
lvMain.ListItems(lvMain.SelectedItem.Index).SubItems(4) = txtObs.Text
EnableObs (True)

End Sub

Private Sub cmdNovo_Click()

If cmbFiscal.ListIndex = -1 Then
    MsgBox "Selecione um fiscal.", vbExclamation, "Atenção"
    Exit Sub
End If
dData = mskDataIni.value
nMes = Month(dData)
nAno = Year(dData)

Sql = "SELECT * FROM PRODUTIVIDADEENCERRA WHERE ANO=" & nAno & " AND MES=" & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este mês já foi encerrado e não pode mais ser alterado.", vbCritical, "Atenção"
    RdoAux.Close
    Exit Sub
End If
RdoAux.Close


bNovo = True
lblPontoTarefa.Caption = "0,00"
txtNumProc.Text = ""
ControlBehaviour False
ClearGrid
End Sub

Private Sub cmdQtde_Click()
Dim n As Variant, nOld As Double
If lvMain.SelectedItem.Checked = False Then
    MsgBox "Marque primeiro o ítem que deseja alterar.", vbExclamation, "Atenção"
Else
    nOld = Val(lvMain.SelectedItem.SubItems(2))
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
    TotalizaTarefa
End If

End Sub

Private Sub cmdT1_Click()
nPointer = nPointer - 1
cmdT2.Enabled = True
If nPointer = 1 Then
    cmdT1.Enabled = False
End If
lblTarefa.Caption = nPointer & " de " & nMaxPointer

CarregaTarefa

End Sub

Private Sub cmdT2_Click()
nPointer = nPointer + 1
cmdT1.Enabled = True
If nPointer = nMaxPointer Then
    cmdT2.Enabled = False
End If
lblTarefa.Caption = nPointer & " de " & nMaxPointer

CarregaTarefa
End Sub

Private Sub Form_Load()
Dim RdoAux As rdoResultset, Sql As String
On Error GoTo Erro
bNovo = False
Centraliza Me
ControlBehaviour True
CarregaEvento


Sql = "select codigo,nome,nomecompleto from produtividadefiscal inner join "
Sql = Sql & "usuario on produtividadefiscal.nome = usuario.nomelogin order by nomecompleto "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFiscal.AddItem !NomeCompleto
        cmbFiscal.ItemData(cmbFiscal.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "NOELI" Then
    If Not ProdIsBossLogin() Then
        cmbFiscal.Text = RetornaUsuarioFullName()
        cmbFiscal_Click
        cmbFiscal.Enabled = False
    End If
End If

cmdT1.Enabled = False
cmdT2.Enabled = False
mskDataIni.value = Now
CarregaLista
CarregaDia

Exit Sub
Erro:
cmbFiscal.Enabled = False
mskDataIni.Enabled = False
MsgBox "Erro Fatal.", vbCritical, "Atenção"

End Sub

Private Sub ControlBehaviour(bStart As Boolean)
cmdNovo.Visible = bStart
cmdAlterar.Visible = bStart
cmdExcluir.Visible = bStart
cmdGravar.Visible = Not bStart
cmdCancel.Visible = Not bStart
cmdQtde.Enabled = Not bStart
cmbObs.Enabled = Not bStart
cmdData.Enabled = Not bStart
cmdAlterar.Enabled = Not bStart
cmbFiscal.Enabled = bStart
mskDataIni.Enabled = bStart

If bStart Then
    cmbFiscal.BackColor = Branco
    mskDataIni.CalendarTitleBackColor = Me.BackColor
    txtNumProc.BackColor = Me.BackColor
    txtNumProc.Locked = True
Else
    cmbFiscal.BackColor = Me.BackColor
    mskDataIni.CalendarTitleBackColor = Branco
    txtNumProc.BackColor = Branco
    txtNumProc.Locked = False
End If

End Sub

Private Sub CarregaLista()
Dim RdoAux As rdoResultset, Sql As String, itmX As ListItem
lvMain.ListItems.Clear
Sql = "select * from vwprodutividade where dataini<='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' and "
Sql = Sql & " datafim>='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' "
'Sql = Sql & " and item not in ('14','22')"
Sql = Sql & " order by item"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, , !Item)
        itmX.SubItems(1) = !descricao
        itmX.SubItems(2) = "1,00"
        itmX.SubItems(3) = FormatNumber(!Valor, 2)
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lvMain_Click()
txtObs.Text = lvMain.SelectedItem.SubItems(4)
End Sub

Private Sub lvMain_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Item.Selected = True
If cmdNovo.Visible Then
    Item.Checked = Not Item.Checked
End If

'If Item.Text = "13" And Not ProdIsBossLogin() Then
'    Item.Checked = Not Item.Checked
'End If

TotalizaTarefa
End Sub

Private Sub CarregaDia()
Dim RdoAux As rdoResultset, Sql As String, itmX As ListItem, nCodEvento As Integer, nCodFiscal As Integer

If cmbFiscal.ListIndex = -1 Then Exit Sub
Limpa
ClearGrid

If bIsBoss Then GoTo Chefe
cmdNovo.Enabled = True
cmdAlterar.Enabled = True
cmdExcluir.Enabled = True

nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
nCodEvento = ProdEventoDia(nCodFiscal, CDate(mskDataIni.value))
If nCodEvento > 0 Then
    txtEvento.Text = aEvento(nCodEvento).sNome
    GoTo OutroEvento
Else
    txtEvento.Text = "NORMAL"
End If

ReDim aDia(0)

Sql = "select DISTINCT Seq,processo,ano,numero from produtividadetarefa where fiscal=" & nCodFiscal
Sql = Sql & " and data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aDia(UBound(aDia) + 1)
        aDia(UBound(aDia)).nSeq = !Seq
        aDia(UBound(aDia)).sNumProc = !Processo
        aDia(UBound(aDia)).nAnoproc = !Ano
        aDia(UBound(aDia)).nNumproc = !Numero
       .MoveNext
    Loop
   .Close
End With

If UBound(aDia) > 0 Then
    lblTarefa.Caption = "1 de " & UBound(aDia)
    nPointer = 1
    cmdT1.Enabled = False
    nMaxPointer = UBound(aDia)
    txtTarefas.Text = nMaxPointer
    If nPointer < nMaxPointer Then
        cmdT2.Enabled = True
    Else
        cmdT2.Enabled = False
    End If
    CarregaTarefa
Else
    lblTarefa.Caption = "0 de 0"
    nPointer = 0
    cmdT1.Enabled = False
    cmdT2.Enabled = False
End If

TotalizaDia
Exit Sub

Chefe:
txtEvento.Text = "CHEFIA"
txtPontos.Text = 30
lblTarefa.Caption = "0 de 0"
nPointer = 0
cmdT1.Enabled = False
cmdT2.Enabled = False
cmdNovo.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
Exit Sub:

OutroEvento:
txtPontos.Text = aEvento(nCodEvento).nPontos
lblTarefa.Caption = "0 de 0"
nPointer = 0
cmdT1.Enabled = False
cmdT2.Enabled = False
'cmdNovo.Enabled = False
'cmdAlterar.Enabled = False
'cmdExcluir.Enabled = False

End Sub

Private Sub lvMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
txtObs.Text = lvMain.SelectedItem.SubItems(4)
End Sub

Private Sub mskDataIni_Change()
CarregaLista
CarregaDia

End Sub

Private Sub Limpa()
txtNumProc.Text = "000000-0/0000"
txtTarefas.Text = 0
txtPontos.Text = "0,00"
lblPontoTarefa.Caption = "0,00"
End Sub

Private Sub ClearGrid()
Dim x As Integer
For x = 1 To lvMain.ListItems.Count
    lvMain.ListItems(x).Checked = False
    lvMain.ListItems(x).SubItems(2) = 1
Next

End Sub

Private Sub CarregaTarefa()
Dim Sql As String, RdoAux As rdoResultset, x As Integer

ClearGrid
txtNumProc.Text = aDia(nPointer).sNumProc
If txtNumProc.Text = "" Then txtNumProc.Text = "000000-0/0000"
lblSeq.Caption = aDia(nPointer).nSeq
Sql = "select item,qtde,processo,obs from produtividadetarefa where fiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex)
Sql = Sql & " and data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' and seq=" & aDia(nPointer).nSeq

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For x = 1 To lvMain.ListItems.Count
            If lvMain.ListItems(x).Text = !Item Then
                lvMain.ListItems(x).SubItems(2) = FormatNumber(!QTDE, 2)
                lvMain.ListItems(x).Checked = True
                lvMain.ListItems(x).SubItems(4) = SubNull(!obs)
            End If
        Next
       .MoveNext
    Loop
   .Close
End With
TotalizaTarefa

End Sub

Private Sub TotalizaTarefa()
Dim x As Integer, nTotal As Double

nTotal = 0
For x = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(x).Checked Then
        nTotal = nTotal + (CDbl(lvMain.ListItems(x).SubItems(2)) * CDbl(lvMain.ListItems(x).SubItems(3)))
    End If
Next

lblPontoTarefa.Caption = FormatNumber(nTotal, 2)

End Sub

Private Sub TotalizaDia()
Dim Sql As String, RdoAux As rdoResultset

Sql = "select sum(valor * qtde) as soma from produtividadetarefa where fiscal=" & cmbFiscal.ItemData(cmbFiscal.ListIndex)
Sql = Sql & " and data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Not IsNull(RdoAux!soma) Then
    txtPontos.Text = FormatNumber(RdoAux!soma, 2)
End If
RdoAux.Close

End Sub

Private Sub Grava()
Dim sNumProcesso As String, nAno As Integer, nNumproc As Long, sObs As String, nAnoproc As Integer, nQtde As Double

If MsgBox("Deseja gravar as alterações?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
If bNovo Then
    Sql = "SELECT max(seq) as maximo from produtividadetarefa where data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' "
    Sql = Sql & " and fiscal=" & nCodFiscal
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq = 1
    Else
        nSeq = RdoAux!maximo + 1
    End If
    RdoAux.Close
    sNumProcesso = txtNumProc.Text
    If Val(sNumProcesso) > 0 Then
        nNumproc = Val(Left$(sNumProcesso, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProcesso, 4))
    Else
        nNumproc = 0
        nAnoproc = 0
    End If
    ReDim Preserve aDia(UBound(aDia) + 1)
    aDia(UBound(aDia)).nSeq = nSeq
    aDia(UBound(aDia)).sNumProc = sNumProcesso
    aDia(UBound(aDia)).nAnoproc = nAnoproc
    aDia(UBound(aDia)).nNumproc = nNumproc
    nMaxPointer = UBound(aDia)
Else
    nSeq = Val(lblSeq.Caption)
    Sql = "delete from produtividadetarefa where fiscal=" & nCodFiscal & " and data='" & Format(mskDataIni.value, "mm/dd/yyyy") & "' and seq=" & nSeq
    cn.Execute Sql, rdExecDirect
    sNumProcesso = txtNumProc.Text
    nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
    nNumproc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1))
End If


For nRow = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(nRow).Checked Then
        sItem = lvMain.ListItems(nRow).Text
        nQtde = CDbl(lvMain.ListItems(nRow).SubItems(2))
        nValor = CDbl(lvMain.ListItems(nRow).SubItems(3))
        sObs = lvMain.ListItems(nRow).SubItems(4)
        
        Sql = "insert produtividadetarefa(data,fiscal,seq,item,qtde,valor,ano,numero,processo,obs) values('"
        Sql = Sql & Format(mskDataIni.value, "mm/dd/yyyy") & "'," & nCodFiscal & "," & nSeq & ",'" & sItem & "',"
        Sql = Sql & Virg2Ponto(CStr(nQtde)) & "," & Virg2Ponto(CStr(nValor)) & "," & nAno & "," & nNumproc & ",'" & sNumProcesso & "','" & Mask(sObs) & "')"
        cn.Execute Sql, rdExecDirect
    End If
Next

End Sub

Private Sub CarregaEvento()
Dim Sql As String, RdoAux As rdoResultset

ReDim aEvento(0)
Sql = "select codigo,nome,pontodia from produtividadeevento order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aEvento(UBound(aEvento) + 1)
        aEvento(UBound(aEvento)).sNome = !Nome
        aEvento(UBound(aEvento)).nPontos = !pontodia
       .MoveNext
    Loop
   .Close
End With

End Sub
