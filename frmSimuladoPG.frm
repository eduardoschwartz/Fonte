VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSimuladoPG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulador de Planta Genérica de Valores"
   ClientHeight    =   8445
   ClientLeft      =   5940
   ClientTop       =   3300
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   12450
   Begin prjChameleon.chameleonButton cmdSelectAll 
      Height          =   285
      Left            =   10470
      TabIndex        =   61
      Top             =   1290
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Selecionar Tudo"
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
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSimuladoPG.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      Height          =   5655
      Left            =   90
      TabIndex        =   30
      Top             =   2790
      Width           =   12255
      Begin VB.CheckBox chkRedutor 
         Caption         =   "Calcular com redutor territorial"
         Height          =   225
         Left            =   180
         TabIndex        =   75
         Top             =   1770
         Value           =   1  'Checked
         Width           =   2595
      End
      Begin VB.CheckBox chkCategNova 
         Caption         =   "Calcular com categoria nova"
         Height          =   225
         Left            =   3000
         TabIndex        =   73
         Top             =   1770
         Value           =   1  'Checked
         Width           =   2385
      End
      Begin VB.CheckBox chkAgrupAntigo 
         Caption         =   "Calcular pelo agrupamento antigo"
         Enabled         =   0   'False
         Height          =   225
         Left            =   6360
         TabIndex        =   72
         Top             =   1800
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.CheckBox chkQuadra 
         Caption         =   "Todas as quadras"
         Height          =   195
         Left            =   5760
         TabIndex        =   59
         Top             =   660
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.TextBox txtQuadraIni 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   56
         Top             =   990
         Width           =   675
      End
      Begin VB.TextBox txtQuadraFim 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   55
         Top             =   1350
         Width           =   675
      End
      Begin VB.TextBox txtAliqT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10080
         TabIndex        =   54
         Text            =   "0,5"
         Top             =   1350
         Width           =   675
      End
      Begin VB.TextBox txtAliqP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   10080
         TabIndex        =   53
         Text            =   "1,5"
         Top             =   1020
         Width           =   675
      End
      Begin VB.ComboBox cmbTipoCalc 
         Height          =   315
         ItemData        =   "frmSimuladoPG.frx":001C
         Left            =   10080
         List            =   "frmSimuladoPG.frx":0029
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   660
         Width           =   1935
      End
      Begin VB.ComboBox cmbSetorSimulado 
         Height          =   315
         ItemData        =   "frmSimuladoPG.frx":005C
         Left            =   1770
         List            =   "frmSimuladoPG.frx":0075
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   570
         Width           =   1215
      End
      Begin VB.ComboBox cmbAgAtual 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   960
         Width           =   3315
      End
      Begin VB.ComboBox cmbAg 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1350
         Width           =   3315
      End
      Begin MSComctlLib.ListView lvImovel 
         Height          =   2595
         Left            =   90
         TabIndex        =   32
         Top             =   2100
         Width           =   12075
         _ExtentX        =   21299
         _ExtentY        =   4577
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
         NumItems        =   29
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Setor"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Quadra"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Proprietario"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Endereço"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Bairro"
            Object.Width           =   3598
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Area T"
            Object.Width           =   1481
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Área C"
            Object.Width           =   1480
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Testada"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "F.Ideal"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   11
            Text            =   "Aliq"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "FCat"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "FPed"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "FSit"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "FPro"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "FTop"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "FGle"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "Agr1"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   19
            Text            =   "VVT"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   20
            Text            =   "VVC"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   21
            Text            =   "VVI"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   22
            Text            =   "IPTU"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   23
            Text            =   "Agr2"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   24
            Text            =   "VVT2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   25
            Text            =   "VVC2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   26
            Text            =   "VVI2"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   27
            Text            =   "IPTU2"
            Object.Width           =   1834
         EndProperty
         BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   28
            Text            =   "Redutor"
            Object.Width           =   1764
         EndProperty
      End
      Begin prjChameleon.chameleonButton cmdCalc 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   180
         TabIndex        =   33
         ToolTipText     =   "Calcular novo IPTU"
         Top             =   5130
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Calcular"
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
         MICON           =   "frmSimuladoPG.frx":0094
         PICN            =   "frmSimuladoPG.frx":00B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   165
         Left            =   210
         TabIndex        =   34
         Top             =   4890
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prjChameleon.chameleonButton cmdExportar 
         Height          =   315
         Left            =   2700
         TabIndex        =   35
         ToolTipText     =   "Exportar para Excel"
         Top             =   5160
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Exportar"
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
         MICON           =   "frmSimuladoPG.frx":01DF
         PICN            =   "frmSimuladoPG.frx":01FB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancelar 
         Height          =   315
         Left            =   1440
         TabIndex        =   36
         ToolTipText     =   "Cancelar o simulado"
         Top             =   5160
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cance&lar  "
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
         MICON           =   "frmSimuladoPG.frx":057D
         PICN            =   "frmSimuladoPG.frx":0599
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblVVPNew 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   10530
         TabIndex        =   71
         Top             =   5340
         Width           =   1395
      End
      Begin VB.Label lblVVPOld 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   10530
         TabIndex        =   70
         Top             =   5070
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Predial:"
         Height          =   225
         Index           =   6
         Left            =   9900
         TabIndex        =   69
         Top             =   5340
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Predial.:"
         Height          =   225
         Index           =   5
         Left            =   9900
         TabIndex        =   68
         Top             =   5070
         Width           =   585
      End
      Begin VB.Label lblVVTNew 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   8340
         TabIndex        =   67
         Top             =   5340
         Width           =   1395
      End
      Begin VB.Label lblVVTOld 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   8340
         TabIndex        =   66
         Top             =   5070
         Width           =   1395
      End
      Begin VB.Label Label4 
         Caption         =   "Territ..:"
         Height          =   225
         Index           =   4
         Left            =   7770
         TabIndex        =   65
         Top             =   5340
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Territ..:"
         Height          =   225
         Index           =   3
         Left            =   7770
         TabIndex        =   64
         Top             =   5070
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra Inicial...:"
         Height          =   225
         Index           =   7
         Left            =   5730
         TabIndex        =   58
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra Final.....:"
         Height          =   225
         Index           =   3
         Left            =   5730
         TabIndex        =   57
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Cálculo.........:"
         Height          =   195
         Left            =   8430
         TabIndex        =   52
         Top             =   780
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Aliquota Territorial %..:"
         Height          =   195
         Left            =   8430
         TabIndex        =   50
         Top             =   1410
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "Aliquota Predial %......:"
         Height          =   195
         Left            =   8430
         TabIndex        =   49
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Setor........................:"
         Height          =   195
         Left            =   150
         TabIndex        =   48
         Top             =   630
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Agrupamento 2017.:"
         Height          =   195
         Left            =   180
         TabIndex        =   46
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Agrupamento novo..:"
         Height          =   195
         Left            =   150
         TabIndex        =   44
         Top             =   1410
         Width           =   1635
      End
      Begin VB.Label Label4 
         Caption         =   "IPTU lançado em 2017.........:"
         Height          =   225
         Index           =   0
         Left            =   4020
         TabIndex        =   43
         Top             =   5070
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "IPTU c/ agrupamento novo..:"
         Height          =   225
         Index           =   1
         Left            =   4020
         TabIndex        =   42
         Top             =   5340
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Quantidade de imóveis..........:"
         Height          =   225
         Index           =   2
         Left            =   4020
         TabIndex        =   41
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label lblQtdeImovel 
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   6270
         TabIndex        =   40
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label lblIptuAtual 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   6270
         TabIndex        =   39
         Top             =   5070
         Width           =   1395
      End
      Begin VB.Label lblIptu2 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   6270
         TabIndex        =   38
         Top             =   5340
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SIMULADOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   2850
         TabIndex        =   37
         Top             =   210
         Width           =   6525
      End
   End
   Begin VB.ListBox lstQuadra 
      Appearance      =   0  'Flat
      Height          =   1830
      Left            =   8610
      Style           =   1  'Checkbox
      TabIndex        =   28
      Top             =   690
      Width           =   1635
   End
   Begin VB.ComboBox cmbSetor2 
      Height          =   315
      ItemData        =   "frmSimuladoPG.frx":06F3
      Left            =   6570
      List            =   "frmSimuladoPG.frx":0709
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   810
      Width           =   855
   End
   Begin VB.ComboBox cmbAg3 
      Height          =   315
      ItemData        =   "frmSimuladoPG.frx":071F
      Left            =   6540
      List            =   "frmSimuladoPG.frx":0721
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   420
      Width           =   1605
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2835
      Left            =   5100
      TabIndex        =   21
      Top             =   -30
      Width           =   105
   End
   Begin VB.ComboBox cmbAg2 
      Height          =   315
      ItemData        =   "frmSimuladoPG.frx":0723
      Left            =   1530
      List            =   "frmSimuladoPG.frx":0725
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1590
      Width           =   1245
   End
   Begin VB.ComboBox cmbSetor 
      Height          =   315
      ItemData        =   "frmSimuladoPG.frx":0727
      Left            =   1530
      List            =   "frmSimuladoPG.frx":073D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox txtQuadra2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1530
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1230
      Width           =   675
   End
   Begin VB.TextBox txtQuadra1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1530
      MaxLength       =   4
      TabIndex        =   1
      Top             =   870
      Width           =   675
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   14415
      MaxLength       =   10
      TabIndex        =   9
      Top             =   13245
      Width           =   915
   End
   Begin VB.ListBox lstBairro 
      Appearance      =   0  'Flat
      Height          =   1155
      Left            =   18840
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   12000
      Width           =   7305
   End
   Begin VB.ComboBox cmbAgrupa 
      Height          =   315
      ItemData        =   "frmSimuladoPG.frx":0753
      Left            =   20100
      List            =   "frmSimuladoPG.frx":0755
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   11580
      Width           =   1965
   End
   Begin prjChameleon.chameleonButton cmdLoadImovel 
      Height          =   345
      Left            =   24540
      TabIndex        =   7
      ToolTipText     =   "Carrega os imóveis"
      Top             =   13320
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Gerar"
      ENAB            =   0   'False
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
      MICON           =   "frmSimuladoPG.frx":0757
      PICN            =   "frmSimuladoPG.frx":0773
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAtualizar 
      Height          =   315
      Left            =   570
      TabIndex        =   22
      ToolTipText     =   "Alterar o valor do agrupamento para as quadras selecionadas"
      Top             =   2010
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Atualizar"
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
      MICON           =   "frmSimuladoPG.frx":0812
      PICN            =   "frmSimuladoPG.frx":082E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton chkMudar 
      Height          =   735
      Left            =   10440
      TabIndex        =   60
      ToolTipText     =   "Exportar para Excel"
      Top             =   1770
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "Mudar as quadras selecionadas para novo agrupamento"
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
      MICON           =   "frmSimuladoPG.frx":0988
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
      Height          =   525
      Left            =   2850
      TabIndex        =   63
      ToolTipText     =   "Editar os valores dos agrupamentos"
      Top             =   2040
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Editar valores dos agrupamentos"
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
      MICON           =   "frmSimuladoPG.frx":09A4
      PICN            =   "frmSimuladoPG.frx":09C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar2 
      Height          =   525
      Left            =   2850
      TabIndex        =   74
      ToolTipText     =   "Editar os valores dos agrupamentos"
      Top             =   1440
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   926
      BTYPE           =   3
      TX              =   "&Editar fatores de construção"
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
      MICON           =   "frmSimuladoPG.frx":0B1A
      PICN            =   "frmSimuladoPG.frx":0B36
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblQtdeQuadra 
      Caption         =   "0 Quadras"
      Height          =   225
      Left            =   8610
      TabIndex        =   62
      Top             =   2520
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quadras..:"
      Height          =   225
      Index           =   6
      Left            =   8610
      TabIndex        =   29
      Top             =   450
      Width           =   1245
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Setor................:"
      Height          =   195
      Left            =   5280
      TabIndex        =   27
      Top             =   870
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agrupamento....:"
      Height          =   225
      Index           =   5
      Left            =   5280
      TabIndex        =   25
      Top             =   510
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CONSULTAR QUADRAS POR AGRUPAMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   2
      Left            =   5370
      TabIndex        =   23
      Top             =   90
      Width           =   5745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agrup. Novo.....:"
      Height          =   225
      Index           =   4
      Left            =   300
      TabIndex        =   20
      Top             =   1620
      Width           =   1245
   End
   Begin VB.Label Label6 
      Caption         =   "Setor................:"
      Height          =   195
      Left            =   300
      TabIndex        =   19
      Top             =   540
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quadra Final.....:"
      Height          =   225
      Index           =   2
      Left            =   300
      TabIndex        =   18
      Top             =   1260
      Width           =   1245
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Quadra Inicial...:"
      Height          =   225
      Index           =   0
      Left            =   300
      TabIndex        =   17
      Top             =   900
      Width           =   1245
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ALTERAR QUADRAS PARA O NOVO AGRUPAMENTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   16
      Top             =   90
      Width           =   4725
   End
   Begin VB.Label lblQtde 
      Caption         =   "0"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   17430
      TabIndex        =   15
      Top             =   13335
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Qtde:"
      Height          =   195
      Index           =   3
      Left            =   16800
      TabIndex        =   14
      Top             =   13335
      Width           =   555
   End
   Begin VB.Label lblIPTUNovo 
      Caption         =   "0,00"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   22560
      TabIndex        =   13
      Top             =   13335
      Width           =   1725
   End
   Begin VB.Label lblIPTU2013 
      Caption         =   "0,00"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   19680
      TabIndex        =   12
      Top             =   13335
      Width           =   1500
   End
   Begin VB.Label Label2 
      Caption         =   "IPTU NOVO:"
      Height          =   195
      Index           =   2
      Left            =   21300
      TabIndex        =   11
      Top             =   13335
      Width           =   1140
   End
   Begin VB.Label lblAnoIPTU 
      Caption         =   "IPTU 2013:"
      Height          =   195
      Left            =   18630
      TabIndex        =   10
      Top             =   13320
      Width           =   915
   End
   Begin VB.Label Label2 
      Caption         =   "Novo valor:"
      Height          =   195
      Index           =   0
      Left            =   13425
      TabIndex        =   8
      Top             =   13290
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Agrupamento..:"
      Height          =   225
      Index           =   1
      Left            =   18930
      TabIndex        =   4
      Top             =   11640
      Width           =   1125
   End
End
Attribute VB_Name = "frmSimuladoPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bStop As Boolean

Private Sub chkMudar_Click()
Dim x As Integer, bFind As Boolean, sQuadras, z As Variant
bFind = False
sQuadras = ""
For x = 0 To lstQuadra.ListCount - 1
    If lstQuadra.Selected(x) = True Then
        bFind = True
        sQuadras = sQuadras & lstQuadra.List(x) & ","
    End If
Next


If Not bFind Then
    MsgBox "Nenhuma quadra selecionada.", vbCritical, "Erro"
Else
    sQuadras = Left(sQuadras, Len(sQuadras) - 1)
    z = InputBox("Digite o novo agrupamento de 1 à 8", "Novo Agrupamento")
    If Val(z) < 1 Or Val(z) > 8 Then
        MsgBox "Código de agrupamento inválido.", vbCritical, "Erro"
    Else
        If MsgBox("Deseja alterar a(s) quadra(s) " & sQuadras & " para o agrupamento " & CStr(z) & "?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            For x = 0 To lstQuadra.ListCount - 1
                If lstQuadra.Selected(x) = True Then
                    Sql = "update facequadra set codagrupanovo=" & Val(z) & " where coddistrito=1 and codsetor=" & Val(cmbSetor2.Text) & " and codquadra=" & lstQuadra.List(x)
                    cn.Execute Sql, rdExecDirect
                End If
            Next
        End If
    End If
End If

End Sub

Private Sub chkQuadra_Click()

If chkQuadra.value = vbChecked Then
    txtQuadraIni.BackColor = vbButtonFace
    txtQuadraFim.BackColor = vbButtonFace
    txtQuadraIni.Locked = True
    txtQuadraFim.Locked = True
    txtQuadraIni.Text = ""
    txtQuadraFim.Text = ""
Else
    txtQuadraIni.BackColor = vbWhite
    txtQuadraFim.BackColor = vbWhite
    txtQuadraIni.Locked = False
    txtQuadraFim.Locked = False
End If

End Sub

Private Sub cmbAg3_Click()
If cmbSetor2.ListIndex > -1 And cmbAg3.ListIndex > -1 Then
    CarregaQuadra cmbSetor2.Text, cmbAg3.ItemData(cmbAg3.ListIndex)
End If
End Sub

Private Sub cmbAgrupa_Click()
Dim Sql As String, RdoAux As rdoResultset

lstBairro.Clear

Sql = "SELECT DISTINCT li_codbairro,descbairro From vwFULLIMOVEL2 Where codagrupa =" & cmbAgrupa.ItemData(cmbAgrupa.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If Not IsNull(!DescBairro) Then
            lstBairro.AddItem (!DescBairro)
            lstBairro.ItemData(lstBairro.NewIndex) = !Li_CodBairro
        End If
       .MoveNext
    Loop
   .Close
End With

For x = 0 To lstBairro.ListCount - 1
    lstBairro.Selected(x) = True
Next

End Sub

Private Sub cmbSetor2_Click()
If cmbSetor2.ListIndex > -1 And cmbAg3.ListIndex > -1 Then
    CarregaQuadra cmbSetor2.Text, cmbAg3.ItemData(cmbAg3.ListIndex)
End If
End Sub

Private Sub cmdAlterar_Click()
frmAgrupamento.show
frmAgrupamento.ZOrder 0
End Sub

Private Sub cmdAlterar2_Click()
frmCategConstr.show
frmCategConstr.ZOrder 0
End Sub

Private Sub cmdAtualizar_Click()

Dim nQuadra1 As Integer, nQuadra2 As Integer, nAgrup As Integer, nSetor As Integer
If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "FACTORE" Then
    MsgBox "Acesso negado!", vbCritical
    Exit Sub
End If
nSetor = Val(cmbSetor.Text)
nQuadra1 = Val(txtQuadra1.Text)
nQuadra2 = Val(txtQuadra2.Text)
nAgrup = cmbAg2.ItemData(cmbAg2.ListIndex)

If nQuadra1 = 0 Or nQuadra2 = 0 Then
    MsgBox "Digite quadra inicial e final", vbCritical, "Erro"
    Exit Sub
End If

If nQuadra1 > nQuadra2 Then
    MsgBox "Quadra inicial não pode ser maior que a final", vbCritical, "Erro"
    Exit Sub
End If

If MsgBox("Deseja atualizar as quadras de " & nQuadra1 & " até " & nQuadra2 & " do setor " & nSetor & " para o agrupamento: " & cmbAg2.Text & "?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

Sql = "update facequadra set codagrupanovo=" & nAgrup & ",alterado='" & NomeDeLogin & "' where codsetor=" & nSetor & " and codquadra between " & nQuadra1 & " and " & nQuadra2
cn.Execute Sql, rdExecDirect

MsgBox "Quadras alteradas.", vbInformation, "Infomação"
txtQuadra1.Text = ""
txtQuadra2.Text = ""
txtQuadra1.SetFocus
End Sub

Private Sub cmdCalc_Click()

'If chkSimula4.value = vbChecked Or chkSimula3.value = vbChecked Then
    If NomeDeLogin = "SCHWARTZ" Then
        Simulado
        Exit Sub
    End If
    
'        MsgBox "Simulador 3% e 4% não disponível para o seu Login.", vbCritical, "ERRO"
'    Else
'        Calculo4
'    End If
'Else
    Calculo
'End If
Exit Sub

Dim Sql As String, qd As New rdoQuery, RdoAux As rdoResultset, x As Integer, nCodReduz As Long, nAliq As Double
Dim nValorIptu As Double, nVVI As Double, nSomaIPTUNovo As Double, nPos As Long, nTot As Long

nSomaIPTUNovo = 0
If Val(txtValor.Text) = 0 Then
    MsgBox "Digite um valor válido", vbCritical, "Erro"
    Exit Sub
End If

If lvImovel.ListItems.Count = 0 Then
    MsgBox "Nenhum imóvel carregado.", vbCritical, "Erro"
    Exit Sub
End If
nPos = 1
If MsgBox("Deseja realizar o simulado de cálculo?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

Ocupado
nTot = lvImovel.ListItems.Count
For x = 1 To lvImovel.ListItems.Count
    If nPos Mod 5 = 0 Then
        CallPb nPos, nTot
    End If
    nCodReduz = Val(lvImovel.ListItems(x).Text)
    nAliq = CDbl(lvImovel.ListItems(x).SubItems(7))
    Set qd.ActiveConnection = cn
    qd.QueryTimeout = 0
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spCALCULOPG(?,?,?) }"
    qd(0) = nCodReduz
    qd(1) = Year(Now)
    qd(2) = Virg2Ponto(txtValor.Text)
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        nVVI = !VVI
        nValorIptu = nVVI * (nAliq / 100)
        lvImovel.ListItems(x).SubItems(9) = FormatNumber(nValorIptu, 2)
        If CDbl(lvImovel.ListItems(x).SubItems(9)) - CDbl(lvImovel.ListItems(x).SubItems(8)) < 1 Then
            lvImovel.ListItems(x).SubItems(9) = lvImovel.ListItems(x).SubItems(8)
            nValorIptu = CDbl(lvImovel.ListItems(x).SubItems(8))
        End If
        nSomaIPTUNovo = nSomaIPTUNovo + nValorIptu
        DoEvents
       .Close
    End With
    nPos = nPos + 1
Next
Liberado
lblIPTUNovo.Caption = FormatNumber(nSomaIPTUNovo, 2)
Pb.value = 0
End Sub

Private Sub cmdCancelar_Click()
If MsgBox("Deseja cancelar a simulação de cálculo?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub
bStop = True

End Sub

Private Sub cmdExportar_Click()
Exporta
Exit Sub
Dim ax As String, z As Long, x As Integer, Y As Integer, sChar As String

If grdMain.Rows = 1 Then
    MsgBox "Calcule primeiro."
    Exit Sub
Else
    If grdMain.TextMatrix(1, 1) = "" Then
        MsgBox "Calcule primeiro."
        Exit Sub
    End If
End If

'If txtSep.Text = "" Then
'    sChar = " "
'Else
'    sChar = txtSep.Text
'End If
sChar = "#"

Ocupado
Open sPathBin & "\SIMULADO.TXT" For Output As #1

With lvImovel
    ax = ""
    For Y = 1 To .ColumnHeaders.Count
        ax = ax & .ColumnHeaders(Y).Text & sChar
    Next
    ax = Chomp(ax, chomp_righT, 1)
    Print #1, ax
    
    For x = 1 To .ListItems.Count
        ax = .ListItems(x).Text & sChar & .ListItems(x).SubItems(1) & sChar & .ListItems(x).SubItems(2) & sChar & .ListItems(x).SubItems(3) & sChar _
        & .ListItems(x).SubItems(4) & sChar & .ListItems(x).SubItems(5) & sChar & .ListItems(x).SubItems(6) & sChar & .ListItems(x).SubItems(7) & sChar _
        & .ListItems(x).SubItems(8) & sChar & .ListItems(x).SubItems(9) & sChar
        ax = Chomp(ax, chomp_righT, 1)
        Print #1, ax
    Next
    

End With

Close #1
Liberado
MsgBox "O arquivo foi salvo em " & sPathBin & "\SIMULADO.TXT"

End Sub

Private Sub cmdLoadImovel_Click()
Dim itmX As ListItem, Sql As String, RdoAux As rdoResultset, nTot As Long, nPos As Long
Dim z As Long, x As Integer, nAgrupa As Integer, nBairro As Integer, nSomaIptu2013 As Double
z = SendMessage(lvImovel.hwnd, LVM_DELETEALLITEMS, 0, 0)
nSomaIptu2013 = 0
nAgrupa = cmbAgrupa.ItemData(cmbAgrupa.ListIndex)
nPos = 1

Ocupado
For x = 0 To lstBairro.ListCount - 1
    If lstBairro.Selected(x) = True Then
        nBairro = lstBairro.ItemData(x)
        Sql = "SELECT DISTINCT vwFULLIMOVEL2.codreduzido, vwFULLIMOVEL2.Ativo, vwFULLIMOVEL2.INSCRICAO, vwFULLIMOVEL2.nomecidadao, vwFULLIMOVEL2.CPF, "
        Sql = Sql & "vwFULLIMOVEL2.CNPJ, vwFULLIMOVEL2.rg, vwFULLIMOVEL2.LOGRADOURO, vwFULLIMOVEL2.li_num, vwFULLIMOVEL2.li_compl, vwFULLIMOVEL2.descbairro,"
        Sql = Sql & "vwFULLIMOVEL2.li_quadras, vwFULLIMOVEL2.li_lotes, vwFULLIMOVEL2.li_codbairro, vwFULLIMOVEL2.codlogr, vwFULLIMOVEL2.inativo,"
        Sql = Sql & "vwFULLIMOVEL2.dt_areaterreno, vwFULLIMOVEL2.dt_codusoterreno, vwFULLIMOVEL2.dt_codbenf, vwFULLIMOVEL2.dt_codtopog, vwFULLIMOVEL2.dt_codcategprop,"
        Sql = Sql & "vwFULLIMOVEL2.dt_codsituacao, vwFULLIMOVEL2.dt_codpedol, vwFULLIMOVEL2.dt_numagua, vwFULLIMOVEL2.dt_fracaoideal, vwFULLIMOVEL2.dc_qtdeedif,"
        Sql = Sql & "vwFULLIMOVEL2.dc_qtdepav, vwFULLIMOVEL2.ee_tipoend, vwFULLIMOVEL2.distrito, vwFULLIMOVEL2.setor, vwFULLIMOVEL2.quadra, vwFULLIMOVEL2.lote,"
        Sql = Sql & "vwFULLIMOVEL2.seq, vwFULLIMOVEL2.unidade, vwFULLIMOVEL2.subunidade, vwFULLIMOVEL2.li_uf, vwFULLIMOVEL2.li_codcidade,"
        Sql = Sql & "vwFULLIMOVEL2.descbenfeitoria, vwFULLIMOVEL2.descusoterreno, vwFULLIMOVEL2.desctopografia, vwFULLIMOVEL2.desccategprop,"
        Sql = Sql & "vwFULLIMOVEL2.descsituacao, vwFULLIMOVEL2.descpedologia, vwFULLIMOVEL2.codcidadao, vwFULLIMOVEL2.NOMELOGRADOURO2, vwFULLIMOVEL2.abrevtitlog,"
        Sql = Sql & "vwFULLIMOVEL2.abrevtipolog, vwFULLIMOVEL2.nomelogradouro, vwFULLIMOVEL2.numimovel, vwFULLIMOVEL2.complemento, vwFULLIMOVEL2.DESCBAIRROP,"
        Sql = Sql & "vwFULLIMOVEL2.siglauf, vwFULLIMOVEL2.codlogradouro, vwFULLIMOVEL2.ee_codlog, vwFULLIMOVEL2.ee_nomelog, vwFULLIMOVEL2.ee_numimovel,"
        Sql = Sql & "vwFULLIMOVEL2.ee_complemento, vwFULLIMOVEL2.BairroEE, vwFULLIMOVEL2.CidadeEE, vwFULLIMOVEL2.ee_uf, vwFULLIMOVEL2.ee_cep,"
        Sql = Sql & "vwFULLIMOVEL2.ee_descbairro, vwFULLIMOVEL2.AbrevTipoLogEE, vwFULLIMOVEL2.AbrevTitLogEE, vwFULLIMOVEL2.cd_nomecond,"
        Sql = Sql & "vwFULLIMOVEL2.codcondominio, vwFULLIMOVEL2.datainclusao, vwFULLIMOVEL2.codagrupa, vwFULLIMOVEL2.desccidade, laseriptu.vvt, laseriptu.vvc, laseriptu.vvi,"
        Sql = Sql & "LaserIPTU.impostopredial , LaserIPTU.impostoterritorial, LaserIPTU.valortotalparc, LaserIPTU.areaconstrucao,areaterreno,qtdeparc,aliquota FROM vwFULLIMOVEL2 INNER JOIN "
        Sql = Sql & "laseriptu ON vwFULLIMOVEL2.codreduzido = laseriptu.codreduzido Where (LaserIPTU.Ano = " & Year(Now) & ") and codagrupa =" & nAgrupa & " and li_codbairro=" & nBairro & " and ativo='S'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nTot = .RowCount
            Do Until .EOF
                If nPos Mod 5 = 0 Then
                    CallPb nPos, nTot
                End If
                Set itmX = lvImovel.ListItems.Add(, , !CODREDUZIDO)
                itmX.SubItems(1) = SubNull(!DescBairro)
                itmX.SubItems(2) = FormatNumber(!AreaTerreno, 2)
                itmX.SubItems(3) = FormatNumber(!areaconstrucao, 2)
                itmX.SubItems(4) = FormatNumber(!vvt, 2)
                itmX.SubItems(5) = FormatNumber(!vvc, 2)
                itmX.SubItems(6) = FormatNumber(!VVI, 2)
                itmX.SubItems(7) = FormatNumber(!Aliquota, 2)
                itmX.SubItems(8) = FormatNumber((!valortotalparc * !qtdeparc), 2)
                nSomaIptu2013 = nSomaIptu2013 + (!valortotalparc * !qtdeparc)
                nPos = nPos + 1
               .MoveNext
            Loop
           .Close
        End With
    End If
Next
lblIPTU2013.Caption = FormatNumber(nSomaIptu2013, 2)
lblQtde.Caption = lvImovel.ListItems.Count
Liberado
Pb.value = 0
End Sub

Private Sub cmdSelectAll_Click()
Dim x As Integer
For x = 0 To lstQuadra.ListCount - 1
    lstQuadra.Selected(x) = True
Next
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset



Dim styles As Long
    styles = SendMessage(lvImovel.hwnd, _
        LVM_GETEXTENDEDLISTVIEWSTYLE, 0, ByVal 0&)
    styles = Style Or LVS_EX_DOUBLEBUFFER Or LVS_EX_BORDERSELECT
    Call SendMessage(lvImovel.hwnd, _
        LVM_SETEXTENDEDLISTVIEWSTYLE, 0, ByVal styles)

Centraliza Me
cmbAg.AddItem "(Todos os agrupamentos)"
cmbAg.ItemData(cmbAg.NewIndex) = 0
cmbSetor.ListIndex = 0
Sql = "select codigo,valor from agrupamento where ano=" & Year(Now) & " order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAg.AddItem !Codigo & " - " & Format(!Valor, "#0.00")
        cmbAg.ItemData(cmbAg.NewIndex) = !Codigo
        cmbAg2.AddItem !Codigo & " - " & Format(!Valor, "#0.00")
        cmbAg2.ItemData(cmbAg2.NewIndex) = !Codigo
        cmbAg3.AddItem !Codigo & " - " & Format(!Valor, "#0.00")
        cmbAg3.ItemData(cmbAg3.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
cmbAg.ListIndex = 0
cmbAg2.ListIndex = 0
cmbAg3.ListIndex = 0


cmbAgAtual.AddItem "(Todos os agrupamentos)"
cmbAgAtual.ItemData(cmbAgAtual.NewIndex) = 0
Sql = "select codagrupamento,valorterreno from terreno where anofator=" & Year(Now) & " order by codagrupamento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAgAtual.AddItem !codagrupamento & " - " & Format(!valorterreno, "#0.00")
        cmbAgAtual.ItemData(cmbAgAtual.NewIndex) = !codagrupamento
       .MoveNext
    Loop
   .Close
End With
cmbAgAtual.ListIndex = 0

cmbSetorSimulado.ListIndex = 0
cmbTipoCalc.ListIndex = 0

End Sub

Private Sub txtAliqP_KeyPress(KeyAscii As Integer)
Tweak txtAliqP, KeyAscii, DecimalPositive
End Sub

Private Sub txtAliqT_KeyPress(KeyAscii As Integer)
Tweak txtAliqT, KeyAscii, DecimalPositive
End Sub

Private Sub txtQuadra1_KeyPress(KeyAscii As Integer)
Tweak txtQuadra1, KeyAscii, IntegerPositive
End Sub

Private Sub txtQuadra2_KeyPress(KeyAscii As Integer)
Tweak txtQuadra2, KeyAscii, IntegerPositive
End Sub


Private Sub txtQuadraFim_KeyPress(KeyAscii As Integer)
Tweak txtQuadraIni, KeyAscii, IntegerPositive
End Sub


Private Sub txtQuadraIni_KeyPress(KeyAscii As Integer)
Tweak txtQuadraIni, KeyAscii, IntegerPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub Calculo()

Dim Sql As String, qd As New rdoQuery, RdoAux As rdoResultset, x As Integer, nCodReduz As Long, nAliq As Double, RdoAux2 As rdoResultset, sTipo As String
Dim nValorIptu As Double, nSomaIPTUNovo As Double, nPos As Long, nTot As Long, nValorIPTU2 As Double, sProprietario As String, sEndereco As String
Dim sBairro As String, nAreaT As Double, nAreaC As Double, nVVT As Double, nVVC As Double, nVVI As Double, nVVT2 As Double, nVVC2 As Double, nVVI2 As Double
Dim nSetor As Integer, nQuadra As Integer, nFracao As Double, nTestada As Double, nFCat As Double, nFPed As Double, nFSit As Double, nFPro As Double, nFTop As Double, nFGle As Double
Dim nValorAgrupa1 As Double, nValorAgrupa2 As Double, nSoma1 As Double, nSoma2 As Double, nLinhas As Long, nSomaVVTOld As Double, nSomaVVTNew As Double, nSomaVVPOld As Double, nSomaVVPNew As Double
Dim nQuadra1 As Integer, nQuadra2 As Integer, nAgrup As Integer, sTipoCalc As String, nRedutor As Double

If Not IsNumeric(txtAliqP.Text) Then
    MsgBox "Digite a aliquota predial", vbCritical, "Erro"
    Exit Sub
End If
If Not IsNumeric(txtAliqT.Text) Then
    MsgBox "Digite a aliquota territorial", vbCritical, "Erro"
    Exit Sub
End If
If Val(txtAliqP.Text) > 100 Then
    MsgBox "Aliquota predial maior que 100%", vbCritical, "Erro"
    Exit Sub
End If
If Val(txtAliqT.Text) > 100 Then
    MsgBox "Aliquota territorial maior que 100%", vbCritical, "Erro"
    Exit Sub
End If

If chkQuadra.value = vbUnchecked And cmbSetorSimulado.ListIndex = 0 Then
    MsgBox "Para calcular todas as quadras você deve especificar um setor", vbCritical, "Erro"
    Exit Sub
End If

If chkQuadra.value = vbUnchecked Then
    If Not IsNumeric(txtQuadraIni.Text) Then
        MsgBox "Digite quadra inicial", vbCritical, "Erro"
        Exit Sub
    End If
    
    If Not IsNumeric(txtQuadraFim.Text) Then
        MsgBox "Digite quadra Final", vbCritical, "Erro"
        Exit Sub
    End If
    
End If

nSetor = Val(cmbSetorSimulado.Text)
nQuadra1 = Val(txtQuadraIni.Text)
nQuadra2 = Val(txtQuadraFim.Text)
sTipoCalc = ""
If cmbTipoCalc.ListIndex = 1 Then
    sTipoCalc = "P"
ElseIf cmbTipoCalc.ListIndex = 2 Then
    sTipoCalc = "T"
End If

If MsgBox("Deseja realizar o simulado de cálculo?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

bStop = False
lblQtde.Caption = "0,00"
lblIptuAtual.Caption = "0,00"
lblIptu2.Caption = "0,00"
lblVVPNew.Caption = "0,00"
lblVVTNew.Caption = "0,00"
lblVVPOld.Caption = "0,00"
lblVVTOld.Caption = "0,00"
nSoma1 = 0: nSoma2 = 0
nSomaVVPNew = 0: nSomaVVPOld = 0: nSomaVVTNew = 0: nSomaVVTOld = 0

Ocupado
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
lvImovel.ListItems.Clear

Sql = "SELECT DISTINCT setor, quadra, codreduzido, codcidadao,nomecidadao,INSCRICAO, LOGRADOURO, li_num, li_compl, descbairro, dt_areaterreno, dc_qtdepav, descbenfeitoria, descusoterreno,"
Sql = Sql & "desctopografia, desccategprop, descsituacao, descpedologia, cd_nomecond, codagrupa, codagrupanovo, desccidade, dt_fracaoideal, setor, quadra,"
Sql = Sql & "Distrito , Lote, Seq FROM vwFULLIMOVEL2 WHERE (Ativo = 'S') "
If cmbAg.ListIndex > 0 Then
    Sql = Sql & " and codagrupanovo=" & cmbAg.ItemData(cmbAg.ListIndex)
End If
If cmbAgAtual.ListIndex > 0 Then
    Sql = Sql & " and codagrupa=" & cmbAgAtual.ItemData(cmbAgAtual.ListIndex)
End If
If cmbSetorSimulado.ListIndex > 0 Then
    Sql = Sql & " and setor=" & nSetor
End If
If chkQuadra.value = vbUnchecked Then
    Sql = Sql & " and quadra between " & nQuadra1 & " and " & nQuadra2
    cn.Execute Sql, rdExecDirect
End If

Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nLinhas = RdoAux2.RowCount
nTot = 0
If nLinhas = 0 Then
    Liberado
    MsgBox "Não existem quadras no Setor " & nSetor & " com o agrupamento " & cmbAg.ItemData(cmbAg.ListIndex)
    Exit Sub
End If
With RdoAux2
    LockWindowUpdate lvImovel.hwnd
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nLinhas
            DoEvents
        End If
        If bStop Then
            MsgBox "Cálculo cancelado", vbCritical, "Atenção"
            Exit Do
        End If
        nCodReduz = !CODREDUZIDO
        sBairro = RdoAux2!DescBairro
        
        sProprietario = !nomecidadao
        sEndereco = !Logradouro & "," & !Li_Num
        nSetor = !Setor
        nQuadra = !Quadra
        nFracao = !Dt_FracaoIdeal
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        qd.Sql = "{ Call spCALCULOPG(?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = Year(Now)
        qd(2) = !CODAGRUPA
        qd(3) = 0
        qd(4) = 1.5
        qd(5) = 3
        qd(6) = 0
        qd(7) = 0
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            nAreaT = !AreaTerreno
            nAreaC = !areapredial
            If nAreaC = 0 Then
                sTipo = "Terreno"
                If sTipoCalc = "P" Then GoTo proximo
            Else
                sTipo = "Predial"
                If sTipoCalc = "T" Then GoTo proximo
            End If
            nVVT = !vvt
            nVVC = !VVP
            nVVI = !VVI
            nAliq = !Aliquota
            nValorIptu = !valorfinal
            nTestada = !TESTADAPRINC
            nFCat = !fcat
            nFPed = !fped
            nFSit = !fsit
            nFPro = !fpro
            nFGle = !fgle
            nValorAgrupa1 = !valorAgrupamento
            nSoma1 = nSoma1 + !valorfinal
            If nVVC = 0 Then
                nSomaVVTOld = nSomaVVTOld + !valorfinal
            Else
                nSomaVVPOld = nSomaVVPOld + !valorfinal
            End If
            DoEvents
           .Close
        End With
        
        qd.Sql = "{ Call spCALCULOPG(?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = Year(Now)
        If chkAgrupAntigo.value = vbChecked Then
            qd(2) = !CODAGRUPA
            qd(3) = 0
        Else
            qd(2) = !codagrupanovo
            qd(3) = 1
        End If
        qd(4) = Virg2Ponto(txtAliqP.Text)
        qd(5) = Virg2Ponto(txtAliqT.Text)
        If chkCategNova.value = vbChecked Then
            qd(6) = 1
        Else
            qd(6) = 0
        End If
        If chkRedutor.value = vbChecked Then
            qd(7) = 1
        Else
            qd(7) = 0
        End If
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            nVVT2 = !vvt
            nVVC2 = !VVP
            nVVI2 = !VVI
            'If nVVC2 = 0 Then MsgBox "teste"
            nValorIPTU2 = !valorfinal
            nValorAgrupa2 = !valorAgrupamento
            nRedutor = !Redutor
            nSoma2 = nSoma2 + !valorfinal
            If nVVC2 = 0 Then
                nSomaVVTNew = nSomaVVTNew + !valorfinal
            Else
                nSomaVVPNew = nSomaVVPNew + !valorfinal
            End If
            DoEvents
           .Close
        End With
        
        Set itmX = lvImovel.ListItems.Add(, , nCodReduz)
        itmX.SubItems(1) = nSetor
        itmX.SubItems(2) = nQuadra
        itmX.SubItems(3) = sProprietario
        itmX.SubItems(4) = sEndereco
        itmX.SubItems(5) = sBairro
        itmX.SubItems(6) = sTipo
        itmX.SubItems(7) = FormatNumber(nAreaT, 2)
        itmX.SubItems(8) = FormatNumber(nAreaC, 2)
        itmX.SubItems(9) = FormatNumber(nTestada, 2)
        itmX.SubItems(10) = FormatNumber(nFracao, 2)
        itmX.SubItems(11) = FormatNumber(nAliq * 100, 2) & "%"
        itmX.SubItems(12) = FormatNumber(nFCat, 2)
        itmX.SubItems(13) = FormatNumber(nFPed, 2)
        itmX.SubItems(14) = FormatNumber(nFSit, 2)
        itmX.SubItems(15) = FormatNumber(nFPro, 2)
        itmX.SubItems(16) = FormatNumber(nFTop, 2)
        itmX.SubItems(17) = FormatNumber(nFGle, 2)
        itmX.SubItems(18) = FormatNumber(nValorAgrupa1, 2)
        itmX.SubItems(19) = FormatNumber(nVVT, 2)
        itmX.SubItems(20) = FormatNumber(nVVC, 2)
        itmX.SubItems(21) = FormatNumber(nVVI, 2)
        itmX.SubItems(22) = FormatNumber(nValorIptu, 2)
        itmX.SubItems(23) = FormatNumber(nValorAgrupa2, 2)
        itmX.SubItems(24) = FormatNumber(nVVT2, 2)
        itmX.SubItems(25) = FormatNumber(nVVC2, 2)
        itmX.SubItems(26) = FormatNumber(nVVI2, 2)
        itmX.SubItems(27) = FormatNumber(nValorIPTU2, 2)
        itmX.SubItems(28) = FormatNumber(nRedutor, 2)
        nTot = nTot + 1
proximo:
        
        nPos = nPos + 1
       .MoveNext
    Loop
    LockWindowUpdate 0&
End With

lblQtdeImovel.Caption = nTot
lblIptuAtual.Caption = FormatNumber(nSoma1, 2)
lblIptu2.Caption = FormatNumber(nSoma2, 2)
lblVVTOld.Caption = FormatNumber(nSomaVVTOld, 2)
lblVVPOld.Caption = FormatNumber(nSomaVVPOld, 2)
lblVVTNew.Caption = FormatNumber(nSomaVVTNew, 2)
lblVVPNew.Caption = FormatNumber(nSomaVVPNew, 2)

Liberado
'lblIPTUNovo.Caption = FormatNumber(nSomaIPTUNovo, 2)
Pb.value = 0

End Sub

Private Sub Exporta()
Dim x As Long, Y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
         
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For Y = 1 To lvImovel.ColumnHeaders.Count
    If lvImovel.ColumnHeaders(Y).Width > 0 Then
        ax = ax & RemoveSpace(lvImovel.ColumnHeaders(Y)) & " char(255), "
    End If
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "[Table1$]", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To lvImovel.ListItems.Count
    Rs.AddNew
    nCont = 0
    Rs.Fields(nCont).value = lvImovel.ListItems(x).Text
    nCont = 1
    For Y = 1 To lvImovel.ColumnHeaders.Count - 1
        Rs.Fields(nCont).value = lvImovel.ListItems(x).SubItems(Y)
        nCont = nCont + 1
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)

End Sub

Private Sub CarregaQuadra(Setor As Integer, Agrupamento As Integer)
Dim Sql As String, RdoAux As rdoResultset
lblQtdeQuadra.Caption = "0 Quadras"
Me.Refresh
lstQuadra.Clear
Sql = "select distinct(codquadra) from facequadra where codsetor=" & Setor & " and codagrupanovo=" & Agrupamento & " order by codquadra"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstQuadra.AddItem !CODQUADRA
       .MoveNext
    Loop
   .Close
End With
lblQtdeQuadra.Caption = lstQuadra.ListCount & " Quadras"
Me.Refresh

End Sub

Private Sub Calculo4()

Dim Sql As String, qd As New rdoQuery, RdoAux As rdoResultset, x As Integer, nCodReduz As Long, nAliq As Double, RdoAux2 As rdoResultset, sTipo As String
Dim nValorIptu As Double, nSomaIPTUNovo As Double, nPos As Long, nTot As Long, nValorIPTU2 As Double, sProprietario As String, sEndereco As String
Dim sBairro As String, nAreaT As Double, nAreaC As Double, nVVT As Double, nVVC As Double, nVVI As Double, nVVT2 As Double, nVVC2 As Double, nVVI2 As Double
Dim nSetor As Integer, nQuadra As Integer, nFracao As Double, nTestada As Double, nFCat As Double, nFPed As Double, nFSit As Double, nFPro As Double, nFTop As Double, nFGle As Double
Dim nValorAgrupa1 As Double, nValorAgrupa2 As Double, nSoma1 As Double, nSoma2 As Double
Dim nQuadra1 As Integer, nQuadra2 As Integer, nAgrup As Integer

nSetor = Val(cmbSetor.Text)
nQuadra1 = Val(txtQuadra1.Text)
nQuadra2 = Val(txtQuadra2.Text)

If chkSetorQuadra.value = vbChecked Then
    If nQuadra1 = 0 Or nQuadra2 = 0 Then
        MsgBox "Digite quadra inicial e final", vbCritical, "Erro"
        Exit Sub
    End If
    
    If nQuadra1 > nQuadra2 Then
        MsgBox "Quadra inicial não pode ser maior que a final", vbCritical, "Erro"
        Exit Sub
    End If
End If


If MsgBox("Deseja realizar o simulado de cálculo?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

bStop = False
lblQtde.Caption = "0,00"
lblIptuAtual.Caption = "0,00"
lblIptu2.Caption = "0,00"
nSoma1 = 0
nSoma2 = 0

Ocupado
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
lvImovel.ListItems.Clear

Sql = "SELECT  DISTINCT setor,quadra,  vwFULLIMOVEL2.codreduzido,nomecidadao, vwFULLIMOVEL2.INSCRICAO, vwFULLIMOVEL2.LOGRADOURO, vwFULLIMOVEL2.li_num, vwFULLIMOVEL2.li_compl,vwFULLIMOVEL2.descbairro, vwFULLIMOVEL2.dt_areaterreno, vwFULLIMOVEL2.dc_qtdepav, vwFULLIMOVEL2.descbenfeitoria, vwFULLIMOVEL2.descusoterreno,"
Sql = Sql & "vwFULLIMOVEL2.desctopografia, vwFULLIMOVEL2.desccategprop, vwFULLIMOVEL2.descsituacao, vwFULLIMOVEL2.descpedologia, vwFULLIMOVEL2.codcidadao,vwFULLIMOVEL2.cd_nomecond, vwFULLIMOVEL2.codagrupa,vwFULLIMOVEL2.codagrupanovo, vwFULLIMOVEL2.desccidade, laseriptu.vvt, laseriptu.vvc, laseriptu.vvi, laseriptu.impostopredial,"
Sql = Sql & "laseriptu.impostoterritorial, laseriptu.valortotalparc, laseriptu.areaconstrucao, laseriptu.areaterreno, laseriptu.qtdeparc, laseriptu.aliquota,vwFULLIMOVEL2.Dt_FracaoIdeal , vwFULLIMOVEL2.Setor, vwFULLIMOVEL2.Quadra, vwFULLIMOVEL2.Distrito, vwFULLIMOVEL2.Lote, vwFULLIMOVEL2.Seq "
Sql = Sql & "FROM vwFULLIMOVEL2 INNER JOIN laseriptu ON vwFULLIMOVEL2.codreduzido = laseriptu.codreduzido WHERE (laseriptu.ano = 2017) AND (vwFULLIMOVEL2.Ativo = 'S') "
If cmbAg.ListIndex > 0 Then
    Sql = Sql & " and codagrupanovo=" & cmbAg.ItemData(cmbAg.ListIndex)
End If
If chkSetorQuadra.value = vbChecked Then
    Sql = Sql & " and setor=" & nSetor & " and quadra between " & nQuadra1 & " and " & nQuadra2
    cn.Execute Sql, rdExecDirect
End If
Sql = Sql & " and  vwFULLIMOVEL2.codreduzido in (SELECT codreduzido From cadimob WHERE (inativo = 0) AND (imune <> 1) AND (dt_codbenf <> 4))"

Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nTot = RdoAux2.RowCount
With RdoAux2
    Do Until .EOF
        If nPos Mod 100 = 0 Then
            CallPb nPos, nTot
        End If
        If bStop Then
            MsgBox "Cálculo cancelado", vbCritical, "Atenção"
            Exit Do
        End If
        nCodReduz = !CODREDUZIDO
        sBairro = RdoAux2!DescBairro
        nAreaT = RdoAux2!Dt_AreaTerreno
        nAreaC = RdoAux2!areaconstrucao
        If nAreaC = 0 Then
            sTipo = "Terreno"
        Else
            sTipo = "Predial"
            GoTo proximo
        End If
        sProprietario = !nomecidadao
        sEndereco = !Logradouro & "," & !Li_Num
        nSetor = !Setor
        nQuadra = !Quadra
        nFracao = !Dt_FracaoIdeal
    
        If chkSimula3.value = vbChecked Then
            qd.Sql = "{ Call spCALCULOPG(?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = Year(Now)
            qd(2) = !CODAGRUPA
            qd(3) = 0
        Else
            qd.Sql = "{ Call spCALCULO4(?,?) }"
            qd(0) = nCodReduz
            qd(1) = Year(Now)
        End If
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            nVVT = !vvt
            nVVC = !VVP
            nVVI = !VVI
            nAliq = !Aliquota
            nValorIptu = !valorfinal
            nTestada = !TESTADAPRINC
            nFCat = !fcat
            nFPed = !fped
            nFSit = !fsit
            nFPro = !fpro
            nFGle = !fgle
            nValorAgrupa1 = !valorAgrupamento
            nSoma1 = nSoma1 + !valorfinal
            DoEvents
           .Close
        End With
        
        qd.Sql = "{ Call spCALCULOPG(?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = Year(Now)
        qd(2) = !codagrupanovo
        qd(3) = 1
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            nVVT2 = !vvt
            nVVC2 = !VVP
            nVVI2 = !VVI
            nValorIPTU2 = !valorfinal
            nValorAgrupa2 = !valorAgrupamento
            nSoma2 = nSoma2 + !valorfinal
            DoEvents
           .Close
        End With
        
        Set itmX = lvImovel.ListItems.Add(, , nCodReduz)
        itmX.SubItems(1) = nSetor
        itmX.SubItems(2) = nQuadra
        itmX.SubItems(3) = sProprietario
        itmX.SubItems(4) = sEndereco
        itmX.SubItems(5) = sBairro
        itmX.SubItems(6) = sTipo
        itmX.SubItems(7) = FormatNumber(nAreaT, 2)
        itmX.SubItems(8) = FormatNumber(nAreaC, 2)
        itmX.SubItems(9) = FormatNumber(nTestada, 2)
        itmX.SubItems(10) = FormatNumber(nFracao, 2)
        itmX.SubItems(11) = FormatNumber(nAliq * 100, 2) & "%"
        itmX.SubItems(12) = FormatNumber(nFCat, 2)
        itmX.SubItems(13) = FormatNumber(nFPed, 2)
        itmX.SubItems(14) = FormatNumber(nFSit, 2)
        itmX.SubItems(15) = FormatNumber(nFPro, 2)
        itmX.SubItems(16) = FormatNumber(nFTop, 2)
        itmX.SubItems(17) = FormatNumber(nFGle, 2)
        itmX.SubItems(18) = FormatNumber(nValorAgrupa1, 2)
        itmX.SubItems(19) = FormatNumber(nVVT, 2)
        itmX.SubItems(20) = FormatNumber(nVVC, 2)
        itmX.SubItems(21) = FormatNumber(nVVI, 2)
        itmX.SubItems(22) = FormatNumber(nValorIptu, 2)
        itmX.SubItems(23) = FormatNumber(nValorAgrupa2, 2)
        itmX.SubItems(24) = FormatNumber(nVVT2, 2)
        itmX.SubItems(25) = FormatNumber(nVVC2, 2)
        itmX.SubItems(26) = FormatNumber(nVVI2, 2)
        itmX.SubItems(27) = FormatNumber(nValorIPTU2, 2)
        
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
End With

lblQtdeImovel.Caption = nTot
lblIptuAtual.Caption = FormatNumber(nSoma1, 2)
lblIptu2.Caption = FormatNumber(nSoma2, 2)

Liberado
'lblIPTUNovo.Caption = FormatNumber(nSomaIPTUNovo, 2)
Pb.value = 0

End Sub

Private Sub Simulado()
Dim Sql As String, qd As New rdoQuery, RdoAux As rdoResultset, x As Integer, nCodReduz As Long, RdoImovel As rdoResultset, nPos As Long, nTot As Long
Dim nAreaT As Double, nAreaC As Double, nVVT1 As Double, nVVC1 As Double, nVVI1 As Double, nVVT2 As Double, nVVC2 As Double, nVVI2 As Double, nRedutor As Double
Dim nAgrup1 As Double, nFatores1 As Double, nFracao1 As Double, nIPTU1 As Double, nAgrup2 As Double, nFatores2 As Double, nFracao2 As Double, nIPTU2 As Double
Dim nCodAgrup1 As Integer, nCodAgrup2 As Integer, sNatureza As String, nOcupacao As Double

If MsgBox("Deseja realizar o simulado de cálculo?", vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

Sql = "truncate table simulado2017"
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn
qd.QueryTimeout = 0

Sql = "select codreduzido from cadimob where inativo=0 order by codreduzido"
'Sql = "select codreduzido from cadimob where  codreduzido=36423 and  inativo=0 order by codreduzido"
Set RdoImovel = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nTot = RdoImovel.RowCount
Do Until RdoImovel.EOF
    
    If nPos Mod 100 = 0 Then
        DoEvents
        CallPb nPos, nTot
    End If
   
    nCodReduz = RdoImovel!CODREDUZIDO
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    
   ' If nCodReduz = 196 Then
   '     MsgBox "teste"
   ' End If
    
    qd.Sql = "{ Call spCALCULO(?,?) }"
    qd(0) = nCodReduz
    qd(1) = 2017
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        If !tipoisencao > 0 Then
            RdoAux.Close
            GoTo proximo
        End If
        nAreaC = !areapredial
        nAreaT = !AreaTerreno
        nVVT1 = !vvt
        nVVC1 = !VVP
        nVVI1 = !VVI
        nIPTU1 = !valorfinal
        nFCat = !fcat
        nFPed = !fped
        nFSit = !fsit
        nFPro = !fpro
        nFGle = !fgle
        nFatores1 = nFPed * nFSit * nFPro * nFGle
        nCodAgrup1 = !Agrupamento
        nAgrup1 = !valorAgrupamento
        nFracao1 = !FRACAO
        sNatureza = !Natureza
    End With
    DoEvents
    RdoAux.Close
    qd.Sql = "{ Call spCALCULOSIMULADO(?,?) }"
    qd(0) = nCodReduz
    qd(1) = 2017
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        If IsNull(!vvt) Then GoTo proximo
        nVVT2 = !vvt
        nVVC2 = !VVP
        nVVI2 = !VVI
        nIPTU2 = !valorfinal
        nFCat = !fcat
        nFPed = !fped
        nFSit = !fsit
        nFPro = !fpro
        nFGle = !fgle
        nFatores2 = nFPed * nFSit * nFPro * nFGle
        nCodAgrup2 = !Agrupamento
        nAgrup2 = !valorAgrupamento
        nFracao2 = !FRACAO
        nRedutor = !Redutor
        nOcupacao = !fatorocupacao
    End With
    nDif = (nIPTU2 - nIPTU1) / nIPTU1 * 100
On Error GoTo 0
    Sql = "insert simulado2017(codreduzido,areapredial,areaterreno,natureza,vvt1,vvp1,vvi1,agrup1,codagrup1,fatores1,fracao1,iptu1,vvt2,vvp2,vvi2,agrup2,codagrup2,fatores2,fracao2,iptu2,redutor,ocupacao,dif) "
    Sql = Sql & " values(" & nCodReduz & "," & Virg2Ponto(CStr(nAreaC)) & "," & Virg2Ponto(CStr(nAreaT)) & ",'" & sNatureza & "'," & Virg2Ponto(CStr(nVVT1)) & "," & Virg2Ponto(CStr(nVVC1)) & "," & Virg2Ponto(CStr(nVVI1)) & ","
    Sql = Sql & Virg2Ponto(CStr(nAgrup1)) & "," & nCodAgrup1 & "," & Virg2Ponto(CStr(nFatores1)) & "," & Virg2Ponto(CStr(nFracao1)) & "," & Virg2Ponto(CStr(nIPTU1)) & "," & Virg2Ponto(CStr(nVVT2)) & ","
    Sql = Sql & Virg2Ponto(CStr(nVVC2)) & "," & Virg2Ponto(CStr(nVVI2)) & "," & Virg2Ponto(CStr(nAgrup2)) & "," & nCodAgrup2 & "," & Virg2Ponto(CStr(nFatores2)) & "," & Virg2Ponto(CStr(nFracao2)) & "," & Virg2Ponto(CStr(nIPTU2)) & ","
    Sql = Sql & Virg2Ponto(CStr(nRedutor)) & "," & Virg2Ponto(CStr(nOcupacao)) & "," & Virg2Ponto(CStr(nDif)) & ")"
    cn.Execute Sql, rdExecDirect
proximo:
    nPos = nPos + 1
    RdoImovel.MoveNext
Loop
RdoImovel.Close
 MsgBox "fim"
End Sub
