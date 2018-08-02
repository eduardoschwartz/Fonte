VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDesmembramento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desmembramento de Lote"
   ClientHeight    =   6165
   ClientLeft      =   2250
   ClientTop       =   2535
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   11565
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   10170
      TabIndex        =   65
      ToolTipText     =   "Sair da Tela"
      Top             =   4905
      Width           =   1170
      _ExtentX        =   2064
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
      MICON           =   "frmDesmembramento.frx":0000
      PICN            =   "frmDesmembramento.frx":001C
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
      Left            =   10170
      TabIndex        =   66
      ToolTipText     =   "Gravar o Registro"
      Top             =   4170
      Width           =   1170
      _ExtentX        =   2064
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDesmembramento.frx":008A
      PICN            =   "frmDesmembramento.frx":00A6
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
      Left            =   10170
      TabIndex        =   67
      ToolTipText     =   "Cancelar o Desmembramento deste imóvel."
      Top             =   4545
      Width           =   1170
      _ExtentX        =   2064
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDesmembramento.frx":044B
      PICN            =   "frmDesmembramento.frx":0467
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame PnWait 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   3540
      TabIndex        =   63
      Top             =   2790
      Visible         =   0   'False
      Width           =   3255
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AGUARDE ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   60
         TabIndex        =   64
         Top             =   60
         Width           =   3225
      End
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1785
      MaxLength       =   6
      TabIndex        =   57
      Top             =   105
      Width           =   1110
   End
   Begin VB.TextBox txtQtdeImovel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4395
      MaxLength       =   3
      TabIndex        =   56
      Top             =   105
      Width           =   765
   End
   Begin VB.ComboBox cmbNumTmp 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6825
      Style           =   2  'Dropdown List
      TabIndex        =   55
      Top             =   90
      Width           =   1380
   End
   Begin VB.Frame FraD 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Áreas Construidas"
      ForeColor       =   &H00000080&
      Height          =   2385
      Index           =   3
      Left            =   4620
      TabIndex        =   42
      Top             =   3735
      Width           =   5400
      Begin prjChameleon.chameleonButton cmdDelArea 
         Height          =   315
         Left            =   1335
         TabIndex        =   43
         ToolTipText     =   "Remover uma Área"
         Top             =   1965
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Remover"
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
         MICON           =   "frmDesmembramento.frx":05C1
         PICN            =   "frmDesmembramento.frx":05DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAddArea 
         Height          =   315
         Left            =   150
         TabIndex        =   44
         ToolTipText     =   "Adicionar uma Área"
         Top             =   1965
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Adicionar"
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
         MICON           =   "frmDesmembramento.frx":0737
         PICN            =   "frmDesmembramento.frx":0753
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdEditArea 
         Height          =   315
         Left            =   2520
         TabIndex        =   54
         ToolTipText     =   "Editar uma Área"
         Top             =   1965
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Editar"
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
         MICON           =   "frmDesmembramento.frx":08AD
         PICN            =   "frmDesmembramento.frx":08C9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvArea 
         Height          =   1605
         Left            =   60
         TabIndex        =   70
         Top             =   270
         Width           =   5265
         _ExtentX        =   9287
         _ExtentY        =   2831
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Seq"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Área"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Data"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "CodUso"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Uso"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CodTipo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CodCat"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Categoria"
            Object.Width           =   2471
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Pav."
            Object.Width           =   1058
         EndProperty
      End
      Begin VB.Label lblQtdeEdif 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   5025
         TabIndex        =   46
         Top             =   2055
         Width           =   285
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de Edific..:"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   3810
         TabIndex        =   45
         Top             =   2040
         Width           =   1140
      End
   End
   Begin VB.Frame FraD 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Proprietários"
      ForeColor       =   &H00000080&
      Height          =   2385
      Index           =   0
      Left            =   15
      TabIndex        =   38
      Top             =   3735
      Width           =   4605
      Begin MSComctlLib.TreeView tvProp 
         Height          =   1650
         Left            =   60
         TabIndex        =   39
         Top             =   240
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   2910
         _Version        =   393217
         Indentation     =   794
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "ilsIcons"
         Appearance      =   1
      End
      Begin prjChameleon.chameleonButton cmdAddCid 
         Height          =   315
         Left            =   150
         TabIndex        =   40
         ToolTipText     =   "Adicionar Proprietário/Compromissário"
         Top             =   1965
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Adicionar"
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
         MICON           =   "frmDesmembramento.frx":0A23
         PICN            =   "frmDesmembramento.frx":0A3F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelCid 
         Height          =   315
         Left            =   1350
         TabIndex        =   41
         ToolTipText     =   "Remover Proprietário/Compromissário"
         Top             =   1965
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Remover"
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
         MICON           =   "frmDesmembramento.frx":0B99
         PICN            =   "frmDesmembramento.frx":0BB5
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
   Begin VB.Frame FraD 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Localização"
      ForeColor       =   &H00000080&
      Height          =   2805
      Index           =   1
      Left            =   15
      TabIndex        =   30
      Top             =   915
      Width           =   5550
      Begin VB.TextBox txtLote 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   330
         Width           =   855
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   0
         Top             =   330
         Width           =   885
      End
      Begin VB.TextBox txtNumImovel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   3
         Top             =   1665
         Width           =   885
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         MaxLength       =   40
         TabIndex        =   4
         Top             =   1995
         Width           =   4005
      End
      Begin VB.TextBox txtQuadras 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2340
         Width           =   1245
      End
      Begin VB.TextBox txtLotes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4140
         MaxLength       =   25
         TabIndex        =   6
         Top             =   2340
         Width           =   1245
      End
      Begin VB.TextBox txtNumFace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1380
         TabIndex        =   2
         Top             =   660
         Width           =   885
      End
      Begin VB.Label lblCEP 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3045
         TabIndex        =   53
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP...:"
         Height          =   225
         Index           =   8
         Left            =   2460
         TabIndex        =   52
         Top             =   1710
         Width           =   570
      End
      Begin VB.Label lblNomeLogr 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1380
         TabIndex        =   51
         Top             =   1380
         Width           =   4080
      End
      Begin VB.Label lblCodLogr 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1380
         TabIndex        =   50
         Top             =   1065
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Lote....:"
         Height          =   225
         Index           =   3
         Left            =   2490
         TabIndex        =   48
         Top             =   375
         Width           =   1065
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da Quadra....:"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   47
         Top             =   375
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Lograd.....:"
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   37
         Top             =   1380
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cód.Logradouro.:"
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   36
         Top             =   1050
         Width           =   1275
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote Origem.....:"
         Height          =   225
         Left            =   2940
         TabIndex        =   35
         Top             =   2370
         Width           =   1125
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra Origem..:"
         Height          =   225
         Left            =   90
         TabIndex        =   34
         Top             =   2370
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.....:"
         Height          =   225
         Left            =   90
         TabIndex        =   33
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Número..............:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   32
         Top             =   1710
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da Face........:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   31
         Top             =   690
         Width           =   1215
      End
   End
   Begin VB.Frame FraD 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Dados do Terreno"
      ForeColor       =   &H00000080&
      Height          =   2805
      Index           =   2
      Left            =   5580
      TabIndex        =   19
      Top             =   915
      Width           =   5925
      Begin VB.ComboBox cmbUso 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2400
         Width           =   1905
      End
      Begin VB.ComboBox cmbPedol 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2040
         Width           =   1905
      End
      Begin VB.TextBox txtFracao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2580
         TabIndex        =   8
         Text            =   "0,00"
         Top             =   270
         Width           =   885
      End
      Begin VB.Frame FraD 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Testadas"
         ForeColor       =   &H00000080&
         Height          =   2505
         Index           =   4
         Left            =   3600
         TabIndex        =   20
         Top             =   150
         Width           =   2265
         Begin VB.TextBox txtFace 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   660
            TabIndex        =   15
            Text            =   "0"
            Top             =   1710
            Width           =   435
         End
         Begin VB.TextBox txtTestada 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   660
            TabIndex        =   16
            Text            =   "0,00"
            Top             =   2040
            Width           =   885
         End
         Begin MSFlexGridLib.MSFlexGrid grdTestada 
            Height          =   1275
            Left            =   90
            TabIndex        =   21
            Top             =   300
            Width           =   2085
            _ExtentX        =   3678
            _ExtentY        =   2249
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            BackColorBkg    =   15658734
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
            FormatString    =   "^Face        |^Área  m²          "
         End
         Begin prjChameleon.chameleonButton cmdAddTestada 
            Height          =   285
            Left            =   1470
            TabIndex        =   17
            ToolTipText     =   "Adicionar Testada"
            Top             =   1680
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDesmembramento.frx":0D0F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelTestada 
            Height          =   285
            Left            =   1800
            TabIndex        =   18
            ToolTipText     =   "Remover Testada"
            Top             =   1680
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "-"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDesmembramento.frx":0D2B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "metros:"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   23
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Face:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   1740
            Width           =   495
         End
      End
      Begin VB.ComboBox cmbTopog 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   1905
      End
      Begin VB.ComboBox cmbSit 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1680
         Width           =   1905
      End
      Begin VB.ComboBox cmbCatProp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Width           =   1905
      End
      Begin VB.TextBox txtAreaTerreno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   7
         Text            =   "0,00"
         Top             =   285
         Width           =   885
      End
      Begin VB.ComboBox cmbBenf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1590
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Uso Terreno..........:"
         Height          =   225
         Left            =   90
         TabIndex        =   69
         Top             =   2490
         Width           =   1425
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedologia..............:"
         Height          =   225
         Left            =   90
         TabIndex        =   49
         Top             =   2130
         Width           =   1425
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Fr.Ideal.:"
         Height          =   225
         Left            =   1920
         TabIndex        =   29
         Top             =   330
         Width           =   675
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Topografia.............:"
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   1035
         Width           =   1425
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação................:"
         Height          =   225
         Left            =   90
         TabIndex        =   27
         Top             =   1770
         Width           =   1425
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Área Ter...:"
         Height          =   225
         Left            =   90
         TabIndex        =   26
         Top             =   345
         Width           =   825
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Benfeitoria.............:"
         Height          =   225
         Left            =   90
         TabIndex        =   25
         Top             =   660
         Width           =   1425
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Categ. Propriedade:"
         Height          =   225
         Left            =   90
         TabIndex        =   24
         Top             =   1395
         Width           =   1425
      End
   End
   Begin MSComctlLib.ImageList ImlTv 
      Left            =   1590
      Top             =   6300
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":0D47
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":0EA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":0FFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":115F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":12BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":141B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":1577
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":16D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDesmembramento.frx":19EF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label pnImovel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informações do Imóvel nº 001"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   60
      TabIndex        =   68
      Top             =   540
      Width           =   11445
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Imóvel Desmembrado.:"
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   62
      Top             =   150
      Width           =   1635
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde de Imóveis..:"
      Height          =   225
      Index           =   5
      Left            =   3045
      TabIndex        =   61
      Top             =   150
      Width           =   1380
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Imóvel Temporário..:"
      Height          =   225
      Index           =   6
      Left            =   5310
      TabIndex        =   60
      Top             =   150
      Width           =   1500
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Inscrição.:"
      Height          =   225
      Index           =   7
      Left            =   8340
      TabIndex        =   59
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lblIC 
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   9150
      TabIndex        =   58
      Top             =   150
      Width           =   2205
   End
End
Attribute VB_Name = "frmDesmembramento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, bImovelEmAndamento As Boolean, nQtdeLote As Integer
Dim nQtdeImovel As Integer, nCodNew As Integer, nCodOld As Integer, bExists As Boolean, nAreaTerreno As Double
Dim sQuadras As String

Private Sub Disable()

For Each Ct In frmDesmembramento
    If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
      Ct.BackColor = Ct.Container.BackColor
      Ct.Locked = True
    End If
Next

txtCod.BackColor = Branco
txtCod.Locked = False
txtQtdeImovel.BackColor = Branco
txtQtdeImovel.Locked = False
cmbNumTmp.BackColor = Branco
cmbNumTmp.Locked = False

pnImovel.Caption = ""

End Sub

Private Sub Enable()

For Each Ct In frmDesmembramento
    If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
      Ct.BackColor = Branco
      Ct.Locked = False
    End If
Next

For x = 0 To 4: FraD(x).Enabled = True: Next

End Sub

Private Sub cmbNumTmp_Click()
If cmbNumTmp.ListIndex = -1 Then Exit Sub
nCodOld = nCodNew
nCodNew = Val(cmbNumTmp.Text)

If nCodOld > 0 Then GravaTmp

Limpa
If cmbNumTmp.ListIndex = 0 Then
    pnImovel.Caption = "Informações do Imóvel Principal"
    Disable
    CarregaImovelPrincipal
Else
    pnImovel.Caption = "Informações do Imóvel Temporário nº " & cmbNumTmp.Text
    CarregaImovelTmp
    Enable
    If cmbNumTmp.ListIndex = 1 Then
       txtQuadra.Text = Mid$(lblIC.Caption, 6, 4)
       txtQuadra.BackColor = Kde
       txtQuadra.Locked = True
       txtLote.Text = Mid$(lblIC.Caption, 11, 5)
       txtLote.BackColor = Kde
       txtLote.Locked = True
       txtNumFace.Text = Right$(lblIC.Caption, 2)
       txtNumFace.BackColor = Kde
       txtNumFace.Locked = True
       txtNumFace_LostFocus
    End If
End If

If txtQuadra.Text = "" Then
   txtQuadra.Text = Mid$(lblIC.Caption, 6, 4)
End If
If txtQuadras.Text = "" Then
   txtQuadras.Text = sQuadras
End If


End Sub

Private Sub GravaTmp()
Dim nCodReduz As Long, nPrinc As Integer
Dim nCodBen As Integer, nCodCat As Integer, nCodSit As Integer, nCodPed As Integer, nCodTop As Integer, nCodUso As Integer
Dim nSeq As Integer, sTipo As String, nArea As Double, nUso As Integer, nTipo As Integer
Dim nCat As Integer, dDataAprova As Date, sNumProc As String, dDataProc As Date, nQtdePav As Integer

Ocupado
PnWait.Visible = True
If cGetInputState() <> 0 Then DoEvents

If cmbBenf.ListIndex = -1 Then
    nCodBen = 0
Else
    nCodBen = cmbBenf.ItemData(cmbBenf.ListIndex)
End If
If cmbUso.ListIndex = -1 Then
    nCodUso = 0
Else
    nCodUso = cmbBenf.ItemData(cmbUso.ListIndex)
End If

If cmbCatProp.ListIndex = -1 Then
    nCodCat = 0
Else
    nCodCat = cmbCatProp.ItemData(cmbCatProp.ListIndex)
End If
If cmbSit.ListIndex = -1 Then
    nCodSit = 0
Else
    nCodSit = cmbSit.ItemData(cmbSit.ListIndex)
End If
If cmbPedol.ListIndex = -1 Then
    nCodPed = 0
Else
    nCodPed = cmbPedol.ItemData(cmbPedol.ListIndex)
End If
If cmbTopog.ListIndex = -1 Then
    nCodTop = 0
Else
    nCodTop = cmbTopog.ItemData(cmbTopog.ListIndex)
End If
If txtFracao.Text = "" Then txtFracao.Text = 0
If txtAreaTerreno.Text = "" Then txtAreaTerreno.Text = 0

nCodReduz = Val(txtCod.Text)

'TABELA DESMTEMP
Sql = "SELECT * FROM DESMTEMP WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodOld
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "UPDATE DESMTEMP SET QUADRA=" & Val(txtQuadra.Text) & " ,LOTE=" & Val(txtLote.Text) & " ,FACE=" & Val(txtNumFace.Text)
        Sql = Sql & " ,NUMIMOVEL=" & Val(txtNumImovel.Text) & ",COMPLEMENTO='" & Mask(txtCompl.Text) & "',QUADRAS='" & Mask(txtQuadras.Text)
        Sql = Sql & "' ,LOTES='" & Mask(txtLotes.Text) & "',AREATERRENO=" & Virg2Ponto(RemovePonto(txtAreaTerreno.Text)) & ",CODBEN=" & nCodBen
        Sql = Sql & " ,CODCAT=" & nCodCat & " ,CODUSO=" & nCodUso & " ,CODSIT=" & nCodSit & " ,CODPED=" & nCodPed & " ,CODTOP=" & nCodTop & " ,FRACAO=" & Virg2Ponto(txtFracao.Text) & " ,QTDEEDIF=" & Val(lblQtdeEdif.Caption)
        Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodOld
    Else
        Sql = "INSERT DESMTEMP(CODREDUZIDO,CODTMP,QUADRA,LOTE,FACE,NUMIMOVEL,COMPLEMENTO,QUADRAS,LOTES,AREATERRENO,"
        Sql = Sql & "CODBEN,CODCAT,CODSIT,CODPED,CODTOP,CODUSO,FRACAO,QTDEEDIF) VALUES(" & nCodReduz & "," & nCodOld & ","
        Sql = Sql & Val(txtQuadra.Text) & "," & Val(txtLote.Text) & "," & Val(txtNumFace.Text) & "," & Val(txtNumImovel.Text) & ",'"
        Sql = Sql & Mask(txtCompl.Text) & "','" & Mask(txtQuadras.Text) & "','" & Mask(txtLotes.Text) & "',"
        Sql = Sql & Virg2Ponto(RemovePonto(txtAreaTerreno.Text)) & "," & nCodBen & "," & nCodCat & "," & nCodSit & ","
        Sql = Sql & nCodPed & "," & nCodTop & "," & nCodUso & "," & Virg2Ponto(txtFracao.Text) & "," & Val(lblQtdeEdif.Caption) & ")"
    End If
    cn.Execute Sql, rdExecDirect
   .Close
End With

'TABELA DESMTESTADA
Sql = "DELETE FROM DESMTESTADA WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodOld
cn.Execute Sql, rdExecDirect

With grdTestada
    For x = 1 To .Rows - 1
        Sql = "INSERT DESMTESTADA(CODREDUZIDO,CODTMP,NUMFACE,AREATESTADA) VALUES(" & nCodReduz & ","
        Sql = Sql & nCodOld & "," & .TextMatrix(x, 0) & "," & Virg2Ponto(.TextMatrix(x, 1)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

If tvProp.Nodes.Count > 2 Then
    'TABELA DESMPROPRIETARIO
    Sql = "DELETE FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodOld
    cn.Execute Sql, rdExecDirect
    
    For x = 1 To tvProp.Nodes.Count
        If Len(tvProp.Nodes(x).Key) > 4 Then
            If tvProp.Nodes("PROP").Children = 0 Then
                nPrinc = 0
            Else
                If tvProp.Nodes("PROP").Child.Text = tvProp.Nodes(x).Text Then
                   nPrinc = 1
                Else
                   nPrinc = 0
                End If
            End If
            Sql = "INSERT DESMPROPRIETARIO (CODREDUZIDO,CODTMP,CODCIDADAO,TIPOPROP,PRINCIPAL) VALUES("
            Sql = Sql & nCodReduz & "," & nCodOld & "," & Val(Right$(tvProp.Nodes(x).Key, 6)) & ",'"
            Sql = Sql & Left$(tvProp.Nodes(x).Key, 1) & "'," & nPrinc & ")"
            cn.Execute Sql, rdExecDirect
        End If
    Next
End If

'TABELA DESMAREAS
Sql = "DELETE FROM DESMAREAS WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodOld
cn.Execute Sql, rdExecDirect
For x = 1 To lvArea.ListItems.Count
    Sql = "INSERT DESMAREAS (CODREDUZIDO,CODTMP,SEQAREA,TIPOAREA,DATAAPROVA,AREACONSTR,USOCONSTR,TIPOCONSTR,"
    Sql = Sql & "CATCONSTR,QTDEPAV) VALUES(" & nCodReduz & "," & nCodOld & "," & x & ",'" & "'," & IIf(IsDate(lvArea.ListItems(x).SubItems(2)), "'" & Format(lvArea.ListItems(x).SubItems(2), "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & Virg2Ponto(Left(lvArea.ListItems(x).SubItems(1), Len(lvArea.ListItems(x).SubItems(1)) - 3)) & "," & lvArea.ListItems(x).SubItems(3) & "," & lvArea.ListItems(x).SubItems(5) & ","
    Sql = Sql & lvArea.ListItems(x).SubItems(7) & "," & lvArea.ListItems(x).SubItems(9) & ")"
    cn.Execute Sql, rdExecDirect
Next

PnWait.Visible = False
Liberado

End Sub

Private Sub CarregaImovelTmp()
Dim nCodReduz As Long, nCodTmp As Integer, itmX As ListItem, z As Long
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)

nCodReduz = Val(txtCod.Text)
nCodTmp = Val(cmbNumTmp.Text)

Sql = "SELECT * FROM DESMTEMP WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtQuadra.Text = !Quadra
        txtLote.Text = !Lote
        txtNumFace.Text = !FACE
        txtNumFace_LostFocus
        txtNumImovel.Text = !NUMIMOVEL
        lblCep.Caption = RetornaCEP(Val(lblCodLogr.Caption), Val(txtNumImovel))
        If Not IsNull(!Complemento) Then
            txtCompl.Text = !Complemento
        End If
        If Not IsNull(!Quadras) Then
            txtQuadras.Text = SubNull(!Quadras)
            If txtQuadras.Text = "" Then txtQuadras.Text = " " 'USADO PARA REPETIR AS QUADRAS
        End If
        txtLotes.Text = SubNull(!Lotes)
        txtAreaTerreno.Text = FormatNumber(!AreaTerreno, 2)
        If !CODBEN > 0 Then
            For x = 0 To cmbBenf.ListCount - 1
                cmbBenf.ListIndex = x
                If cmbBenf.ItemData(cmbBenf.ListIndex) = !CODBEN Then
                   Exit For
                End If
            Next
        End If
        If !CODTOP > 0 Then
            For x = 0 To cmbTopog.ListCount - 1
                cmbTopog.ListIndex = x
                If cmbTopog.ItemData(cmbTopog.ListIndex) = !CODTOP Then
                   Exit For
                End If
            Next
        End If
        If !CODSIT > 0 Then
            For x = 0 To cmbSit.ListCount - 1
                cmbSit.ListIndex = x
                If cmbSit.ItemData(cmbSit.ListIndex) = !CODSIT Then
                   Exit For
                End If
            Next
        End If
        If !CODCAT > 0 Then
            For x = 0 To cmbCatProp.ListCount - 1
                cmbCatProp.ListIndex = x
                If cmbCatProp.ItemData(cmbCatProp.ListIndex) = !CODCAT Then
                   Exit For
                End If
            Next
        End If
        If !CODPED > 0 Then
            For x = 0 To cmbPedol.ListCount - 1
                cmbPedol.ListIndex = x
                If cmbPedol.ItemData(cmbPedol.ListIndex) = !CODPED Then
                   Exit For
                End If
            Next
         End If
        If !coduso > 0 Then
            For x = 0 To cmbUso.ListCount - 1
                cmbUso.ListIndex = x
                If cmbUso.ItemData(cmbUso.ListIndex) = !coduso Then
                   Exit For
                End If
            Next
         End If
         txtFracao.Text = FormatNumber(!FRACAO, 2)
    Else
                      
    End If
End With

'TESTADA
Sql = "SELECT NUMFACE,AREATESTADA FROM DESMTESTADA WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdTestada.AddItem Format(!NUMFACE, "00") & Chr(9) & FormatNumber(!AREATESTADA, 2)
       .MoveNext
    Loop
   .Close
End With

'CARREGA PROPRIETARIO

Sql = "SELECT DESMPROPRIETARIO.CODCIDADAO,CIDADAO.NOMECIDADAO,"
Sql = Sql & "DESMPROPRIETARIO.TIPOPROP,DESMPROPRIETARIO.PRINCIPAL "
Sql = Sql & "FROM DESMPROPRIETARIO INNER JOIN CIDADAO ON "
Sql = Sql & "DESMPROPRIETARIO.CODCIDADAO = CIDADAO.CODCIDADAO "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
Inicio:
    If .RowCount > 0 Then
        For i = 1 To tvProp.Nodes.Count
            tvProp.Nodes.Remove (i)
            GoTo Inicio
        Next
        BuildTreeProp
    End If
    Do Until .EOF
       If !tipoprop = "P" Then
          If !principal = 0 Then
            Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(!CodCidadao, "000000"), !nomecidadao, 1)
          Else
            Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(!CodCidadao, "000000"), !nomecidadao & " - Principal", 1)
          End If
          tvProp.Nodes("PROP" & Format(!CodCidadao, "000000")).ForeColor = vbBlue
       Else
          Set NodX = tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(!CodCidadao, "000000"), !nomecidadao, 2)
          tvProp.Nodes("COMP" & Format(!CodCidadao, "000000")).ForeColor = vbBlue
       End If
      .MoveNext
    Loop
End With

For x = 1 To tvProp.Nodes.Count
    tvProp.Nodes(x).EnsureVisible
Next
tvProp.Refresh


'Areas
Sql = "SELECT DESMAREAS.SEQAREA,DESMAREAS.QTDEPAV,DESMAREAS.TIPOAREA,DESMAREAS.DATAAPROVA,DESMAREAS.AREACONSTR,DESMAREAS.NUMPROCESSO,DESMAREAS.DATAPROCESSO,"
Sql = Sql & "DESMAREAS.USOCONSTR,USOCONSTR.DESCUSOCONSTR,DESMAREAS.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
Sql = Sql & "DESMAREAS.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM DESMAREAS INNER JOIN USOCONSTR ON "
Sql = Sql & "DESMAREAS.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
Sql = Sql & "DESMAREAS.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
Sql = Sql & "DESMAREAS.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblQtdeEdif.Caption = .RowCount
    Do Until .EOF
       '****ListView
        Set itmX = lvArea.ListItems.Add(, "A" & Format(.AbsolutePosition, "00"), Format(.AbsolutePosition, "00"))
        itmX.SubItems(1) = FormatNumber(!AREACONSTR, 2) & " m²"
        itmX.SubItems(2) = IIf(IsNull(!DATAAPROVA), "", Format(!DATAAPROVA, "dd/mm/yyyy"))
        itmX.SubItems(3) = !USOCONSTR
        itmX.SubItems(4) = !descusoconstr
        itmX.SubItems(5) = !TIPOCONSTR
        itmX.SubItems(6) = !DESCTIPOCONSTR
        itmX.SubItems(7) = !CATCONSTR
        itmX.SubItems(8) = !desccategconstr
        itmX.SubItems(9) = Val(SubNull(!QTDEPAV))
      
      .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub CarregaImovelPrincipal()
Dim nCodReduz As Long
Dim qd As New rdoQuery
Dim itmX As ListItem, z As Long
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)

nCodReduz = Val(txtCod.Text)

Ocupado
PnWait.Visible = True
If cGetInputState() <> 0 Then DoEvents
Set qd.ActiveConnection = cn
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spDADOSDEUMIMOVEL(?) }"
qd(0) = nCodReduz
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
    txtQuadra.Text = !Quadra
    txtLote.Text = !Lote
    txtNumFace.Text = !Seq
    txtNumFace_LostFocus
    txtNumImovel.Text = Val(SubNull(!Li_Num))
    lblCep.Caption = RetornaCEP(Val(lblCodLogr.Caption), Val(txtNumImovel))
    txtAreaTerreno.Text = FormatNumber(!Dt_AreaTerreno, 2)
    lblCodLogr.Caption = !CodLogr
    lblNomeLogr.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
    txtCompl.Text = SubNull(!Li_Compl)
    If Not IsNull(!Li_Quadras) Then
        txtQuadras.Text = SubNull(!Li_Quadras)
    End If
    sQuadras = SubNull(!Li_Quadras)
    txtLotes.Text = SubNull(!Li_Lotes)
    For x = 0 To cmbBenf.ListCount - 1
        cmbBenf.ListIndex = x
        If cmbBenf.ItemData(cmbBenf.ListIndex) = !Dt_CodBenf Then
           Exit For
        End If
    Next
    For x = 0 To cmbTopog.ListCount - 1
        cmbTopog.ListIndex = x
        If cmbTopog.ItemData(cmbTopog.ListIndex) = !Dt_CodTopog Then
           Exit For
        End If
    Next
    For x = 0 To cmbSit.ListCount - 1
        cmbSit.ListIndex = x
        If cmbSit.ItemData(cmbSit.ListIndex) = !Dt_CodSituacao Then
           Exit For
        End If
    Next
    For x = 0 To cmbCatProp.ListCount - 1
        cmbCatProp.ListIndex = x
        If cmbCatProp.ItemData(cmbCatProp.ListIndex) = !Dt_CodCategProp Then
           Exit For
        End If
    Next
    For x = 0 To cmbPedol.ListCount - 1
        cmbPedol.ListIndex = x
        If cmbPedol.ItemData(cmbPedol.ListIndex) = !Dt_CodPedol Then
           Exit For
        End If
    Next
    For x = 0 To cmbPedol.ListCount - 1
        cmbPedol.ListIndex = x
        If cmbPedol.ItemData(cmbPedol.ListIndex) = !Dt_CodPedol Then
           Exit For
        End If
    Next
    For x = 0 To cmbUso.ListCount - 1
        cmbUso.ListIndex = x
        If cmbUso.ItemData(cmbUso.ListIndex) = !Dt_CodUsoTerreno Then
           Exit For
        End If
    Next
    txtFracao.Text = FormatNumber(!Dt_FracaoIdeal, 2)
   .Close
End With

'CARREGA TESTADA
grdTestada.Rows = 1
Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdTestada.AddItem Format(!NUMFACE, "00") & Chr(9) & FormatNumber(!AREATESTADA, 2)
       .MoveNext
    Loop
   .Close
End With

'CARREGA PROPRIETARIO
Inicio:
For i = 1 To tvProp.Nodes.Count
    tvProp.Nodes.Remove (i)
    GoTo Inicio
Next
BuildTreeProp

Sql = "SELECT PROPRIETARIO.CODCIDADAO,CIDADAO.NOMECIDADAO,PROPRIETARIO.TIPOPROP "
Sql = Sql & "FROM PROPRIETARIO INNER JOIN CIDADAO ON PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO = " & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !tipoprop = "P" Then
           Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(!CodCidadao, "000000"), !nomecidadao, 1)
           tvProp.Nodes("PROP" & Format(!CodCidadao, "000000")).ForeColor = vbBlue
        Else
           Set NodX = tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(!CodCidadao, "000000"), !nomecidadao, 2)
           tvProp.Nodes("COMP" & Format(!CodCidadao, "000000")).ForeColor = vbBlue
        End If
       .MoveNext
    Loop
   .Close
End With
For x = 1 To tvProp.Nodes.Count
    tvProp.Nodes(x).EnsureVisible
Next

'CARREGA ÁREA
       
Sql = "SELECT AREAS.SEQAREA,AREAS.TIPOAREA,AREAS.AREACONSTR,AREAS.USOCONSTR,USOCONSTR.DESCUSOCONSTR,"
Sql = Sql & "AREAS.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,AREAS.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR,"
Sql = Sql & "AREAS.DATAAPROVA, AREAS.NUMPROCESSO,Areas.DATAPROCESSO,Areas.QTDEPAV FROM AREAS INNER JOIN "
Sql = Sql & "USOCONSTR ON AREAS.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
Sql = Sql & "AREAS.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON Areas.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
Sql = Sql & "Where Areas.CODREDUZIDO = " & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       '****ListView
        Set itmX = lvArea.ListItems.Add(, "A" & Format(.AbsolutePosition, "00"), Format(.AbsolutePosition, "00"))
        itmX.SubItems(1) = FormatNumber(!AREACONSTR, 2) & " m²"
        itmX.SubItems(2) = IIf(IsNull(!DATAAPROVA), "", Format(!DATAAPROVA, "dd/mm/yyyy"))
        itmX.SubItems(3) = !USOCONSTR
        itmX.SubItems(4) = !descusoconstr
        itmX.SubItems(5) = !TIPOCONSTR
        itmX.SubItems(6) = !DESCTIPOCONSTR
        itmX.SubItems(7) = !CATCONSTR
        itmX.SubItems(8) = !desccategconstr
        itmX.SubItems(9) = Val(SubNull(!QTDEPAV))
       .MoveNext
    Loop
End With

PnWait.Visible = False
Liberado

End Sub

Private Sub cmdAddArea_Click()
If cmbNumTmp.ListIndex < 1 Then
    MsgBox "Impossivel adicionar área.", vbCritical, "Atenção"
    Exit Sub
End If

Set frm = frmAreas
frm.sForm = Me.Name
frm.sEvento = "Novo"
frm.show 1
lblQtdeEdif.Caption = lvArea.ListItems.Count

End Sub

Private Sub cmdAddCid_Click()

If cmbNumTmp.ListIndex < 1 Then
    MsgBox "Impossivel adicionar proprietário.", vbCritical, "Atenção"
    Exit Sub
End If

On Error GoTo Erro:
   Set frm = frmCnsCidadao
   frm.sForm = Me.Name
   If tvProp.Nodes("PROP").Children = 0 Then
      frm.sTipoCidadao = "P"
   Else
      frm.sTipoCidadao = Left$(tvProp.SelectedItem.Key, 1)
   End If
   frmCnsCidadao.show
   Exit Sub
   
Erro:
   MsgBox "Selecione na árvore Proprietário ou Proprietário Solidário.", vbExclamation, "Atenção"

End Sub

Private Sub cmdAddTestada_Click()

If Val(txtFace.Text) = 0 Then
   MsgBox "Digite a Face da Testada.", vbExclamation, "Atenção"
   Exit Sub
End If

If CDbl(txtTestada.Text) = 0 Then
   MsgBox "Digite a Área da Testada.", vbExclamation, "Atenção"
   Exit Sub
End If

For x = 1 To grdTestada.Rows - 1
    If Val(grdTestada.TextMatrix(x, 0)) = Val(txtFace.Text) Then
        MsgBox "Testada já incluida na lista.", vbExclamation, "Atenção"
        txtFace.SetFocus
        Exit Sub
    End If
Next

grdTestada.AddItem Format(txtFace.Text, "00") & Chr(9) & FormatNumber(txtTestada, 2)
txtFace.Text = Val(grdTestada.TextMatrix(grdTestada.Rows - 1, 0)) + 1
txtTestada.Text = "0,00"
txtFace.SetFocus

End Sub

Private Sub cmdCancel_Click()

If cmbNumTmp.ListIndex = -1 Then
    MsgBox "Nada a Cancelar.", vbCritical, "Atenção"
    Exit Sub
End If

If MsgBox("Deseja cancelar a execução do Desmembramento deste Imóvel." & vbcrl & vbCrLf & "Todas as informações previamente armazenadas serão excluidas.", vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação de Cancelamento") = vbNo Then Exit Sub

'TABELA DESMEMBRAMENTO
Sql = "DELETE FROM DESMTESTADA WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DESMAREAS WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DESMTEMP WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DESMEMBRAMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

Limpa

cmbNumTmp.Clear
txtQtdeImovel.Text = ""
txtCod.Text = ""
txtCod.SetFocus

End Sub

Private Sub cmdDelArea_Click()

If cmbNumTmp.ListIndex < 1 Then
    MsgBox "Impossivel remover área.", vbCritical, "Atenção"
    Exit Sub
End If

If lvArea.ListItems.Count = 0 Then Exit Sub

Dim x As Integer
If MsgBox("Excluir esta área ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
    lvArea.ListItems.Remove (lvArea.SelectedItem.Index)
    For x = 1 To lvArea.ListItems.Count
        lvArea.ListItems(x).Text = Format(x, "00")
    Next
    lblQtdeEdif.Caption = x
End If

End Sub

Private Sub cmdDelCid_Click()
On Error GoTo Erro

If cmbNumTmp.ListIndex < 1 Then
    MsgBox "Impossivel remover proprietário.", vbCritical, "Atenção"
    Exit Sub
End If

n = tvProp.SelectedItem.Parent.Index
nc = tvProp.SelectedItem.Index
tvProp.Nodes.Remove (nc)
If tvProp.Nodes("PROP").Children > 0 Then
    If Right$(tvProp.Nodes("PROP").Child.Text, 9) <> "Principal" Then
          tvProp.Nodes("PROP").Child.Text = tvProp.Nodes("PROP").Child.Text & " - Principal"
    End If
End If

Exit Sub
Erro:
MsgBox "Selecione o proprietário a ser removido.", vbExclamation, "Atenção"


End Sub

Private Sub cmdDelTestada_Click()

If grdTestada.Rows = 1 Then
   MsgBox "Selecione a Face a ser excluída.", vbExclamation, "Atenção"
Else
   If grdTestada.Rows > 2 Then
      grdTestada.RemoveItem (grdTestada.Row)
   Else
      grdTestada.Rows = 1
   End If
End If

End Sub

Private Sub cmdEditArea_Click()
If lvArea.ListItems.Count = 0 Then Exit Sub
Set frm = frmAreas
frm.sEvento = "Alterar"
frm.sForm = Me.Name
frm.nSequenciaArea = Val(lvArea.SelectedItem.SubItems(1))

frm.sUso = lvArea.SelectedItem.SubItems(3)
frm.sTipo = lvArea.SelectedItem.SubItems(5)
frm.sCat = lvArea.SelectedItem.SubItems(7)
frm.nQtdePavimento = lvArea.SelectedItem.SubItems(9)
frm.dDataConstrucao = lvArea.SelectedItem.SubItems(2)
frm.nAreaConstrucao = CDbl(Left(lvArea.SelectedItem.SubItems(1), Len(lvArea.SelectedItem.SubItems(1)) - 3))

frm.show 1

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, bValido As Boolean

If cmbNumTmp.ListIndex = -1 Then
    MsgBox "Nada a Gravar.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtQtdeImovel) < 2 Then
    MsgBox "Desmembramento exige no mínimo 2 lotes", vbCritical, "ERRO"
    Exit Sub
End If

If Val(cmbNumTmp.Text) > 0 Then
    nCodOld = Val(cmbNumTmp.Text)
    GravaTmp
End If

Sql = "SELECT SUM(AREATERRENO) AS SOMA FROM DESMTEMP WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If Not IsNull(!soma) Then
        If nAreaTerreno <> FormatNumber(!soma, 2) Then
            MsgBox "A soma das áreas do terreno (" & FormatNumber(!soma, 2) & ") é diferente da área do terreno principal (" & FormatNumber(nAreaTerreno, 2) & ")", vbExclamation, "Atenção"
            Exit Sub
        End If
    End If
   .Close
End With

PnWait.Visible = True
If cGetInputState() <> 0 Then DoEvents
Ocupado

bValido = True
For x = 1 To nQtdeImovel
    If Not ValidaLote(x) Then
        bValido = False
        Exit For
    End If
Next

If bValido Then
   Grava
End If

Liberado
PnWait.Visible = False

End Sub

Private Function ValidaLote(nCodTmp As Integer) As Boolean
Dim nCodReduz As Long, nDist As Integer, nSetor As Integer, nQuadra As Integer, nLote As Integer, nFace As Integer
nCodReduz = Val(txtCod.Text)
ValidaLote = False

Sql = "SELECT * FROM DESMTEMP WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Digite os dados para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
        Exit Function
    Else
        nDist = Val(Left$(lblIC.Caption, 1))
        nSetor = Val(Mid$(lblIC.Caption, 3, 2))
        nQuadra = !Quadra
        nLote = !Lote
        nFace = !FACE
        If nQuadra = 0 Or nLote = 0 Or nFace = 0 Then
            MsgBox "Digite a Quadra, Lote e Face para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
            Exit Function
        End If
        If !AreaTerreno = 0 Then
            MsgBox "Digite a área do terreno para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
            Exit Function
        End If
        If !CODBEN = 0 Or !CODCAT = 0 Or !CODSIT = 0 Or !CODPED = 0 Or !CODTOP = 0 Then
            MsgBox "Digite os dados do terreno para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
            Exit Function
        End If
    End If
   .Close
End With

If nCodTmp > 1 Then
    Sql = "SELECT CODREDUZIDO,DV FROM CADIMOB WHERE DISTRITO=" & nDist & " AND SETOR=" & nSetor & " AND "
    Sql = Sql & "QUADRA=" & nQuadra & " AND LOTE=" & nLote
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Já existe um imóvel (" & Format(!CODREDUZIDO, "0000000") & "-" & !DV & ") com o nº de quadra/lote do Lote Temporário nº" & nCodTmp & vbCrLf & "Deseja continuar ?", vbQuestion + vbYesNo, "Atenção") = vbNo Then
               Exit Function
            End If
        End If
       .Close
    End With
End If

If nCodTmp > 1 Then
    Sql = "SELECT CODREDUZIDO,CODTMP FROM DESMTEMP WHERE CODREDUZIDO=" & nCodReduz & " AND QUADRA=" & nQuadra & " AND LOTE=" & nLote & " AND CODTMP <> " & nCodTmp
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Lotes com quadra/lote duplicados.", vbExclamation, "Atenção"
            Exit Function
        End If
       .Close
    End With
End If

Sql = "SELECT * FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Selecione o proprietário para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
        Exit Function
    End If
   .Close
End With

Sql = "SELECT * FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp & " AND TIPOPROP='P'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Selecione o proprietário principal para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
        Exit Function
    End If
   .Close
End With

Sql = "SELECT * FROM DESMTESTADA WHERE CODREDUZIDO=" & nCodReduz & " AND CODTMP=" & nCodTmp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Selecione a(s) testada(s) para o Lote Temporário nº " & nCodTmp, vbExclamation, "Atenção"
        Exit Function
    End If
   .Close
End With

ValidaLote = True

End Function

Private Sub Grava()
Dim x As Integer, aNovosCodigos() As Long, nNovoCod As Long, s As String, nBairro As Integer, sHist As String
Dim nDist As Integer, nSetor As Integer, nUso As Integer
Dim nSeq As Integer, RdoAux2 As rdoResultset, bSucesso As Boolean

On Error GoTo Erro

ReDim aNovosCodigos(nQtdeImovel)

bSucesso = True
nDist = Left$(lblIC.Caption, 1)
nSetor = Mid$(lblIC.Caption, 3, 2)

'BUSCA O BAIRRO PRINCIPAL
 Sql = "SELECT LI_CODBAIRRO,DT_CODUSOTERRENO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCod.Text)
 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
 With RdoAux
     nBairro = !Li_CodBairro
     nUso = !Dt_CodUsoTerreno
    .Close
 End With

For x = 1 To nQtdeImovel
   'BUSCA O PROXIMO CÓDIGO DE IMÓVEL
    Sql = "SELECT MAX(CODREDUZIDO) AS MAXIMO FROM CADIMOB WHERE CODREDUZIDO<40000"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nNovoCod = !maximo + 1
        aNovosCodigos(x) = nNovoCod
       .Close
    End With
    
    Sql = "delete from debitoparcela where codreduzido=" & nNovoCod
    cn.Execute Sql, rdExecDirect
    Sql = "delete from debitotributo where codreduzido=" & nNovoCod
    cn.Execute Sql, rdExecDirect
    Sql = "delete from parceladocumento where codreduzido=" & nNovoCod
    cn.Execute Sql, rdExecDirect
    
   'CARREGA TABELA DESMTEMP
    Sql = "SELECT * FROM DESMTEMP WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP=" & x
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
       'GRAVA NA TABELA CADIMOB
        Sql = "INSERT CADIMOB (CODREDUZIDO,DV,CODCONDOMINIO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,"
        Sql = Sql & "LI_NUM,LI_COMPL,LI_UF,LI_CODCIDADE,LI_CODBAIRRO,LI_QUADRAS,LI_LOTES,DT_AREATERRENO,"
        Sql = Sql & "DT_CODUSOTERRENO,DT_CODBENF,DT_CODTOPOG,DT_CODCATEGPROP,DT_CODSITUACAO,DT_CODPEDOL,"
        Sql = Sql & "DT_FRACAOIDEAL,DC_QTDEEDIF,EE_TIPOEND,INATIVO,resideimovel) VALUES(" & nNovoCod & "," & RetornaDVCodReduzido(nNovoCod) & ","
        Sql = Sql & 999 & "," & nDist & "," & nSetor & "," & !Quadra & "," & !Lote & "," & !FACE & "," & 0 & "," & 0 & ","
        Sql = Sql & !NUMIMOVEL & ",'" & Mask(!Complemento) & "','" & "SP" & "'," & 413 & "," & nBairro & ",'" & Mask(!Quadras) & "','"
        Sql = Sql & Mask(!Lotes) & "'," & Virg2Ponto(!AreaTerreno) & "," & !coduso & "," & !CODBEN & "," & !CODTOP & "," & !CODCAT & ","
        Sql = Sql & !CODSIT & "," & !CODPED & "," & Virg2Ponto(!FRACAO) & "," & !QTDEEDIF & "," & 0 & "," & 0 & "," & 1 & ")"
        cn.Execute Sql, rdExecDirect
       .Close
    End With
   'CARREGA TABELA DESMPROPRIETARIO
    Sql = "SELECT * FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP=" & x
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
           'GRAVA NA TABELA PROPRIETARIO
            Sql = "INSERT PROPRIETARIO (CODREDUZIDO,CODCIDADAO,TIPOPROP,PRINCIPAL) VALUES(" & nNovoCod & ","
            Sql = Sql & !CodCidadao & ",'" & !tipoprop & "'," & IIf(!principal, 1, 0) & ")"
            cn.Execute Sql, rdExecDirect
            If !tipoprop = "P" And !principal Then
                AtualizaPropDuplicado nNovoCod, !CodCidadao
            End If
           .MoveNext
        Loop
       .Close
    End With
   'CARREGA TABELA DESMTESTADA
    Sql = "SELECT * FROM DESMTESTADA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP=" & x
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
           'GRAVA NA TABELA TESTADA
            Sql = "INSERT TESTADA (CODREDUZIDO,NUMFACE,AREATESTADA) VALUES(" & nNovoCod & ","
            Sql = Sql & !NUMFACE & "," & Virg2Ponto(!AREATESTADA) & ")"
            cn.Execute Sql, rdExecDirect
           .MoveNext
        Loop
       .Close
    End With
   'CARREGA TABELA DESMAREAS
    Sql = "SELECT * FROM DESMAREAS WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP=" & x
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
           'GRAVA NA TABELA AREAS
            Sql = "INSERT AREAS (CODREDUZIDO,SEQAREA,TIPOAREA,AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,DATAAPROVA,NUMPROCESSO,DATAPROCESSO,QTDEPAV) VALUES(" & nNovoCod & ","
            Sql = Sql & !SEQAREA & ",'" & !TIPOAREA & "'," & Virg2Ponto(!AREACONSTR) & "," & !USOCONSTR & "," & !TIPOCONSTR & "," & !CATCONSTR & ",'"
            Sql = Sql & Format(!DATAAPROVA, "mm/dd/yyyy") & "','" & !numprocesso & "','" & Format(!DATAPROCESSO, "mm/dd/yyyy") & "'," & !QTDEPAV & ")"
            cn.Execute Sql, rdExecDirect
           .MoveNext
        Loop
       .Close
    End With
   'GRAVA NA TABELA HISTÓRICO
    sHist = "O imóvel foi criado a partir do desmembramento do lote: " & Val(txtCod.Text)
    Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST,DATAHIST2) VALUES("
    Sql = Sql & nNovoCod & "," & 1 & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sHist & "','" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
Next

'IMPORTAÇÃO DA DIVIDA DO PRINCIPAL PARA O PRIMEIRO
Sql = "INSERT DEBITOPARCELA SELECT " & aNovosCodigos(1) & ",ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,"
Sql = Sql & "VALORJUROS,NUMPROCESSO,INTACTO,NOTIFICADO,USERID,NUMEXECFISCAL,ANOEXECFISCAL From DEBITOPARCELA Where CODREDUZIDO = " & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

Sql = "INSERT DEBITOTRIBUTO SELECT " & aNovosCodigos(1) & ",ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
Sql = Sql & "VALORTRIBUTO,VALORCORRECAO,VALORMULTA,VALORJUROS,INTACTO,VALORPORBAIXA From DEBITOTRIBUTO Where CODREDUZIDO = " & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

Sql = "INSERT DEBITOPAGO SELECT " & aNovosCodigos(1) & ",ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQPAG,"
Sql = Sql & "DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,RESTITUIDO,NUMDOCUMENTO,VALORPAGOREAL,INTACTO,VALORTARIFA,ARQUIVOBANCO,VALORDIF,DATAPAGAMENTOCALC,DATAINTEGRACAO From DEBITOPAGO Where CODREDUZIDO = " & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

Sql = "SELECT * FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & aNovosCodigos(1) & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
       Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
       Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            If .RowCount > 0 Then
                Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & aNovosCodigos(1) & ","
                Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
                Sql = Sql & !CODCOMPLEMENTO & "," & RdoAux!NumDocumento & ")"
                cn.Execute Sql, rdExecDirect
                Sql = "DELETE FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                cn.Execute Sql, rdExecDirect
            End If
           .Close
       End With
      .MoveNext
    Loop
   .Close
End With


Sql = "INSERT ORIGEMREPARC SELECT NUMPROCESSO," & aNovosCodigos(1) & ",ANOEXERCICIO,CODLANCAMENTO,"
Sql = Sql & "NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,PRINCIPAL,JUROS,MULTA,CORRECAO FROM ORIGEMREPARC WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

Sql = "INSERT DESTINOREPARC SELECT NUMPROCESSO," & aNovosCodigos(1) & ",ANOEXERCICIO,CODLANCAMENTO,"
Sql = Sql & "NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,VALORLIQUIDO,JUROS,MULTA,CORRECAO,VALORPRINCIPAL,SALDO,JUROSPERC,JUROSVALOR,JUROSAPL,HONORARIO,TOTAL FROM DESTINOREPARC WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

Sql = "UPDATE PROCESSOREPARC SET CODIGORESP=" & aNovosCodigos(1) & " WHERE CODIGORESP=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

'STATUS TRANSFERIDO
Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND STATUSLANC=3"
cn.Execute Sql, rdExecDirect

'OPTANTES DA
Sql = "UPDATE DEBITOAUTOMATICO SET CODREDUZ=" & aNovosCodigos(1) & " WHERE CODREDUZ=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

'INFORMA LOTES CRIADOS
For x = 1 To UBound(aNovosCodigos)
    s = s & CStr(aNovosCodigos(x)) & ","
Next
s = Chomp(s, chomp_righT, 1)
MsgBox "Imóveis criados (" & s & ")", vbInformation, "Desmembramento com Sucesso"

'HISTORICO IMOVEL PRINCIPAL
Sql = "SELECT max(SEQ) as maximo FROM HISTORICO WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
If IsNull(RdoAux!maximo) Then
    nSeq = 1
Else
    nSeq = RdoAux!maximo + 1
End If
sHist = "O imóvel foi desmembrado e criou os imóveis: (" & s & ")"
Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST,DATAHIST2) VALUES("
Sql = Sql & Val(txtCod.Text) & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sHist & "','" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect

'INATIVA O IMOVEL PRINCIPAL
Sql = "UPDATE CADIMOB SET INATIVO=1 WHERE CODREDUZIDO=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

'EXCLUI O DESMEMBRAMENTO
If bSucesso Then
    Sql = "DELETE FROM DESMTESTADA WHERE CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM DESMAREAS WHERE CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM DESMTEMP WHERE CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM DESMEMBRAMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
    Disable
    Limpa
    txtCod.Text = ""
    txtQtdeImovel.Text = ""
    cmbNumTmp.Clear
End If

Exit Sub
Erro:
bSucesso = False
For k = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(k).Description
Next
Resume Next

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Disable
CarregaCombo
txtCod.Locked = False
txtCod.BackColor = Branco
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Val(cmbNumTmp.Text) > 0 Then
    nCodOld = Val(cmbNumTmp.Text)
    GravaTmp
End If

End Sub

Private Sub CarregaCombo()

bExec = False
Sql = "SELECT CODSITUACAO,DESCSITUACAO FROM SITUACAO WHERE CODSITUACAO<>999 ORDER BY DESCSITUACAO; " & _
      "SELECT CODBENFEITORIA,DESCBENFEITORIA FROM BENFEITORIA WHERE CODBENFEITORIA<>999 ORDER BY DESCBENFEITORIA; " & _
      "SELECT CODPEDOLOGIA,DESCPEDOLOGIA FROM PEDOLOGIA WHERE CODPEDOLOGIA<>999 ORDER BY DESCPEDOLOGIA; " & _
      "SELECT CODTOPOGRAFIA,DESCTOPOGRAFIA FROM TOPOGRAFIA WHERE CODTOPOGRAFIA<>999 ORDER BY DESCTOPOGRAFIA; " & _
      "SELECT CODUSOTERRENO,DESCUSOTERRENO FROM USOTERRENO WHERE CODUSOTERRENO<>999 ORDER BY DESCUSOTERRENO; " & _
      "SELECT CODCATEGPROP,DESCCATEGPROP FROM CATEGPROP WHERE CODCATEGPROP<>999 ORDER BY DESCCATEGPROP"

Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbSit.AddItem !DescSituacao
       cmbSit.ItemData(cmbSit.NewIndex) = !Codsituacao
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbBenf.AddItem !DescBenfeitoria
       cmbBenf.ItemData(cmbBenf.NewIndex) = !CODBENFEITORIA
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbPedol.AddItem !DescPedologia
       cmbPedol.ItemData(cmbPedol.NewIndex) = !CODPEDOLOGIA
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbTopog.AddItem !DescTopografia
       cmbTopog.ItemData(cmbTopog.NewIndex) = !CODTOPOGRAFIA
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbUso.AddItem !DescUsoTerreno
       cmbUso.ItemData(cmbUso.NewIndex) = !CODUSOTERRENO
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbCatProp.AddItem !DescCategProp
       cmbCatProp.ItemData(cmbCatProp.NewIndex) = !CODCATEGPROP
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub txtAreaTerreno_KeyPress(KeyAscii As Integer)
Tweak txtAreaTerreno, KeyAscii, DecimalPositive
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Dim RdoAux2 As rdoResultset
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    Sql = "SELECT CODREDUZIDO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,INATIVO,DT_AREATERRENO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "Código não cadastrado.", vbCritical, "Atenção"
            txtQtdeImovel.Text = ""
            txtQtdeImovel.Locked = True
            txtQtdeImovel.BackColor = txtQtdeImovel.Container.BackColor
            txtCod.SetFocus
        Else
            Sql = "SELECT * FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & Val(txtCod.Text)
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    'MsgBox "Este imóvel não pode ser desmembrado pois possue cadastro no débito automático.", vbCritical, "Atenção"
                    MsgBox "O débito automático deste imóvel sera transferido para o primeiro imóvel desmembrado.", vbCritical, "Atenção"
                    'Exit Sub
                   .Close
                End If
'                .Close
            End With
            nAreaTerreno = FormatNumber(!Dt_AreaTerreno, 2)
            lblIC.Caption = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00")
            If !Inativo Then
                MsgBox "Este imóvel encontra-se inativo.", vbExclamation, "Atenção"
                txtCod.SetFocus
                Exit Sub
                txtQtdeImovel.Text = ""
                txtQtdeImovel.Locked = True
                txtQtdeImovel.BackColor = txtQtdeImovel.Container.BackColor
                txtCod.SetFocus
                
            Else
                Limpa
                Sql = "SELECT * FROM DESMTEMP WHERE CODREDUZIDO=" & Val(txtCod.Text)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                   If .RowCount = 0 Then
                      bImovelEmAndamento = False
                      nQtdeImovel = 0
                      txtQtdeImovel.Locked = False
                      txtQtdeImovel.BackColor = Branco
                      txtQtdeImovel.SetFocus
                      'txtCod_LostFocus
                      
                   Else
                      bImovelEmAndamento = True
                      nQtdeImovel = .RowCount
                      txtQtdeImovel.Text = nQtdeImovel
                      txtQtdeImovel.Locked = True
                      txtQtdeImovel.BackColor = txtQtdeImovel.Container.BackColor
                      If cmbNumTmp.ListCount = 0 Then
                        cmbNumTmp.AddItem "Principal"
                        For x = 1 To .RowCount
                          cmbNumTmp.AddItem x
                        Next
                        cmbNumTmp.SetFocus
                     End If
                   End If
                  .Close
                End With
            End If
        End If
       .Close
    End With
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If

End Sub

Private Sub txtCod_LostFocus()
If Val(txtCod.Text) > 0 Then
   txtCod_KeyPress vbKeyReturn
Else
   Limpa
   txtQtdeImovel.Text = ""
   cmbNumTmp.Clear
   lblIC.Caption = ""
   cmdGravar.SetFocus
End If
End Sub

Private Sub txtFace_GotFocus()
txtFace.SelStart = 0
txtFace.SelLength = Len(txtFace.Text)
End Sub

Private Sub txtFracao_KeyPress(KeyAscii As Integer)
Tweak txtFracao, KeyAscii, DecimalPositive
End Sub

Private Sub txtLote_GotFocus()
txtLote.SelStart = 0
txtLote.SelLength = Len(txtLote.Text)
End Sub

Private Sub txtNumFace_GotFocus()
txtNumFace.SelStart = 0
txtNumFace.SelLength = Len(txtNumFace.Text)
End Sub

Private Sub txtNumFace_LostFocus()
If Val(txtQuadra.Text) > 0 And Val(txtNumFace.Text) > 0 Then
   If Not ValidaFace Then
      txtNumFace.SetFocus
   Else
      lblCep.Caption = ""
   End If
Else
    lblCep.Caption = ""
End If

End Sub

Private Sub txtNumImovel_LostFocus()
lblCep.Caption = RetornaCEP(Val(lblCodLogr.Caption), Val(txtNumImovel))
End Sub

Private Sub txtQtdeImovel_GotFocus()
nQtdeLote = Val(txtQtdeImovel.Text)
Sql = "SELECT * FROM DESMEMBRAMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        bExists = False
    Else
        bExists = True
    End If
   .Close
End With

End Sub

Private Sub txtQtdeImovel_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    With cmbNumTmp
        If Val(txtQtdeImovel.Text) = 0 Then
            MsgBox "Digite a qtde de imóveis em que deseja desmembrar.", vbExclamation, "Atenção"
           .BackColor = .Container.BackColor
           .Locked = True
           .Clear
            pnImovel.Caption = ""
            txtQtdeImovel.SetFocus
        Else
           .Clear
           .AddItem "Principal"
            If bImovelEmAndamento Then
                For x = 1 To Val(nQtdeImovel)
                    .AddItem Format(x, "000")
                Next
            Else
                nQtdeImovel = Val(txtQtdeImovel.Text)
                For x = 1 To nQtdeImovel
                    .AddItem Format(x, "000")
                Next
                'GRAVA NA TABELA DESMEMBRAMENTO
                Sql = "SELECT * FROM DESMEMBRAMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                     If .RowCount = 0 Then
                        Sql = "INSERT DESMEMBRAMENTO(CODREDUZIDO,QTDEIMOVEL) VALUES("
                        Sql = Sql & Val(txtCod.Text) & "," & Val(nQtdeImovel) & ")"
                        cn.Execute Sql, rdExecDirect
                     End If
                    .Close
                End With
            End If
           .BackColor = Branco
           .Locked = False
           .ListIndex = 0
           .SetFocus
            nCodOld = 0: nCodNew = 0
        End If
    End With
Else
    Tweak txtQtdeImovel, KeyAscii, IntegerPositive
End If

End Sub

Private Sub txtQtdeImovel_LostFocus()

If Val(txtQtdeImovel.Text) = 0 And Val(txtCod.Text) > 0 And txtQtdeImovel.BackColor = Branco Then
    MsgBox "Digite a qtde de lotes.", vbCritical, "Atenção"
    txtCod.SetFocus
    Exit Sub
End If

If bExists And nQtdeLote <> Val(txtQtdeImovel.Text) Then
    If MsgBox("Alterar a qtde de lotes?" & vbCrLf & " Os lotes acima desta quantidade serão excluidos.", vbQuestion + vbYesNo, "Atenção") = vbYes Then
       Sql = "UPDATE DESMEMBRAMENTO SET QTDEIMOVEL=" & Val(txtQtdeImovel.Text)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM DESMPROPRIETARIO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP > " & Val(txtQtdeImovel.Text)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM DESMAREAS WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP > " & Val(txtQtdeImovel.Text)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM DESMTESTADA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP > " & Val(txtQtdeImovel.Text)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM DESMTEMP WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTMP > " & Val(txtQtdeImovel.Text)
       cn.Execute Sql, rdExecDirect
       nQtdeImovel = Val(txtQtdeImovel.Text)
       cmbNumTmp.Clear
       cmbNumTmp.AddItem "Principal"
       For x = 1 To Val(nQtdeImovel)
          cmbNumTmp.AddItem Format(x, "000")
       Next
    End If
End If

If Val(txtQtdeImovel.Text) > 0 Then
   txtQtdeImovel_KeyPress vbKeyReturn
Else
   txtQtdeImovel.BackColor = txtQtdeImovel.Container.BackColor
   txtQtdeImovel.Locked = True
   pnImovel.Caption = ""
   cmbNumTmp.BackColor = cmbNumTmp.Container.BackColor
   cmbNumTmp.Locked = True
   cmbNumTmp.Clear
End If

End Sub

Private Sub BuildTreeProp()

With tvProp
Inicio:
    For x = 1 To .Nodes.Count
       .Nodes.Remove (x)
       GoTo Inicio:
    Next
End With

With tvProp
   .ImageList = ImlTv
    Set NodX = .Nodes.Add(, , "PROP", "Proprietários", 1)
    Set NodX = .Nodes.Add(, , "COMP", "Proprietário Solidário", 1)
End With

With tvProp
    For x = 1 To .Nodes.Count
       .Nodes(x).EnsureVisible
    Next
   .Nodes("PROP").Bold = True
   .Nodes("COMP").Bold = True
End With

End Sub


Private Sub Limpa()
Dim z As Long
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)

txtQuadra.Text = ""
txtLote.Text = ""
txtNumFace.Text = ""
lblCodLogr.Caption = ""
lblCep.Caption = ""
txtNumImovel.Text = ""
lblNomeLogr.Caption = ""
'txtCompl.Text = ""
'txtQuadras.text = ""
txtLotes.Text = ""
txtAreaTerreno.Text = ""
cmbBenf.ListIndex = -1
cmbCatProp.ListIndex = -1
cmbPedol.ListIndex = -1
cmbSit.ListIndex = -1
cmbTopog.ListIndex = -1
cmbUso.ListIndex = -1
txtFracao.Text = ""
grdTestada.Rows = 1


End Sub

Private Function ValidaFace() As Boolean
Dim nDist As Integer, nSetor As Integer, nQuadra As Integer, nFace As Integer

nDist = Left$(lblIC.Caption, 1)
nSetor = Mid$(lblIC.Caption, 3, 2)
nQuadra = Val(txtQuadra.Text)
nFace = Val(txtNumFace.Text)

Sql = "SELECT * FROM vwFACEQUADRA WHERE CODDISTRITO=" & nDist & " AND CODSETOR=" & nSetor
Sql = Sql & " AND CODQUADRA=" & nQuadra & " AND CODFACE=" & nFace
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        ValidaFace = True
        lblCodLogr.Caption = Format(!CodLogr, "00000")
        lblNomeLogr.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
        lblCep.Caption = RetornaCEP(Val(lblCodLogr.Caption), Val(txtNumImovel))
    Else
        MsgBox "Face de Quadra não cadastrada.", vbExclamation, "atenção"
        ValidaFace = False
    End If
   .Close
End With

End Function

Private Sub txtQuadra_GotFocus()
txtQuadra.SelStart = 0
txtQuadra.SelLength = Len(txtQuadra.Text)
End Sub

Private Sub txtQuadra_LostFocus()

If Val(txtQuadra.Text) > 0 And Val(txtNumFace.Text) > 0 Then
   If Not ValidaFace Then
      txtQuadra.SetFocus
   Else
      lblCep.Caption = ""
   End If
Else
   lblCep.Caption = ""
End If

End Sub

Private Sub txtTestada_GotFocus()
txtTestada.SelStart = 0
txtTestada.SelLength = Len(txtTestada.Text)

End Sub

Private Sub txtTestada_KeyPress(KeyAscii As Integer)
Tweak txtTestada, KeyAscii, DecimalPositive
End Sub
