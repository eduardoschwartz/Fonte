VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmUnifica 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unificação de Lote"
   ClientHeight    =   5940
   ClientLeft      =   3345
   ClientTop       =   3570
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvLote 
      Height          =   2235
      Left            =   6870
      TabIndex        =   43
      Top             =   0
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3942
      _Version        =   393217
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImlTv 
      Left            =   1650
      Top             =   5190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":02B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":0418
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":0578
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":06D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUnifica.frx":0830
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvLote2 
      Height          =   2235
      Left            =   6870
      TabIndex        =   60
      Top             =   2940
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   3942
      _Version        =   393217
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin prjChameleon.chameleonButton cmdPFrame 
      Height          =   285
      Left            =   90
      TabIndex        =   62
      ToolTipText     =   "Voltar"
      Top             =   5520
      Width           =   555
      _ExtentX        =   979
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUnifica.frx":098C
      PICN            =   "frmUnifica.frx":09A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNFrame 
      Height          =   285
      Left            =   690
      TabIndex        =   63
      ToolTipText     =   "Avançar"
      Top             =   5520
      Width           =   555
      _ExtentX        =   979
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUnifica.frx":0B02
      PICN            =   "frmUnifica.frx":0B1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   285
      Left            =   2670
      TabIndex        =   64
      ToolTipText     =   "Busca Lote a ser Desmembrado"
      Top             =   60
      Width           =   435
      _ExtentX        =   767
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUnifica.frx":0C78
      PICN            =   "frmUnifica.frx":0C94
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel2 
      Height          =   285
      Left            =   2670
      TabIndex        =   65
      ToolTipText     =   "Busca Lote a ser Desmembrado"
      Top             =   660
      Width           =   435
      _ExtentX        =   767
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUnifica.frx":0DEE
      PICN            =   "frmUnifica.frx":0E0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraD 
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H00000080&
      Height          =   4065
      Index           =   3
      Left            =   0
      TabIndex        =   42
      Top             =   1260
      Width           =   6855
      Begin MSComctlLib.ListView lvArea 
         Height          =   2655
         Left            =   90
         TabIndex        =   61
         Top             =   1290
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   8388672
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lote"
            Object.Width           =   883
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Seq"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Tp"
            Object.Width           =   742
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Área"
            Object.Width           =   1744
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ano"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Uso"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Categoria"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Áreas do Imóvel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   10
         Left            =   180
         TabIndex        =   50
         Top             =   270
         Width           =   1935
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade de Edificações:"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   180
         TabIndex        =   46
         Top             =   600
         Width           =   2055
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione a Área Principal"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   1020
         Width           =   1965
      End
      Begin VB.Label lblQtdeEdif 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2250
         TabIndex        =   44
         Top             =   600
         Width           =   345
      End
   End
   Begin VB.Frame fraD 
      BackColor       =   &H00EEEEEE&
      Height          =   4065
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   1260
      Width           =   6855
      Begin MSComctlLib.TreeView tvProp 
         Height          =   3450
         Left            =   90
         TabIndex        =   5
         Top             =   510
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   6085
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
         Left            =   5250
         TabIndex        =   68
         ToolTipText     =   "Adicionar Proprietário/Compromissário"
         Top             =   930
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUnifica.frx":0F64
         PICN            =   "frmUnifica.frx":0F80
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
         Left            =   5250
         TabIndex        =   69
         ToolTipText     =   "Remover Proprietário/Compromissário"
         Top             =   1320
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUnifica.frx":10DA
         PICN            =   "frmUnifica.frx":10F6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietários/Compromissários"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   47
         Top             =   210
         Width           =   4575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTes1 
      Height          =   525
      Left            =   8430
      TabIndex        =   70
      Top             =   2340
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   926
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdTes2 
      Height          =   525
      Left            =   8400
      TabIndex        =   71
      Top             =   5310
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   926
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      FocusRect       =   0
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin prjChameleon.chameleonButton cmdHelp 
      Height          =   315
      Left            =   2370
      TabIndex        =   72
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   5460
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Ajuda"
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
      MICON           =   "frmUnifica.frx":1250
      PICN            =   "frmUnifica.frx":126C
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
      Left            =   5640
      TabIndex        =   73
      ToolTipText     =   "Sair da Tela"
      Top             =   5460
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
      MICON           =   "frmUnifica.frx":13C6
      PICN            =   "frmUnifica.frx":13E2
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
      Left            =   4560
      TabIndex        =   74
      ToolTipText     =   "Gravar o Registro"
      Top             =   5460
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
      MICON           =   "frmUnifica.frx":1450
      PICN            =   "frmUnifica.frx":146C
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
      Left            =   3450
      TabIndex        =   75
      ToolTipText     =   "Cancelar Edição"
      Top             =   5460
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmUnifica.frx":1811
      PICN            =   "frmUnifica.frx":182D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraD 
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H00000080&
      Height          =   4065
      Index           =   2
      Left            =   0
      TabIndex        =   19
      Top             =   1260
      Width           =   6855
      Begin VB.ComboBox cmbBenf 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1440
         Width           =   1905
      End
      Begin VB.ComboBox cmbUso 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1080
         Width           =   1905
      End
      Begin VB.TextBox txtAreaTerreno 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "0,00"
         Top             =   660
         Width           =   1185
      End
      Begin VB.ComboBox cmbCatProp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2160
         Width           =   1905
      End
      Begin VB.ComboBox cmbSit 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2520
         Width           =   1905
      End
      Begin VB.ComboBox cmbTopog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   1800
         Width           =   1905
      End
      Begin VB.ComboBox cmbPedol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   2880
         Width           =   1905
      End
      Begin VB.TextBox txtFracaoIdeal 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1230
         TabIndex        =   24
         Text            =   "0,00"
         Top             =   3555
         Width           =   795
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Testadas"
         ForeColor       =   &H00000080&
         Height          =   2655
         Left            =   4050
         TabIndex        =   20
         Top             =   330
         Width           =   2475
         Begin VB.TextBox txtTestada 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   660
            TabIndex        =   33
            Text            =   "0,00"
            Top             =   2280
            Width           =   885
         End
         Begin VB.TextBox txtFace 
            Appearance      =   0  'Flat
            Height          =   255
            Left            =   660
            TabIndex        =   32
            Text            =   "0"
            Top             =   1950
            Width           =   435
         End
         Begin MSFlexGridLib.MSFlexGrid grdTestada 
            Height          =   1455
            Left            =   90
            TabIndex        =   21
            Top             =   330
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   2566
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            BackColorBkg    =   15658734
            FocusRect       =   0
            SelectionMode   =   1
            BorderStyle     =   0
            Appearance      =   0
            FormatString    =   "^Face        |^Área  m²          "
         End
         Begin prjChameleon.chameleonButton cmdAddTestada 
            Height          =   285
            Left            =   1680
            TabIndex        =   66
            ToolTipText     =   "Adicionar Testada"
            Top             =   2250
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
            MICON           =   "frmUnifica.frx":1987
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
            Left            =   2010
            TabIndex        =   67
            ToolTipText     =   "Remover Testada"
            Top             =   2250
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
            MICON           =   "frmUnifica.frx":19A3
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
            BackColor       =   &H00EEEEEE&
            Caption         =   "Face:"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   1980
            Width           =   495
         End
         Begin VB.Label Label4 
            BackColor       =   &H00EEEEEE&
            Caption         =   "metros:"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   2280
            Width           =   495
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Dados do Terreno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   9
         Left            =   180
         TabIndex        =   49
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Categ. Propriedade:"
         Height          =   225
         Left            =   120
         TabIndex        =   41
         Top             =   2235
         Width           =   1425
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Benfeitoria.............:"
         Height          =   225
         Left            =   120
         TabIndex        =   40
         Top             =   1500
         Width           =   1425
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Uso do Terreno.....:"
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Área do Terreno....:"
         Height          =   225
         Left            =   120
         TabIndex        =   38
         Top             =   690
         Width           =   1425
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação................:"
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   2610
         Width           =   1425
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Topografia.............:"
         Height          =   225
         Left            =   120
         TabIndex        =   36
         Top             =   1875
         Width           =   1425
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedologia..............:"
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   2970
         Width           =   1425
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Fração Ideal:"
         Height          =   225
         Left            =   180
         TabIndex        =   34
         Top             =   3555
         Width           =   1005
      End
   End
   Begin VB.Frame fraD 
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H00000080&
      Height          =   4065
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   1260
      Width           =   6855
      Begin VB.TextBox txtNumFace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Text            =   "0"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtLotes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   11
         Top             =   3240
         Width           =   2835
      End
      Begin VB.TextBox txtQuadras 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   25
         TabIndex        =   10
         Top             =   2910
         Width           =   2835
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2160
         Width           =   4845
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "0"
         Top             =   1830
         Width           =   855
      End
      Begin VB.TextBox txtCodLogrLI 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   52
         Top             =   1065
         Width           =   855
      End
      Begin VB.TextBox txtNomeLogLI 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1425
         Width           =   4830
      End
      Begin VB.Label lblBairro 
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   1470
         TabIndex        =   54
         Top             =   2580
         Width           =   3765
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da Face........:"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   51
         Top             =   735
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Local do Imóvel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   48
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Número..............:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   1845
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.....:"
         Height          =   225
         Left            =   180
         TabIndex        =   17
         Top             =   2205
         Width           =   1305
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro.................:"
         Height          =   225
         Left            =   180
         TabIndex        =   16
         Top             =   2565
         Width           =   1305
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadras............:"
         Height          =   225
         Left            =   180
         TabIndex        =   15
         Top             =   2940
         Width           =   1305
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes.................:"
         Height          =   225
         Left            =   180
         TabIndex        =   14
         Top             =   3300
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cód.Logradouro.:"
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   13
         Top             =   1110
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Lograd.....:"
         Height          =   225
         Index           =   6
         Left            =   180
         TabIndex        =   12
         Top             =   1470
         Width           =   1275
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o 2º Lote a ser Unificado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   30
      TabIndex        =   59
      Top             =   690
      Width           =   2625
   End
   Begin VB.Label lblNumInsc2 
      BackStyle       =   0  'Transparent
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
      Left            =   3150
      TabIndex        =   58
      Top             =   690
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   30
      TabIndex        =   57
      Top             =   960
      Width           =   885
   End
   Begin VB.Label lblProp2 
      BackStyle       =   0  'Transparent
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
      Left            =   960
      TabIndex        =   56
      Top             =   960
      Width           =   5745
   End
   Begin VB.Label lblOK 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1380
      TabIndex        =   55
      Top             =   5520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblProp 
      BackStyle       =   0  'Transparent
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
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   5745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   885
   End
   Begin VB.Label lblNumInsc 
      BackStyle       =   0  'Transparent
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
      Left            =   3180
      TabIndex        =   1
      Top             =   90
      Width           =   3765
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o 1º Lote a ser Unificado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   2625
   End
End
Attribute VB_Name = "frmUnifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frAtivo As Integer
Dim RdoAux As rdoResultset
Dim RdoAux2 As rdoResultset
Dim Sql As String, bExec As Boolean
Dim nSomaTestada As Double
Dim nLoteSel As Integer
Private Type Testada
       nFace As Integer
       nArea As Double
End Type
Dim aTestada() As Testada
Dim xImovel As New clsImovel
Dim nCodReduz As Long
Dim nCodReduz1 As Long
Dim nCodReduz2 As Long

Private Sub cmdAddCid_Click()
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
   MsgBox "Selecione na árvore Proprietário ou Compromissário.", vbExclamation, "Atenção"

End Sub

Private Sub cmdAddTestada_Click()
Dim Achou As Boolean
Dim x As Integer

If Val(txtFace.text) = 0 Then
   MsgBox "Digite a Face da Testada.", vbExclamation, "Atenção"
   txtFace.SetFocus
   Exit Sub
End If

If CDbl(txtTestada.text) = 0 Then
   MsgBox "Digite a Área da Testada.", vbExclamation, "Atenção"
   txtTestada.SetFocus
   Exit Sub
End If

If grdTestada.Rows = 1 And (Val(txtFace.text) <> Val(Mid$(frmUnifica.lblNumInsc.Caption, 17, 2)) And Val(txtFace.text) <> Val(Mid(frmUnifica.lblNumInsc2.Caption, 17, 2))) Then
   MsgBox "O número da testada principal deve ser igual a testada principal de um dos lotes unificados.", vbExclamation, "Atenção"
   txtFace.SetFocus
   Exit Sub
End If

Achou = False
For x = 1 To grdTestada.Rows - 1
   If Val(grdTestada.TextMatrix(x, 0)) = Val(txtFace.text) Then
      Achou = True
      Exit For
   End If
Next
If Achou Then
   MsgBox "Face já cadastrada.", vbExclamation, "Atenção"
   txtFace.SetFocus
   Exit Sub
End If

nSomaTestada = 0
For x = 0 To grdTes1.Rows - 1
    nSomaTestada = nSomaTestada + CDbl(grdTes1.TextMatrix(x, 1))
Next
For x = 0 To grdTes2.Rows - 1
    nSomaTestada = nSomaTestada + CDbl(grdTes2.TextMatrix(x, 1))
Next
nSoma = 0
For x = 1 To grdTestada.Rows - 1
    nSoma = nSoma + CDbl(grdTestada.TextMatrix(x, 1))
Next

If CDbl(txtTestada.text) + nSoma > nSomaTestada Then
     MsgBox "A Soma de todas as testadas não pode ultrapassar a soma das testadas dos lotes unificados (" & FormatNumber(nSomaTestada, 2) & ")."
     Exit Sub
End If

nSoma = 0

grdTestada.AddItem Format(txtFace.text, "00") & Chr(9) & Format(txtTestada, "0#.00")
txtFace.text = Val(grdTestada.TextMatrix(grdTestada.Rows - 1, 0)) + 1
txtTestada.text = "0,00"
txtFace.SetFocus

End Sub

Private Sub cmdCancel_Click()
If MsgBox("Cancelar a Unificação ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    cmdCnsImovel.Enabled = True
    cmdCnsImovel2.Enabled = True
    lblNumInsc.Caption = ""
    lblNumInsc2.Caption = ""
    lblProp.Caption = ""
    lblProp2.Caption = ""
    txtNumFace.text = ""
    txtCodLogrLI.text = 0
    txtNomeLogLI.text = ""
    txtNum.text = ""
    txtCompl.text = ""
    lblBairro.Caption = ""
    txtQuadras.text = ""
    txtLotes.text = ""
    txtAreaTerreno.text = 0
    cmbUso.ListIndex = -1
    cmbBenf.ListIndex = -1
    cmbTopog.ListIndex = -1
    cmbCatProp.ListIndex = -1
    cmbSit.ListIndex = -1
    cmbPedol.ListIndex = -1
    txtFracaoIdeal.text = "0,00"
    grdTestada.Rows = 1
    grdTes1.Rows = 1
    grdTes1.TextMatrix(0, 0) = ""
    grdTes1.TextMatrix(0, 1) = ""
    grdTes2.Rows = 1
    grdTes2.TextMatrix(0, 0) = ""
    grdTes2.TextMatrix(0, 1) = ""
    BuildTreeProp
    BuildTreeLote
    ReDim aTestada(0)
End If

End Sub

Private Sub cmdCnsImovel_Click()
sForm = "UN"
nLoteSel = 1
frmCnsImovel.show
frmCnsImovel.ZOrder 0
End Sub

Private Sub cmdCnsImovel2_Click()
sForm = "UN"
nLoteSel = 2
frmCnsImovel.show
frmCnsImovel.ZOrder 0
End Sub

Private Sub cmdDelCid_Click()
On Error GoTo Erro
'remove da arvore
n = frmUnifica.tvProp.SelectedItem.Parent.Index
nc = tvProp.SelectedItem.Index
tvProp.Nodes.Remove (nc)
If frmUnifica.tvProp.Nodes("PROP").Children > 0 Then
    If Right$(frmUnifica.tvProp.Nodes("PROP").Child.text, 9) <> "Principal" Then
          frmUnifica.tvProp.Nodes("PROP").Child.text = frmUnifica.tvProp.Nodes("PROP").Child.text & " - Principal"
    End If
End If
Exit Sub
   
Erro:
   MsgBox "Selecione o Proprietário/Compromissário que deseja Remover.", vbExclamation, "Atenção"
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

Private Sub cmdGravar_Click()

If Valida() Then
    GravaLote
    'nCodReduz1 = 26974
    'nCodReduz2 = 23371
    'nCodReduz = 26983
    TransfereDivida nCodReduz1, nCodReduz2, nCodReduz
    MsgBox "Unificação efetuada com sucesso.", vbExclamation, "Atenção"
    Unload Me
End If

End Sub

Public Sub TransfereDivida(nCodAntigo1 As Long, nCodAntigo2 As Long, nCodNovo As Long)
Dim nSeq As Integer

'*************************************
'CARREGA DIVIDA DO PRIMEIRO IMÓVEL
'*************************************

'DEBITOPARCELA
Sql = "INSERT DEBITOPARCELA SELECT " & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,"
Sql = Sql & "VALORJUROS,NUMPROCESSO,INTACTO From DEBITOPARCELA Where CODREDUZIDO = " & nCodAntigo1
cn.Execute Sql, rdExecDirect

'DEBITOTRIBUTO
Sql = "INSERT DEBITOTRIBUTO SELECT " & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
Sql = Sql & "VALORTRIBUTO,VALORCORRECAO,VALORMULTA,VALORJUROS,INTACTO,VALORPORBAIXA From DEBITOTRIBUTO Where CODREDUZIDO = " & nCodAntigo1
cn.Execute Sql, rdExecDirect

'DEBITOPAGO
Sql = "INSERT DEBITOPAGO SELECT " & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQPAG,"
Sql = Sql & "DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,RESTITUIDO,NUMDOCUMENTO,VALORPAGOREAL,INTACTO,VALORTARIFA,ARQUIVOBANCO,VALORDIF From DEBITOPAGO Where CODREDUZIDO = " & nCodAntigo1
cn.Execute Sql, rdExecDirect

'PARCELADOCUMENTO
Sql = "SELECT * FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & nCodAntigo1
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodNovo & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
       Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
       Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            If .RowCount > 0 Then
                Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodNovo & ","
                Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
                Sql = Sql & !CODCOMPLEMENTO & "," & RdoAux!NumDocumento & ")"
                cn.Execute Sql, rdExecDirect
                Sql = "DELETE FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & nCodAntigo1 & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
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
Sql = "DELETE FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & nCodAntigo1
cn.Execute Sql, rdExecDirect

'REPARCELAMENTO
Sql = "INSERT ORIGEMREPARC SELECT NUMPROCESSO," & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO,"
Sql = Sql & "NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,PRINCIPAL,JUROS,MULTA,CORRECAO FROM ORIGEMREPARC WHERE CODREDUZIDO=" & nCodAntigo1
cn.Execute Sql, rdExecDirect

Sql = "INSERT DESTINOREPARC SELECT NUMPROCESSO," & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO,"
Sql = Sql & "NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,VALORLIQUIDO,JUROS,MULTA,CORRECAO,VALORPRINCIPAL,SALDO,JUROSPERC,JUROSVALOR,JUROSAPL,HONORARIO,TOTAL FROM DESTINOREPARC WHERE CODREDUZIDO=" & nCodAntigo1
cn.Execute Sql, rdExecDirect

Sql = "UPDATE PROCESSOREPARC SET CODIGORESP=" & nCodNovo & " WHERE CODIGORESP=" & nCodAntigo1
cn.Execute Sql, rdExecDirect

'GoTo fim

 'Exit Sub
'*************************************
'CARREGA DIVIDA DO SEGUNDO IMÓVEL
'*************************************

'NOVA SEQUENCIA
Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO = " & nCodAntigo2
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodNovo
        Sql = Sql & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento
        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
        Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If IsNull(!MAXIMO) Then
                nSeq = RdoAux!SeqLancamento
            Else
                nSeq = !MAXIMO + 1
            End If
           .Close
        End With
        On Error Resume Next
        'DEBITOPARCELA
        Sql = "INSERT DEBITOPARCELA SELECT " & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO," & nSeq & ",NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,"
        Sql = Sql & "VALORJUROS,NUMPROCESSO,INTACTO From DEBITOPARCELA Where CODREDUZIDO = " & nCodAntigo2 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento
        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
                
        'DEBITOTRIBUTO
        Sql = "INSERT DEBITOTRIBUTO SELECT " & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO," & nSeq & ",NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
        Sql = Sql & "VALORTRIBUTO,VALORCORRECAO,VALORMULTA,VALORJUROS,INTACTO From DEBITOTRIBUTO Where CODREDUZIDO = " & nCodAntigo2 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento
        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        'DEBITOPAGO
        Sql = "INSERT DEBITOPAGO SELECT " & nCodNovo & ",ANOEXERCICIO,CODLANCAMENTO," & nSeq & ",NUMPARCELA,CODCOMPLEMENTO,SEQPAG,"
        Sql = Sql & "DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,RESTITUIDO,NUMDOCUMENTO,VALORPAGOREAL,INTACTO,VALORTARIFA,ARQUIVOBANCO From DEBITOPAGO Where CODREDUZIDO = " & nCodAntigo2 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento
        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        'PARCELADOCUMENTO
        Sql = "SELECT * FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & nCodAntigo2
        Sql = Sql & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento
        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodNovo & ","
                Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & nSeq & "," & !NumParcela & ","
                Sql = Sql & !CODCOMPLEMENTO & "," & RdoAux2!NumDocumento & ")"
                cn.Execute Sql, rdExecDirect
                Sql = "DELETE FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & nCodAntigo2 & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
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

Fim:

'STATUS TRANSFERIDO
Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & nCodAntigo1 & " AND STATUSLANC=3"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & nCodAntigo2 & " AND STATUSLANC=3"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub GravaLote()
Dim nDist As Integer
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Long
Dim nUnidade As Integer
Dim sHist As String
Dim qd As New rdoQuery

Set qd.ActiveConnection = cn

nCodReduz1 = Val(Left$(Right$(lblNumInsc.Caption, 10), 7))
nCodReduz2 = Val(Left$(Right$(lblNumInsc2.Caption, 10), 7))

nDist = Left$(lblNumInsc.Caption, 1)
nSetor = Mid(lblNumInsc.Caption, 3, 2)
nQuadra = Mid(lblNumInsc.Caption, 6, 4)
nUnidade = Mid(lblNumInsc.Caption, 20, 2)

Sql = "SELECT MAX(LOTE) AS ULTIMOLOTE FROM CADIMOB WHERE "
Sql = Sql & "DISTRITO=" & nDist & " AND SETOR=" & nSetor & " AND QUADRA=" & nQuadra
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
nLote = RdoAux!ULTIMOLOTE + 1

Sql = "SELECT MAX(CODREDUZIDO) AS ULTIMOCODREDUZ FROM CADIMOB WHERE CODREDUZIDO<40000"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
nCodReduz = RdoAux!ULTIMOCODREDUZ + 1

GoSub GravaImovel
GoSub GravaProprietario
GoSub GravaTestada
GoSub GravaArea

'Inativa os 2 Imoveis
Sql = "UPDATE CADIMOB SET INATIVO=1 WHERE CODREDUZIDO=" & nCodReduz1
cn.Execute Sql, rdExecDirect
Sql = "UPDATE CADIMOB SET INATIVO=1 WHERE CODREDUZIDO=" & nCodReduz2
cn.Execute Sql, rdExecDirect

'Grava o Histórico dos Lotes
'Novo Lote
Sql = "SELECT CODREDUZIDO,SEQ FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
If RdoAux.RowCount > 0 Then
    nSeq = RdoAux!Seq + 1
Else
   nSeq = 1
End If
sHist = "O imóvel foi criado a partir da Unificação dos Imóveis: " & vbCrLf & Format(nCodReduz1, "0000000") & "-" & RetornaDVCodReduzido(nCodReduz1) & " e " & Format(nCodReduz2, "0000000") & "-" & RetornaDVCodReduzido(nCodReduz2)
Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST) VALUES("
Sql = Sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sHist & "')"
cn.Execute Sql, rdExecDirect

'Lote 1
Sql = "SELECT max(SEQ) as MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz1
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
If Not IsNull(RdoAux!MAXIMO) Then
    nSeq = RdoAux!MAXIMO + 1
Else
   nSeq = 1
End If
sHist = "O imóvel foi unificado com o imóvel " & Format(nCodReduz2, "0000000") & "-" & RetornaDVCodReduzido(nCodReduz2) & " e criou o Imóvel " & Format(nCodReduz, "0000000") & "-" & RetornaDVCodReduzido(nCodReduz)
Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST) VALUES("
Sql = Sql & nCodReduz1 & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "')"
cn.Execute Sql, rdExecDirect

'Lote 2
Sql = "SELECT CODREDUZIDO,SEQ FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz2
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
If RdoAux.RowCount > 0 Then
    nSeq = RdoAux!Seq + 1
Else
   nSeq = 1
End If
sHist = "O imóvel foi unificado com o imóvel " & Format(nCodReduz1, "0000000") & "-" & RetornaDVCodReduzido(nCodReduz1) & " e criou o Imóvel " & Format(nCodReduz, "0000000") & "-" & RetornaDVCodReduzido(nCodReduz)
Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST) VALUES("
Sql = Sql & nCodReduz2 & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "')"
cn.Execute Sql, rdExecDirect

'FINALIZA
lblOK.Caption = "GRAVADO"
MsgBox "O Imóvel foi unificado com sucesso, criando o imóvel " & nCodReduz & "-" & RetornaDVCodReduzido(nCodReduz) & ".", vbInformation, "Informação"
Exit Sub
'SAI
'*******GRAVA IMOVEL**********************************************

GravaImovel:
Sql = "INSERT CADIMOB(CODREDUZIDO,DV,CODCONDOMINIO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,LI_NUM,LI_COMPL,"
Sql = Sql & "LI_UF,LI_CODCIDADE,LI_CODBAIRRO,LI_QUADRAS,LI_LOTES,DT_AREATERRENO,DT_CODUSOTERRENO,DT_CODBENF,DT_CODTOPOG,"
Sql = Sql & "DT_CODCATEGPROP,DT_CODSITUACAO,DT_CODPEDOL,DT_NUMAGUA,DT_FRACAOIDEAL,DC_QTDEEDIF,DC_QTDEPAV,EE_TIPOEND) values("
Sql = Sql & nCodReduz & "," & RetornaDVCodReduzido(nCodReduz) & "," & 999 & ","
Sql = Sql & nDist & "," & nSetor & "," & nQuadra & "," & nLote & ","
Sql = Sql & Val(txtNumFace.text) & "," & 0 & "," & 0 & "," & Val(txtNum.text) & ",'"
Sql = Sql & Mask(txtCompl.text) & "','" & "SP" & "'," & 413 & "," & IIf(lblBairro.Caption <> "", Val(Left$(lblBairro.Caption, 3)), "Null") & ",'"
Sql = Sql & Mask(txtQuadras.text) & "','" & Mask(txtLotes.text) & "'," & Virg2Ponto(RemovePonto(txtAreaTerreno.text)) & "," & IIf(cmbUso.ListIndex > -1, cmbUso.ItemData(cmbUso.ListIndex), "Null") & ","
Sql = Sql & IIf(cmbBenf.ListIndex > -1, cmbBenf.ItemData(cmbBenf.ListIndex), "Null") & "," & IIf(cmbTopog.ListIndex > -1, cmbTopog.ItemData(cmbTopog.ListIndex), "Null") & ","
Sql = Sql & IIf(cmbCatProp.ListIndex > -1, cmbCatProp.ItemData(cmbCatProp.ListIndex), "Null") & "," & IIf(cmbSit.ListIndex > -1, cmbSit.ItemData(cmbSit.ListIndex), "Null") & ","
Sql = Sql & IIf(cmbPedol.ListIndex > -1, cmbPedol.ItemData(cmbPedol.ListIndex), "Null") & "," & "Null" & "," & Virg2Ponto(txtFracaoIdeal.text) & "," & Val(lblQtdeEdif.Caption) & ",0,0)"
cn.Execute Sql, rdExecDirect

Return
''*********************************************************************
''*******GRAVA PROPRIETARIO *******************************************
GravaProprietario:
For x = 1 To tvProp.Nodes.Count
    If Len(tvProp.Nodes(x).Key) > 4 Then
        Sql = "INSERT PROPRIETARIO (CODREDUZIDO,CODCIDADAO,TIPOPROP,PRINCIPAL) VALUES("
        Sql = Sql & nCodReduz & "," & Val(Right$(tvProp.Nodes(x).Key, 6)) & ",'"
        Sql = Sql & Left$(tvProp.Nodes(x).Key, 1) & "'," & IIf(tvProp.Nodes("PROP").Child.text = tvProp.Nodes(x).text, 1, 0) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next
Return
''*********************************************************************
''*******GRAVA TESTADA *******************************************
GravaTestada:
For x = 1 To grdTestada.Rows - 1
    Sql = "INSERT TESTADA(CODREDUZIDO,NUMFACE,AREATESTADA) VALUES("
    Sql = Sql & nCodReduz & "," & Val(grdTestada.TextMatrix(x, 0)) & "," & Virg2Ponto(grdTestada.TextMatrix(x, 1)) & ")"
    cn.Execute Sql, rdExecDirect
Next
Return
''*********************************************************************
''*******GRAVA AREA *******************************************
GravaArea:
For x = 1 To lvArea.ListItems.Count
    sData = lvArea.ListItems(x).ListSubItems(4).text
    Sql = "INSERT AREAS (CODREDUZIDO,SEQAREA,TIPOAREA,DATAAPROVA,AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR) VALUES(" & nCodReduz & "," & x & ",'" & IIf(lvArea.ListItems(x).Checked, "P", "C") & "',"
    Sql = Sql & IIf(IsDate(sData), "'" & Format(sData, "mm/dd/yyyy") & "'", "Null") & "," & Virg2Ponto(RemovePonto(Left$(lvArea.ListItems(x).ListSubItems(3).text, Len(lvArea.ListItems(x).ListSubItems(3).text) - 3))) & ","
    Sql = Sql & Val(Left$(lvArea.ListItems(x).ListSubItems(5).text, 2)) & "," & Val(Left$(lvArea.ListItems(x).ListSubItems(6).text, 2)) & "," & Val(Left$(lvArea.ListItems(x).ListSubItems(7).text, 2)) & ")"
    cn.Execute Sql, rdExecDirect
Next
Return
''*********************************************************************

End Sub

Private Function Valida() As Boolean
Dim Achou As Boolean
Dim nDist1 As Integer, nDist2 As Integer
Dim nSetor1 As Integer, nSetor2 As Integer
Dim nQuadra1 As Integer, nQuadra2 As Integer
Dim Erro As Boolean
Dim nLote1 As Integer
Dim nLote2 As Integer
Dim nUnidade1 As Integer
Dim nUnidade2 As Integer
Dim nCodReduz1 As Long
Dim nCodReduz2 As Long

Valida = True

If lblNumInsc.Caption = "" Or lblNumInsc2.Caption = "" Then
     MsgBox "Selecione os 2 Lotes que serão unificados.", vbExclamation, "Atenção"
     GoTo Falso
End If

nDist1 = Left$(lblNumInsc.Caption, 1)
nSetor1 = Mid$(lblNumInsc.Caption, 3, 2)
nQuadra1 = Mid$(lblNumInsc.Caption, 6, 4)
nLote1 = Mid$(lblNumInsc.Caption, 11, 5)
nUnidade1 = Mid$(lblNumInsc.Caption, 20, 2)

nDist2 = Left$(lblNumInsc2.Caption, 1)
nSetor2 = Mid$(lblNumInsc2.Caption, 3, 2)
nQuadra2 = Mid$(lblNumInsc2.Caption, 6, 4)
nLote2 = Mid$(lblNumInsc.Caption, 11, 5)
nUnidade2 = Mid$(lblNumInsc.Caption, 20, 2)

If (nUnidade1 > 0 And nUnidade2 = 0) Or (nUnidade1 = 0 And nUnidade2 > 0) Then
    MsgBox "Não é possivel unificar um lote normal com uma subunidade de um condomínio.", vbExclamation, "Atenção"
    GoTo Falso
End If

If nUnidade1 = 0 Then 'se não for condominio
    If nDist1 = nDist2 And nSetor1 = nSetor2 And nQuadra1 = nQuadra2 Then
    Else
        MsgBox "Os 2 lotes a serem unificados devem pertencer a mesma Quadra.", vbExclamation, "Atenção"
        GoTo Falso
    End If
Else ' se for
    If nDist1 = nDist2 And nSetor1 = nSetor2 And nQuadra1 = nQuadra2 And nLote1 = nLote2 Then
    Else
        MsgBox "Os 2 lotes a serem unificados devem pertencer ao mesmo Lote.", vbExclamation, "Atenção"
        GoTo Falso
    End If
End If

nCodReduz1 = Val(Left$(Right$(lblNumInsc.Caption, 10), 7))
nCodReduz2 = Val(Left$(Right$(lblNumInsc2.Caption, 10), 7))

Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & nCodReduz1
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
If RdoAux!Inativo = True Then
     MsgBox "O Imóvel 1 (" & Left$(Right$(lblNumInsc.Caption, 10), 9) & ") encontra-se Inativo, verifique o Histórico do Imóvel para maiores informações.", vbInformation, "Atenção"
     RdoAux.Close
     GoTo Falso
End If

Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & nCodReduz2
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
If RdoAux!Inativo = True Then
     MsgBox "O Imóvel 2 (" & Left$(Right$(lblNumInsc2.Caption, 10), 9) & ") encontra-se Inativo, verifique o Histórico do Imóvel para maiores informações.", vbInformation, "Atenção"
     RdoAux.Close
     GoTo Falso
End If

'Proprietário
If tvProp.Nodes("PROP").Children = 0 Then
   MsgBox "Selecione um Proprietário.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

'Local do Imovel
If Val(txtNumFace.text) = 0 Then
   MsgBox "Digite a Face do Imóvel.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If Val(txtNumFace.text) <> tvLote.Nodes("FACE").Tag And Val(txtNumFace.text) <> tvLote2.Nodes("FACE").Tag Then
   MsgBox "O Imóvel só pode ter Face para uma das faces dos lotes unificados.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If Val(txtNum.text) = 0 Then
   MsgBox "Digite o Nº do Imóvel.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If Val(txtNum.text) <> tvLote.Nodes("NUM").Tag Then
     If tvLote.Nodes("NUM").Tag > 0 Then
          If Val(txtNum.text) <> tvLote2.Nodes("NUM").Tag Then
               If tvLote2.Nodes("NUM").Tag > 0 Then
                    Erro = True
                    MsgBox "O imovel só pode ter os numeros dos lotes unificados.", vbExclamation, "Atenção"
                    GoTo Falso
               End If
          End If
     End If
Else
End If

'Dados do Terreno

If cmbUso.ListIndex = -1 Then
   MsgBox "Selecione o Uso do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If cmbBenf.ListIndex = -1 Then
   MsgBox "Selecione a Benfeitoria do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If cmbTopog.ListIndex = -1 Then
   MsgBox "Selecione a Topografia do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If cmbCatProp.ListIndex = -1 Then
   MsgBox "Selecione a Categoria do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If cmbSit.ListIndex = -1 Then
   MsgBox "Selecione a Situação do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

If cmbPedol.ListIndex = -1 Then
   MsgBox "Selecione a Pedologia do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

'Area
If lvArea.ListItems.Count > 0 Then
    Achou = False
    For x = 1 To lvArea.ListItems.Count
          If lvArea.ListItems(x).Checked = True Then
               Achou = True
          End If
    Next
    
    If Not Achou Then
       MsgBox "Selecione a Área Principal.", vbCritical, "Erro de Validação."
       GoTo Falso
    End If
End If

Valida = True
Exit Function

Falso:
   Valida = False

End Function

Private Sub cmdHelp_Click()
  With hHelp
    .CHMFile = sPathHelp & "\Tribut.chm"
    .HHTopicID = 1300
    .HHWindow = "Main"
    .HHDisplayTopicID
  End With
End Sub

Private Sub cmdNFrame_Click()
Dim nDist1 As Integer, nDist2 As Integer
Dim nSetor1 As Integer, nSetor2 As Integer
Dim nQuadra1 As Integer, nQuadra2 As Integer
Dim nLote1 As Integer, nLote2 As Integer

If lblNumInsc.Caption = "" Or lblNumInsc2.Caption = "" Then
     MsgBox "Selecione os 2 Lote a serem Unificados.", vbExclamation, "Atenção"
     Exit Sub
End If

nDist1 = Left$(lblNumInsc.Caption, 1)
nSetor1 = Mid$(lblNumInsc.Caption, 3, 2)
nQuadra1 = Mid$(lblNumInsc.Caption, 6, 4)
nLote1 = Mid$(lblNumInsc.Caption, 11, 5)
nUnidade1 = Mid$(lblNumInsc.Caption, 20, 2)

nDist2 = Left$(lblNumInsc2.Caption, 1)
nSetor2 = Mid$(lblNumInsc2.Caption, 3, 2)
nQuadra2 = Mid$(lblNumInsc2.Caption, 6, 4)
nLote2 = Mid$(lblNumInsc.Caption, 11, 5)
nUnidade2 = Mid$(lblNumInsc.Caption, 20, 2)

If (nUnidade1 > 0 And nUnidade2 = 0) Or (nUnidade1 = 0 And nUnidade2 > 0) Then
    MsgBox "Não é possivel unificar um lote normal com uma subunidade de um condomínio.", vbExclamation, "Atenção"
End If

If frAtivo < 3 Then
     HabiltaFrame frAtivo + 1
End If
End Sub

Private Sub cmdPFrame_Click()
If frAtivo > 0 Then
     HabiltaFrame frAtivo - 1
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim x As Integer
Dim itmX As ListItem
Dim nCodImovel As Long

bResize = True
If Val(CodImovel) > 0 Then
    lblOK.Caption = ""
    Ocupado
    nCodImovel = CLng(Left$(CodImovel, 7))
    With xImovel
      .CarregaImovel nCodImovel
       If .Inativo Then
          MsgBox "Este imóvel encontra-se Inativo, verifique o Histórico do Imóvel para maiores informações.", vbInformation, "Atenção"
          If nLoteSel = 1 Then
               lblNumInsc.Caption = ""
               lblProp.Caption = ""
          Else
               lblNumInsc2.Caption = ""
               lblProp2.Caption = ""
          End If
          CodImovel = 0
          Exit Sub
       End If
       If nLoteSel = 1 Then
            lblNumInsc.Caption = .Inscricao & " (" & Format(.CodigoImovel, "0000000") & "-" & RetornaDVCodReduzido(.CodigoImovel) & ")"
            lblProp.Caption = .NomePropPrincipal
            tvLote.Nodes("LOTE").text = "Lote nº 1: " & Format(.Lote, "00000")
            tvLote.Nodes("FACE").text = "Face de Quadra: " & Format(.Seq, "00")
            tvLote.Nodes("FACE").Tag = .Seq
            tvLote.Nodes("PROP").text = "Proprietário: " & Format(.CodPropPrincipal, "00000") & " - " & .NomePropPrincipal
            tvLote.Nodes("PROP").Tag = .CodPropPrincipal
            tvLote.Nodes("LOG").text = "Logradouro: " & Format(.CodLogr, "00000") & " - " & Trim$(.AbrevTipoLog) & " " & IIf(IsNull(.AbrevTitLog), "", Trim$(.AbrevTitLog) & " ") & .NomeLogradouro
            tvLote.Nodes("LOG").Tag = .CodLogr
            tvLote.Nodes("NUM").text = "Número: " & .Li_Num
            tvLote.Nodes("NUM").Tag = .Li_Num
            tvLote.Nodes("COM").text = "Complemento: " & .Li_Compl
            tvLote.Nodes("BAI").text = "Bairro: " & Format(.Li_CodBairro, "000") & " - " & .DescBairro
            tvLote.Nodes("BAI").Tag = .Li_CodBairro
            tvLote.Nodes("QDS").text = "Quadras: " & .Li_Quadras
            tvLote.Nodes("LTS").text = "Lotes: " & .Li_Lotes
            tvLote.Nodes("AREA").text = "Área do Terreno: " & FormatNumber(.Dt_AreaTerreno, 2) & " m²"
            tvLote.Nodes("AREA").Tag = .Dt_AreaTerreno
            tvLote.Nodes("USO").text = "Uso do Terreno: " & Format(.Dt_CodUsoTerreno, "00") & " - " & .DescUsoTerreno
            tvLote.Nodes("USO").Tag = .Dt_CodUsoTerreno
            tvLote.Nodes("BENF").text = "Benfeitoria: " & Format(.Dt_CodBenf, "00") & " - " & .DescBenfeitoria
            tvLote.Nodes("BENF").Tag = .Dt_CodBenf
            tvLote.Nodes("TOPO").text = "Topografia: " & Format(.Dt_CodTopog, "00") & " - " & .DescTopografia
            tvLote.Nodes("TOPO").Tag = .Dt_CodTopog
            tvLote.Nodes("CATP").text = "Categoria: " & Format(.Dt_CodCategProp, "00") & " - " & .DescCategProp
            tvLote.Nodes("CATP").Tag = .Dt_CodCategProp
            tvLote.Nodes("SITU").text = "Situação: " & Format(.Dt_CodSituacao, "00") & " - " & .DescSituacao
            tvLote.Nodes("SITU").Tag = .Dt_CodSituacao
            tvLote.Nodes("PEDO").text = "Pedologia: " & Format(.Dt_CodPedol, "00") & " - " & .DescPedologia
            tvLote.Nodes("PEDO").Tag = .Dt_CodPedol
            tvLote.Nodes("TEST").text = "Testadas: "
            For x = 1 To tvLote.Nodes.Count
                  tvLote.Nodes(x).EnsureVisible
            Next
            tvLote.Nodes("LOTE").EnsureVisible
       Else
            lblNumInsc2.Caption = .Inscricao & " (" & Format(.CodigoImovel, "0000000") & "-" & RetornaDVCodReduzido(.CodigoImovel) & ")"
            If lblNumInsc2.Caption = lblNumInsc.Caption Then
               MsgBox "Não é possível unificar 2 lotes com a mesma inscrição cadastral.", vbCritical, "ERRO!!!"
               lblNumInsc2.Caption = ""
               Liberado
               bExec = True
               CodImovel = 0
               Exit Sub
            End If
            lblProp2.Caption = .NomePropPrincipal
            tvLote2.Nodes("LOTE").text = "Lote nº 2: " & Format(.Lote, "00000")
            tvLote2.Nodes("FACE").text = "Face de Quadra: " & Format(.Seq, "00")
            tvLote2.Nodes("FACE").Tag = .Seq
            tvLote2.Nodes("PROP").text = "Proprietário: " & Format(.CodPropPrincipal, "00000") & " - " & .NomePropPrincipal
            tvLote2.Nodes("PROP").Tag = .CodPropPrincipal
            tvLote2.Nodes("LOG").text = "Logradouro: " & Format(.CodLogr, "00000") & " - " & Trim$(.AbrevTipoLog) & " " & IIf(IsNull(.AbrevTitLog), "", Trim$(.AbrevTitLog) & " ") & .NomeLogradouro
            tvLote2.Nodes("LOG").Tag = .CodLogr
            tvLote2.Nodes("NUM").text = "Número: " & .Li_Num
            tvLote2.Nodes("NUM").Tag = .Li_Num
            tvLote2.Nodes("COM").text = "Complemento: " & SubNull(.Li_Compl)
            tvLote2.Nodes("BAI").text = "Bairro: " & Format(.Li_CodBairro, "000") & " - " & .DescBairro
            tvLote2.Nodes("BAI").Tag = .Li_CodBairro
            tvLote2.Nodes("QDS").text = "Quadras: " & SubNull(.Li_Quadras)
            tvLote2.Nodes("LTS").text = "Lotes: " & SubNull(.Li_Lotes)
            tvLote2.Nodes("AREA").text = "Área do Terreno: " & FormatNumber(.Dt_AreaTerreno, 2) & " m²"
            tvLote2.Nodes("AREA").Tag = .Dt_AreaTerreno
            tvLote2.Nodes("USO").text = "Uso do Terreno: " & Format(.Dt_CodUsoTerreno, "00") & " - " & .DescUsoTerreno
            tvLote2.Nodes("USO").Tag = .Dt_CodUsoTerreno
            tvLote2.Nodes("BENF").text = "Benfeitoria: " & Format(.Dt_CodBenf, "00") & " - " & .DescBenfeitoria
            tvLote2.Nodes("BENF").Tag = .Dt_CodBenf
            tvLote2.Nodes("TOPO").text = "Topografia: " & Format(.Dt_CodTopog, "00") & " - " & .DescTopografia
            tvLote2.Nodes("TOPO").Tag = .Dt_CodTopog
            tvLote2.Nodes("CATP").text = "Categoria: " & Format(.Dt_CodCategProp, "00") & " - " & .DescCategProp
            tvLote2.Nodes("CATP").Tag = .Dt_CodCategProp
            tvLote2.Nodes("SITU").text = "Situação: " & Format(.Dt_CodSituacao, "00") & " - " & .DescSituacao
            tvLote2.Nodes("SITU").Tag = .Dt_CodSituacao
            tvLote2.Nodes("PEDO").text = "Pedologia: " & Format(.Dt_CodPedol, "00") & " - " & .DescPedologia
            tvLote2.Nodes("PEDO").Tag = .Dt_CodPedol
            For x = 1 To tvLote2.Nodes.Count
                  tvLote2.Nodes(x).EnsureVisible
            Next
            tvLote2.Nodes("LOTE").EnsureVisible
       End If
       If .Li_CodBairro <> 999 Then
            lblBairro.Caption = Format(.Li_CodBairro, "000") & " - " & .DescBairro
       Else
            lblBairro.Caption = ""
       End If
       If tvLote.Nodes("AREA").Tag <> "" And tvLote2.Nodes("AREA").Tag <> "" Then
            txtAreaTerreno.text = FormatNumber(CDbl(tvLote.Nodes("AREA").Tag) + CDbl(tvLote2.Nodes("AREA").Tag), 2)
       End If
   
       If nLoteSel = 1 Then
            grdTes1.Rows = 0
       Else
            grdTes2.Rows = 0
       End If
      .CarregaTestada
       For x = 1 To .QtdeTestada
           If nLoteSel = 1 Then
              grdTes1.AddItem Format(.Testada(x, 1), "00") & Chr(9) & FormatNumber(.Testada(x, 2), 2)
           Else
              grdTes2.AddItem Format(.Testada(x, 1), "00") & Chr(9) & FormatNumber(.Testada(x, 2), 2)
           End If
       Next
       cmdAddCid.Enabled = True
       cmdDelCid.Enabled = True
    End With
    Set xImovel = Nothing
End If

Inicio:
For i = 1 To lvArea.ListItems.Count
    lvArea.ListItems.Remove (i)
    GoTo Inicio
Next

Sql = "SELECT AREAS.SEQAREA,AREAS.TIPOAREA,AREAS.DATAAPROVA,AREAS.AREACONSTR,"
Sql = Sql & "AREAS.USOCONSTR,USOCONSTR.DESCUSOCONSTR,AREAS.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
Sql = Sql & "AREAS.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM AREAS INNER JOIN USOCONSTR ON "
Sql = Sql & "AREAS.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
Sql = Sql & "AREAS.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
Sql = Sql & "AREAS.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
Sql = Sql & "WHERE CODREDUZIDO=" & Val(Mid$(Right$(lblNumInsc.Caption, 11), 2, 7))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
      Do Until .EOF
           Set itmX = lvArea.ListItems.Add(, , "1")
            itmX.SubItems(1) = Format(!SEQAREA, "00")
            itmX.SubItems(2) = !TIPOAREA
            itmX.SubItems(3) = FormatNumber(!AREACONSTR, 2) & " m²"
            itmX.SubItems(4) = Format(!DATAAPROVA, "dd/mm/yyyy")
            itmX.SubItems(5) = Format(!USOCONSTR, "00") & " - " & !DESCUSOCONSTR
            itmX.SubItems(6) = Format(!TIPOCONSTR, "00") & " - " & !DESCTIPOCONSTR
            itmX.SubItems(7) = Format(!CATCONSTR, "00") & " - " & !DESCCATEGCONSTR
          .MoveNext
      Loop
End With

Sql = "SELECT AREAS.SEQAREA,AREAS.TIPOAREA,AREAS.DATAAPROVA,AREAS.AREACONSTR,"
Sql = Sql & "AREAS.USOCONSTR,USOCONSTR.DESCUSOCONSTR,AREAS.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
Sql = Sql & "AREAS.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM AREAS INNER JOIN USOCONSTR ON "
Sql = Sql & "AREAS.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
Sql = Sql & "AREAS.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
Sql = Sql & "AREAS.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
Sql = Sql & "WHERE CODREDUZIDO=" & Val(Mid(Right$(lblNumInsc2.Caption, 11), 2, 7))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
      Do Until .EOF
           Set itmX = lvArea.ListItems.Add(, , "2")
            itmX.SubItems(1) = Format(!SEQAREA, "00")
            itmX.SubItems(2) = !TIPOAREA
            itmX.SubItems(3) = FormatNumber(!AREACONSTR, 2) & " m²"
            itmX.SubItems(4) = Format(!DATAAPROVA, "dd/mm/yyyy")
            itmX.SubItems(5) = Format(!USOCONSTR, "00") & " - " & !DESCUSOCONSTR
            itmX.SubItems(6) = Format(!TIPOCONSTR, "00") & " - " & !DESCTIPOCONSTR
            itmX.SubItems(7) = Format(!CATCONSTR, "00") & " - " & !DESCCATEGCONSTR
          .MoveNext
      Loop
End With
lblQtdeEdif.Caption = Format(lvArea.ListItems.Count, "00")

If Trim$(lblNumInsc.Caption) = "" Or Trim$(lblNumInsc2.Caption) = "" Then
    cmdCnsImovel.Enabled = True
    cmdCnsImovel2.Enabled = True
Else
    cmdCnsImovel.Enabled = False
    cmdCnsImovel2.Enabled = False
End If

Liberado
bExec = True
CodImovel = 0
End Sub

Private Sub Form_Load()
frmMdi.AddWindow Me.Name, Me.Caption
Ocupado
HabiltaFrame 0
CarregaCombo
Set xImovel = New clsImovel
Centraliza Me
cmdAddCid.Enabled = False
cmdDelCid.Enabled = False
BuildTreeProp
BuildTreeLote
Liberado

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
CodImovel = 0
Set xImovel = Nothing
End Sub

Private Sub HabiltaFrame(NumFrame As Integer)

FraD(NumFrame).Visible = True
For x = 0 To 3
      If NumFrame <> x Then
           FraD(x).Visible = False
      End If
Next
frAtivo = NumFrame

If NumFrame = 0 Then
     cmdPFrame.Enabled = False
     cmdNFrame.Enabled = True
ElseIf NumFrame = 3 Then
     cmdPFrame.Enabled = True
     cmdNFrame.Enabled = False
Else
     cmdPFrame.Enabled = True
     cmdNFrame.Enabled = True
End If

If NumFrame = 1 Then
     txtNumFace.SetFocus
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
       cmbSit.ItemData(cmbSit.NewIndex) = !CODSITUACAO
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

Private Sub lvArea_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim x As Integer

If Item.ListSubItems(2).text = "C" Then
     Item.Checked = False
     MsgBox "Só é possível selecionar uma Área Principal dentre as áreas Principais.", vbExclamation, "Atenção"
     Exit Sub
End If

With lvArea
    For x = 1 To .ListItems.Count
          If .ListItems(x).Index <> Item.Index Then
               .ListItems(x).Checked = False
          End If
    Next
End With



End Sub

Private Sub tvLote_DblClick()
Dim x As Integer

Select Case tvLote.SelectedItem.Key
      Case "FACE"
               txtNumFace.text = tvLote.SelectedItem.Tag
               txtNumFace_LostFocus
      Case "NUM"
               txtNum.text = tvLote.SelectedItem.Tag
               txtNum_LostFocus
      Case "QDS"
               txtQuadras.text = Mid(tvLote.SelectedItem.text, 10, Len(tvLote.SelectedItem.text) - 9)
      Case "LTS"
               txtLotes.text = Mid(tvLote.SelectedItem.text, 8, Len(tvLote.SelectedItem.text) - 7)
      Case "USO"
               For x = 0 To cmbUso.ListCount - 1
                     cmbUso.ListIndex = x
                     If cmbUso.ItemData(cmbUso.ListIndex) = tvLote.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "BENF"
               For x = 0 To cmbBenf.ListCount - 1
                     cmbBenf.ListIndex = x
                     If cmbBenf.ItemData(cmbBenf.ListIndex) = tvLote.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "TOPO"
               For x = 0 To cmbTopog.ListCount - 1
                     cmbTopog.ListIndex = x
                     If cmbTopog.ItemData(cmbTopog.ListIndex) = tvLote.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "CATP"
               For x = 0 To cmbCatProp.ListCount - 1
                     cmbCatProp.ListIndex = x
                     If cmbCatProp.ItemData(cmbCatProp.ListIndex) = tvLote.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "SITU"
               For x = 0 To cmbSit.ListCount - 1
                     cmbSit.ListIndex = x
                     If cmbSit.ItemData(cmbSit.ListIndex) = tvLote.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "PEDO"
               For x = 0 To cmbPedol.ListCount - 1
                     cmbPedol.ListIndex = x
                     If cmbPedol.ItemData(cmbPedol.ListIndex) = tvLote.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
End Select

End Sub

Private Sub tvLote2_DblClick()

Select Case tvLote2.SelectedItem.Key
      Case "FACE"
               txtNumFace.text = tvLote2.SelectedItem.Tag
               txtNumFace_LostFocus
      Case "NUM"
               txtNum.text = tvLote2.SelectedItem.Tag
               txtNum_LostFocus
      Case "QDS"
               txtQuadras.text = Mid(tvLote2.SelectedItem.text, 10, Len(tvLote2.SelectedItem.text) - 9)
      Case "LTS"
               txtLotes.text = Mid(tvLote2.SelectedItem.text, 8, Len(tvLote2.SelectedItem.text) - 7)
      Case "USO"
               For x = 0 To cmbUso.ListCount - 1
                     cmbUso.ListIndex = x
                     If cmbUso.ItemData(cmbUso.ListIndex) = tvLote2.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "BENF"
               For x = 0 To cmbBenf.ListCount - 1
                     cmbBenf.ListIndex = x
                     If cmbBenf.ItemData(cmbBenf.ListIndex) = tvLote2.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "TOPO"
               For x = 0 To cmbTopog.ListCount - 1
                     cmbTopog.ListIndex = x
                     If cmbTopog.ItemData(cmbTopog.ListIndex) = tvLote2.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "CATP"
               For x = 0 To cmbCatProp.ListCount - 1
                     cmbCatProp.ListIndex = x
                     If cmbCatProp.ItemData(cmbCatProp.ListIndex) = tvLote2.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "SITU"
               For x = 0 To cmbSit.ListCount - 1
                     cmbSit.ListIndex = x
                     If cmbSit.ItemData(cmbSit.ListIndex) = tvLote2.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
      Case "PEDO"
               For x = 0 To cmbPedol.ListCount - 1
                     cmbPedol.ListIndex = x
                     If cmbPedol.ItemData(cmbPedol.ListIndex) = tvLote2.SelectedItem.Tag Then
                          Exit For
                     End If
               Next
End Select

End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
Tweak txtNum, KeyAscii, IntegerPositive
End Sub

Private Sub txtNum_LostFocus()
Dim sNum As String
Dim Erro As Boolean

Erro = False
If Val(txtNum.text) = 0 Then Exit Sub

If Val(txtNum.text) <> tvLote.Nodes("NUM").Tag Then
     If tvLote.Nodes("NUM").Tag > 0 Then
          If Val(txtNum.text) <> tvLote2.Nodes("NUM").Tag Then
               If tvLote2.Nodes("NUM").Tag > 0 Then
                    Erro = True
               End If
          End If
     End If
Else
     GoTo Erro
End If

Erro:
If Erro Then
    If tvLote.Nodes("NUM").text <> tvLote2.Nodes("NUM").Tag Then
        sNum = tvLote.Nodes("NUM").Tag & "," & tvLote2.Nodes("NUM").Tag
    Else
       sNum = tvLote.Nodes("NUM").Tag
    End If
    MsgBox "O Imóvel unificado só pode ter o(s) número(s) nº " & sNum, vbExclamation, "Atenção"
    txtNum.SetFocus
End If

End Sub

Private Sub txtNumFace_GotFocus()
txtNumFace.SelStart = 0
txtNumFace.SelLength = Len(txtNumFace.text)
End Sub

Private Sub txtNumFace_KeyPress(KeyAscii As Integer)
Tweak txtNumFace, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumFace_LostFocus()
Dim nFace1 As Integer, nFace2 As Integer
If Val(txtNumFace.text) = 0 Then
    txtCodLogrLI.text = 0
    txtNomeLogLI.text = ""
    Exit Sub
End If

nFace1 = Val(Right$(tvLote.Nodes("FACE").text, 2))
nFace2 = Val(Right$(tvLote2.Nodes("FACE").text, 2))
If Val(txtNumFace.text) <> nFace1 And Val(txtNumFace.text) <> nFace2 Then
     If nFace1 <> nFace2 Then
        MsgBox "O Imóvel unificado só pode ter a(s) face(s) nº " & nFace1 & " e " & nFace2, vbExclamation, "Atenção"
     Else
        MsgBox "O Imóvel unificado só pode ter a(s) face(s) nº " & nFace1, vbExclamation, "Atenção"
     End If
     txtCodLogrLI.text = 0
     txtNomeLogLI.text = ""
     txtNumFace.SetFocus
     Exit Sub
Else
    If Val(Right$(tvLote.Nodes("FACE").text, 2)) = Val(txtNumFace.text) Then
       txtCodLogrLI.text = Mid$(tvLote.Nodes("LOG").text, 13, 5)
       txtNomeLogLI.text = Mid$(tvLote.Nodes("LOG").text, 21, Len(tvLote.Nodes("LOG").text) - 20)
    Else
       txtCodLogrLI.text = Mid$(tvLote2.Nodes("LOG").text, 13, 5)
       txtNomeLogLI.text = Mid$(tvLote2.Nodes("LOG").text, 21, Len(tvLote2.Nodes("LOG").text) - 20)
    End If
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
   Set NodX = .Nodes.Add(, , "COMP", "Compromissários", 2)
End With

With tvProp
    For x = 1 To .Nodes.Count
       .Nodes(x).EnsureVisible
    Next
   .Nodes("PROP").Bold = True
   .Nodes("COMP").Bold = True
End With

End Sub

Private Sub BuildTreeLote()

With tvLote
Inicio:
    For x = 1 To .Nodes.Count
       .Nodes.Remove (x)
       GoTo Inicio:
    Next
End With

With tvLote2
Inicio2:
    For x = 1 To .Nodes.Count
       .Nodes.Remove (x)
       GoTo Inicio2:
    Next
End With

With tvLote
    .ImageList = ImlTv
    Set NodX = .Nodes.Add(, , "LOTE", "1º Lote: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "FACE", "Face de Quadra: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "PROP", "Proprietário: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "LOG", "Logradouro: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "NUM", "Número: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "COM", "Complemento: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "BAI", "Bairro: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "QDS", "Quadras: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "LTS", "Lotes: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "AREA", "Área do Terreno: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "USO", "Uso do Terreno: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "BENF", "Benfeitoria: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "TOPO", "Topografia: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "CATP", "Categoria: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "SITU", "Situação: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "PEDO", "Pedologia: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "TEST", "Testadas: ")
   .Nodes("LOTE").Bold = True
End With

With tvLote2
    .ImageList = ImlTv
    Set NodX = .Nodes.Add(, , "LOTE", "2º Lote: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "FACE", "Face de Quadra: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "PROP", "Proprietário: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "LOG", "Logradouro: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "NUM", "Número: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "COM", "Complemento: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "BAI", "Bairro: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "QDS", "Quadras: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "LTS", "Lotes: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "AREA", "Área do Terreno: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "USO", "Uso do Terreno: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "BENF", "Benfeitoria: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "TOPO", "Topografia: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "CATP", "Categoria: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "SITU", "Situação: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "PEDO", "Pedologia: ")
    Set NodX = .Nodes.Add("LOTE", tvwChild, "TEST", "Testadas: ")
   .Nodes("LOTE").Bold = True
End With

End Sub

Private Sub txtFace_GotFocus()
txtFace.SelStart = 0
txtFace.SelLength = Len(txtFace)
End Sub

Private Sub txtFace_KeyPress(KeyAscii As Integer)
Tweak txtFace, KeyAscii, IntegerPositive
End Sub

Private Sub txtTestada_GotFocus()
txtTestada.SelStart = 0
txtTestada.SelLength = Len(txtTestada)
End Sub
