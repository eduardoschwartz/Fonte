VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCadCondominio 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Condomínios"
   ClientHeight    =   5805
   ClientLeft      =   3870
   ClientTop       =   3165
   ClientWidth     =   9720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame pnlProp 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Outros dados"
      Height          =   3825
      Left            =   1770
      TabIndex        =   71
      Top             =   1500
      Visible         =   0   'False
      Width           =   5595
      Begin VB.Frame FraD 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Áreas Construidas"
         ForeColor       =   &H00000080&
         Height          =   2385
         Index           =   3
         Left            =   90
         TabIndex        =   74
         Top             =   1290
         Width           =   5400
         Begin prjChameleon.chameleonButton cmdDelArea 
            Height          =   315
            Left            =   1335
            TabIndex        =   75
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
            MICON           =   "frmCadCondominio.frx":0000
            PICN            =   "frmCadCondominio.frx":001C
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
            TabIndex        =   76
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
            MICON           =   "frmCadCondominio.frx":0176
            PICN            =   "frmCadCondominio.frx":0192
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
            TabIndex        =   77
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
            MICON           =   "frmCadCondominio.frx":02EC
            PICN            =   "frmCadCondominio.frx":0308
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
            TabIndex        =   83
            Top             =   240
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
            TabIndex        =   79
            Top             =   2055
            Width           =   285
         End
         Begin VB.Label Label24 
            BackStyle       =   0  'Transparent
            Caption         =   "Qtde de Edific..:"
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   3810
            TabIndex        =   78
            Top             =   2040
            Width           =   1140
         End
      End
      Begin VB.TextBox txtFracao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1230
         TabIndex        =   73
         Text            =   "0,00"
         Top             =   900
         Width           =   885
      End
      Begin VB.TextBox txtCodProp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1230
         TabIndex        =   72
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Fração Ideal.:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   82
         Top             =   930
         Width           =   1035
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário...:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   81
         Top             =   270
         Width           =   1065
      End
      Begin VB.Label lblNomeProp 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   150
         TabIndex        =   80
         Top             =   600
         Width           =   5265
      End
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   8280
      TabIndex        =   69
      ToolTipText     =   "Novo Registro"
      Top             =   180
      Width           =   1155
      _ExtentX        =   2037
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadCondominio.frx":0462
      PICN            =   "frmCadCondominio.frx":047E
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
      Left            =   8280
      TabIndex        =   70
      ToolTipText     =   "Sair da Tela"
      Top             =   2220
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadCondominio.frx":05D8
      PICN            =   "frmCadCondominio.frx":05F4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBusca 
      Height          =   315
      Left            =   8280
      TabIndex        =   63
      ToolTipText     =   "Buscar Imóvel"
      Top             =   180
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Buscar"
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
      MICON           =   "frmCadCondominio.frx":0662
      PICN            =   "frmCadCondominio.frx":067E
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
      Left            =   8280
      TabIndex        =   64
      ToolTipText     =   "Editar Registro"
      Top             =   510
      Width           =   1155
      _ExtentX        =   2037
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadCondominio.frx":07D8
      PICN            =   "frmCadCondominio.frx":07F4
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
      Left            =   8280
      TabIndex        =   65
      ToolTipText     =   "Excluir Registro"
      Top             =   840
      Width           =   1155
      _ExtentX        =   2037
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadCondominio.frx":094E
      PICN            =   "frmCadCondominio.frx":096A
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
      Left            =   8280
      TabIndex        =   66
      ToolTipText     =   "Gravar os Dados"
      Top             =   1890
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadCondominio.frx":0A0C
      PICN            =   "frmCadCondominio.frx":0A28
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
      Left            =   8280
      TabIndex        =   67
      ToolTipText     =   "Cancelar Edição"
      Top             =   2220
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadCondominio.frx":0DCD
      PICN            =   "frmCadCondominio.frx":0DE9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdOutro 
      Height          =   315
      Left            =   8280
      TabIndex        =   68
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   1560
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Outros"
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
      MICON           =   "frmCadCondominio.frx":0F43
      PICN            =   "frmCadCondominio.frx":0F5F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Testadas"
      ForeColor       =   &H00000080&
      Height          =   2895
      Left            =   7200
      TabIndex        =   54
      Top             =   2850
      Width           =   2475
      Begin VB.TextBox txtTestada 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   660
         TabIndex        =   22
         Text            =   "0,00"
         Top             =   2250
         Width           =   885
      End
      Begin VB.TextBox txtFace 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   660
         TabIndex        =   21
         Text            =   "0"
         Top             =   1920
         Width           =   435
      End
      Begin MSFlexGridLib.MSFlexGrid grdTestada 
         Height          =   1455
         Left            =   90
         TabIndex        =   23
         Top             =   330
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2566
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColorFixed  =   15658734
         BackColorBkg    =   -2147483633
         FocusRect       =   0
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "^Face        |^Área              "
      End
      Begin prjChameleon.chameleonButton cmdAddTestada 
         Height          =   285
         Left            =   1680
         TabIndex        =   61
         ToolTipText     =   "Adicionar Testada"
         Top             =   1920
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
         MICON           =   "frmCadCondominio.frx":10B9
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
         TabIndex        =   62
         ToolTipText     =   "Remover Testada"
         Top             =   1920
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
         MICON           =   "frmCadCondominio.frx":10D5
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
         Caption         =   "Face:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Área:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   55
         Top             =   2280
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Outros"
      ForeColor       =   &H00000080&
      Height          =   1515
      Left            =   0
      TabIndex        =   46
      Top             =   4230
      Width           =   7185
      Begin VB.TextBox txtSU 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3630
         TabIndex        =   20
         Text            =   "0"
         Top             =   870
         Width           =   555
      End
      Begin MSFlexGridLib.MSFlexGrid grdUnid 
         Height          =   1035
         Left            =   5010
         TabIndex        =   58
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1826
         _Version        =   393216
         Rows            =   1
         FixedCols       =   0
         BackColorFixed  =   15658734
         BackColorBkg    =   12632256
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "^Unid  |^SubUnidades  "
      End
      Begin VB.TextBox txtNumUnid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   990
         Width           =   1185
      End
      Begin VB.TextBox txtAreaCon 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   18
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox txtAreaTer 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   17
         Top             =   270
         Width           =   1185
      End
      Begin prjChameleon.chameleonButton cmdEditSU 
         Height          =   285
         Left            =   4260
         TabIndex        =   60
         ToolTipText     =   "Sair da Tela"
         Top             =   840
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "-->"
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
         MICON           =   "frmCadCondominio.frx":10F1
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
         Caption         =   "Digite o nº de Sub- Unidades por Unidade"
         Height          =   405
         Index           =   19
         Left            =   3270
         TabIndex        =   59
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Unidades.........:"
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   49
         Top             =   1050
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Área do Terreno.........:"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   48
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Área Total Construída:"
         Height          =   255
         Index           =   17
         Left            =   120
         TabIndex        =   47
         Top             =   690
         Width           =   1605
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Dados do Terreno"
      ForeColor       =   &H00000080&
      Height          =   1365
      Left            =   0
      TabIndex        =   39
      Top             =   2850
      Width           =   7185
      Begin VB.ComboBox cmbUso 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2085
      End
      Begin VB.ComboBox cmbCat 
         Height          =   315
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   2085
      End
      Begin VB.ComboBox cmbTop 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   2085
      End
      Begin VB.ComboBox cmbBen 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2085
      End
      Begin VB.ComboBox cmbPed 
         Height          =   315
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   960
         Width           =   2085
      End
      Begin VB.ComboBox cmbSit 
         Height          =   315
         Left            =   4950
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   600
         Width           =   2085
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação..:"
         Height          =   255
         Index           =   10
         Left            =   3630
         TabIndex        =   45
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Categoria Prop..:"
         Height          =   255
         Index           =   11
         Left            =   3630
         TabIndex        =   44
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Topografia..:"
         Height          =   255
         Index           =   12
         Left            =   90
         TabIndex        =   43
         Top             =   1020
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Benfeitoria..:"
         Height          =   255
         Index           =   13
         Left            =   90
         TabIndex        =   42
         Top             =   645
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Uso do Terreno.:"
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   41
         Top             =   285
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedologia..:"
         Height          =   255
         Index           =   16
         Left            =   3630
         TabIndex        =   40
         Top             =   1020
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Local do Imóvel"
      ForeColor       =   &H00000080&
      Height          =   1755
      Left            =   0
      TabIndex        =   31
      Top             =   1080
      Width           =   8115
      Begin VB.TextBox txtLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   5595
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7380
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   7
         Top             =   600
         Width           =   4965
      End
      Begin VB.ComboBox cmbBairro 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   2925
      End
      Begin VB.TextBox txtQuadras 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1140
         TabIndex        =   9
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtLotes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   10
         Top             =   1320
         Width           =   2835
      End
      Begin VB.Label lblCep 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   6750
         TabIndex        =   57
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lotes.............:"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   38
         Top             =   1380
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadras........:"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   37
         Top             =   1380
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP..:"
         Height          =   255
         Index           =   3
         Left            =   6180
         TabIndex        =   36
         Top             =   660
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento:"
         Height          =   255
         Index           =   4
         Left            =   90
         TabIndex        =   35
         Top             =   660
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº....:"
         Height          =   255
         Index           =   5
         Left            =   6840
         TabIndex        =   34
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro............:"
         Height          =   255
         Index           =   22
         Left            =   90
         TabIndex        =   33
         Top             =   1020
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro...:"
         Height          =   255
         Index           =   23
         Left            =   90
         TabIndex        =   32
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Nome e Inscrição Cadastral"
      ForeColor       =   &H00000080&
      Height          =   1065
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   8115
      Begin VB.TextBox txtCond 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         MaxLength       =   40
         TabIndex        =   0
         Top             =   270
         Width           =   5595
      End
      Begin VB.ComboBox cmbCond 
         Height          =   315
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   270
         Width           =   5595
      End
      Begin VB.TextBox txtSetor 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2670
         MaxLength       =   2
         TabIndex        =   2
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtDist 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         MaxLength       =   2
         TabIndex        =   1
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtCodCond 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7290
         TabIndex        =   51
         Top             =   270
         Width           =   675
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4470
         MaxLength       =   4
         TabIndex        =   3
         Top             =   630
         Width           =   735
      End
      Begin VB.TextBox txtLote 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5850
         MaxLength       =   5
         TabIndex        =   4
         Top             =   630
         Width           =   915
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7290
         MaxLength       =   2
         TabIndex        =   5
         Top             =   630
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cód.:"
         Height          =   255
         Index           =   21
         Left            =   6840
         TabIndex        =   52
         Top             =   330
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Condomínio..:"
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   50
         Top             =   330
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito..........:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   690
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Face:"
         Height          =   255
         Index           =   6
         Left            =   6840
         TabIndex        =   29
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         Height          =   255
         Index           =   7
         Left            =   5370
         TabIndex        =   28
         Top             =   690
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra:"
         Height          =   255
         Index           =   8
         Left            =   3780
         TabIndex        =   27
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor:"
         Height          =   255
         Index           =   9
         Left            =   2130
         TabIndex        =   26
         Top             =   690
         Width           =   465
      End
   End
   Begin MSComctlLib.ImageList ImlTv 
      Left            =   0
      Top             =   0
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
            Picture         =   "frmCadCondominio.frx":110D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":1269
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":13C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":1525
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":1685
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":17E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":193D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":1A99
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadCondominio.frx":1DB5
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCadCondominio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim bExec As Boolean
Dim nOldCod As Integer
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean
Dim xImovel As New clsImovel, bResize As Boolean

Private Sub cmbCond_Click()
If bExec Then
    Limpa
    txtCodCond.Text = Format(cmbCond.ItemData(cmbCond.ListIndex), "0000")
    Le
End If
End Sub

Private Sub cmdAddArea_Click()

Set frm = frmAreas
frm.sForm = Me.Name
frm.sEvento = "Novo"
frm.show 1
lblQtdeEdif.Caption = lvArea.ListItems.Count


End Sub

Private Sub cmdAlterar_Click()
If cmbCond.ListIndex = -1 Then
    MsgBox "Selecione o condominio.", vbExclamation, "Atenção"
    Exit Sub
End If
cmbCond.Visible = False
txtCond.Visible = True
Evento = "Alterar"
Eventos "INCLUIR"
txtCond.SetFocus
End Sub

Private Sub cmdBusca_Click()

Dim z As Variant

z = InputBox("Digite o Código do Imóvel que deseja importar os dados.", "Importação de Dados de um Imóvel")
If z = "" Then Exit Sub
If Val(z) = 0 Then
    MsgBox "Imóvel Inválido.", vbCritical, "Atenção"
Else
    CarregaImovel CLng(z)
End If

End Sub

Private Sub CarregaImovel(nCodReduz As Long)
Limpa
With xImovel
    .CarregaImovel nCodReduz
    If Val(.CodigoImovel) = 0 Then
        MsgBox "Imóvel não cadastrado.", vbExclamation, "Atenção"
        Limpa
        Exit Sub
    End If
    txtDist.Text = .Distrito
    txtSetor.Text = .Setor
    txtQuadra.Text = .Quadra
    txtLote.Text = .Lote
    txtSeq.Text = .Seq
    txtSeq_Change
    txtNum.Text = .Li_Num
    txtCompl.Text = .Li_Compl
    lblCEP.Caption = RetornaCEP(.CodLogr, .Li_Num)
    If .Li_CodBairro <> 999 Then
       For x = 0 To cmbBairro.ListCount - 1
           cmbBairro.ListIndex = x
           If cmbBairro.ItemData(cmbBairro.ListIndex) = .Li_CodBairro Then
              Exit For
           End If
       Next
    End If
    txtQuadras.Text = .Li_Quadras
    txtLotes.Text = .Li_Lotes
    If Not IsNull(.Dt_CodUsoTerreno) Then
        For x = 0 To cmbUso.ListCount - 1
            cmbUso.ListIndex = x
            If cmbUso.ItemData(cmbUso.ListIndex) = .Dt_CodUsoTerreno Then
               Exit For
            End If
        Next
    End If
    If Not IsNull(.Dt_CodBenf) Then
        For x = 0 To cmbBen.ListCount - 1
            cmbBen.ListIndex = x
            If cmbBen.ItemData(cmbBen.ListIndex) = .Dt_CodBenf Then
               Exit For
            End If
        Next
    End If
    If Not IsNull(.Dt_CodTopog) Then
        For x = 0 To cmbTop.ListCount - 1
            cmbTop.ListIndex = x
            If cmbTop.ItemData(cmbTop.ListIndex) = .Dt_CodTopog Then
               Exit For
            End If
        Next
    End If
    If Not IsNull(.Dt_CodCategProp) Then
        For x = 0 To cmbCat.ListCount - 1
            cmbCat.ListIndex = x
            If cmbCat.ItemData(cmbCat.ListIndex) = .Dt_CodCategProp Then
               Exit For
            End If
        Next
    End If
    If Not IsNull(.Dt_CodSituacao) Then
        For x = 0 To cmbSit.ListCount - 1
            cmbSit.ListIndex = x
            If cmbSit.ItemData(cmbSit.ListIndex) = .Dt_CodSituacao Then
               Exit For
            End If
        Next
    End If
    If Not IsNull(.Dt_CodPedol) Then
        For x = 0 To cmbPed.ListCount - 1
            cmbPed.ListIndex = x
            If cmbPed.ItemData(cmbPed.ListIndex) = .Dt_CodPedol Then
               Exit For
            End If
        Next
    End If
    txtAreaTer.Text = FormatNumber(.Dt_AreaTerreno, 2)
   .CarregaTestada
    For x = 1 To .QtdeTestada
       If Val(.Testada(x, 1)) = Val(txtSeq.Text) Then
          grdTestada.AddItem Format(.Testada(x, 1), "00") & Chr(9) & FormatNumber(.Testada(x, 2), 2)
          Exit For
       End If
    Next
    For x = 1 To .QtdeTestada
       If Val(.Testada(x, 1)) <> Val(txtSeq.Text) Then
          grdTestada.AddItem Format(.Testada(x, 1), "00") & Chr(9) & FormatNumber(.Testada(x, 2), 2)
       End If
    Next
     
    
End With


End Sub

Private Sub cmdCancel_Click()
If MsgBox("Cancelar a Edição deste Condomínio ?", vbQuestion + vbYesNo + vbDefaultButton2, "Atenção") = vbYes Then
    txtCodCond.Text = nOldCod
    Limpa
    Le
    Evento = ""
    Eventos "INICIAR"
End If

End Sub

Private Sub cmdDelArea_Click()

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

Private Sub cmdEditSU_Click()

If Val(txtSU.Text) = 0 Then
     MsgBox "Digite o nº de subunidades.", vbExclamation, "Atenção"
     Exit Sub
Else
    grdUnid.TextMatrix(grdUnid.Row, 1) = txtSU.Text
End If

End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Erro

If cmbCond.ListIndex = -1 Then
    MsgBox "Selecione o condominio.", vbExclamation, "Atenção"
    Exit Sub
End If

If txtCodCond.Text = "" Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

Sql = "SELECT CODREDUZIDO,CODCONDOMINIO FROM CADIMOB WHERE CODCONDOMINIO=" & cmbCond.ItemData(cmbCond.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Não é possível excluir pois existem imóveis cadastrados para este condomínio.", vbExclamation, "Atenção"
        Exit Sub
    End If
End With

If MsgBox("Excluir este Condomínio ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM TESTADACONDOMINIO  WHERE CODCOND=" & Val(txtCodCond.Text)
   cn.Execute Sql, rdExecDirect
   Sql = "DELETE FROM CONDOMINIOUNIDADE  WHERE CD_CODIGO=" & Val(txtCodCond.Text)
   cn.Execute Sql, rdExecDirect
   Sql = "DELETE FROM CONDOMINIO  WHERE CD_CODIGO=" & Val(txtCodCond.Text)
   cn.Execute Sql, rdExecDirect
   Log Form, Me.Caption, Exclusão, "Excluído Condomínio " & Format(txtCodCond.Text, "000") & "-" & txtCond.Text
   Limpa
   cmbCond.Clear
   CarregaLista
   Le
End If
    
Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next
Resume Next
End Sub

Private Sub cmdGravar_Click()

Ocupado
If Not Valida() Then
     Liberado
     Exit Sub
End If
Grava
Evento = ""
Eventos "INICIAR"
Liberado

End Sub

Private Function Valida() As Boolean

Valida = False

If Trim$(txtCond.Text) = "" Then
     MsgBox "Digite o Nome do Condomínio.", vbExclamation, "Atenção"
     txtCond.SetFocus
     Exit Function
End If

If Trim$(lblNomeProp.Caption) = "" Then
     MsgBox "Digite o Nome do Proprietário.", vbExclamation, "Atenção"
     Exit Function
End If

If Trim$(txtLogr.Text) = "" Then
     MsgBox "Face de Quadra não cadastrada.", vbExclamation, "Atenção"
     txtSeq.SetFocus
     Exit Function
End If

If Evento = "Novo" Then
    
    Sql = "SELECT CD_CODIGO,CD_NOMECOND FROM CONDOMINIO WHERE "
    Sql = Sql & "CD_DISTRITO=" & Val(txtDist.Text) & " AND "
    Sql = Sql & "CD_SETOR=" & Val(txtSetor.Text) & " AND "
    Sql = Sql & "CD_QUADRA=" & Val(txtQuadra.Text) & " AND "
    Sql = Sql & "CD_LOTE=" & Val(txtLote.Text) & " AND "
    Sql = Sql & "CD_SEQ=" & Val(txtSeq.Text) & " AND "
    Sql = Sql & "CD_CODIGO<>" & 999
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
        MsgBox "Número de Inscrição Cadastral já existente." & vbCrLf & "(Condominio: " & RdoAux!cd_nomecond & ")", vbExclamation, "Atenção"
        Exit Function
    End If
End If

If cmbBairro.ListIndex = -1 Then
   cmbBairro.ListIndex = 0
End If

If cmbBairro.ListIndex = 0 Then
     MsgBox "Selecione o Bairro.", vbExclamation, "Atenção"
     cmbBairro.SetFocus
     Exit Function
End If

If cmbUso.ListIndex = -1 Then
     MsgBox "Selecione o Uso do Terreno.", vbExclamation, "Atenção"
     cmbUso.SetFocus
     Exit Function
End If

If cmbBen.ListIndex = -1 Then
     MsgBox "Selecione a Benfeitoria.", vbExclamation, "Atenção"
     cmbBen.SetFocus
     Exit Function
End If

If cmbTop.ListIndex = -1 Then
     MsgBox "Selecione a Topografia.", vbExclamation, "Atenção"
     cmbTop.SetFocus
     Exit Function
End If

If cmbCat.ListIndex = -1 Then
     MsgBox "Selecione a Categoria da Propriedade.", vbExclamation, "Atenção"
     cmbCat.SetFocus
     Exit Function
End If

If cmbSit.ListIndex = -1 Then
     MsgBox "Selecione a Situação.", vbExclamation, "Atenção"
     cmbSit.SetFocus
     Exit Function
End If

If cmbPed.ListIndex = -1 Then
     MsgBox "Selecione a Pedologia.", vbExclamation, "Atenção"
     cmbPed.SetFocus
     Exit Function
End If

If Val(txtAreaTer.Text) = 0 Then
     MsgBox "Digite a Área do Terreno.", vbExclamation, "Atenção"
     txtAreaTer.SetFocus
     Exit Function
End If

If Val(txtNumUnid.Text) = 0 Then
     MsgBox "Digite o nº de unidades.", vbExclamation, "Atenção"
     txtNumUnid.SetFocus
     Exit Function
End If

For x = 1 To grdUnid.Rows - 1
      If Val(grdUnid.TextMatrix(x, 1)) = 0 Then
           MsgBox "Digite as subunidades de todas as unidades.", vbExclamation, "Atenção"
           txtSU.SetFocus
           Exit Function
      End If
Next

If (cmbSit.ItemData(cmbSit.ListIndex) = 3 Or cmbSit.ItemData(cmbSit.ListIndex) = 4) And grdTestada.Rows > 1 Then
   MsgBox "O Lote Interno ou Encravado não pode ter testadas .", vbCritical, "Erro de Validação."
   Exit Function
End If

If grdTestada.Rows = 1 And (cmbSit.ItemData(cmbSit.ListIndex) <> 3 And cmbSit.ItemData(cmbSit.ListIndex) <> 4) Then
     MsgBox "Digite as Testadas.", vbExclamation, "Atenção"
     txtFace.SetFocus
     Exit Function
End If

If cmbSit.ItemData(cmbSit.ListIndex) = 2 And grdTestada.Rows <= 2 Then
   MsgBox "O Lote de esquina deve ter mais de 1 testada .", vbCritical, "Erro de Validação."
   If cmbSit.Enabled = True Then cmbSit.SetFocus
   Exit Function
End If
If cmbSit.ItemData(cmbSit.ListIndex) = 6 And grdTestada.Rows <= 3 Then
   MsgBox "O Lote de Quadra Inteira deve ter mais de 2 testadas .", vbCritical, "Erro de Validação."
   If cmbSit.Enabled = True Then cmbSit.SetFocus
   Exit Function
End If
If (cmbSit.ItemData(cmbSit.ListIndex) = 3 Or cmbSit.ItemData(cmbSit.ListIndex) = 4) And grdTestada.Rows > 1 Then
   MsgBox "O Lote Interno ou Encravado não pode ter testadas .", vbCritical, "Erro de Validação."
   If cmbSit.Enabled = True Then cmbSit.SetFocus
   Exit Function
End If

Valida = True

End Function


Private Sub cmdNovo_Click()

Limpa
Evento = "Novo"
Eventos "INCLUIR"
cmbCond.Visible = False
txtCond.Visible = True
nOldCod = Val(txtCodCond.Text)
txtCodCond.Text = ""
txtCond.Text = ""
txtCond.SetFocus

End Sub

Private Sub cmdOutro_Click()
If cmdOutro.Value = True Then
    pnlProp.Visible = True
Else
    pnlProp.Visible = False
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Liberado
bResize = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If cmdNovo.Visible = True Then
      cmdNovo_Click
   Else
      cmdGravar_Click
   End If
End If
End Sub

Private Sub Form_Load()
Ocupado
bExec = False
Centraliza Me
Set xImovel = New clsImovel
sRet = RetEventUserForm(Me.Name)
CarregaLista
Eventos "INICIAR"
bExec = True
Le

End Sub

Private Sub CarregaLista()
bExec = False
Sql = "SELECT CD_CODIGO,CD_NOMECOND FROM CONDOMINIO WHERE CD_CODIGO<>999 ORDER BY CD_NOMECOND;" & _
          "SELECT CODSITUACAO,DESCSITUACAO FROM SITUACAO WHERE CODSITUACAO<>999 ORDER BY DESCSITUACAO; " & _
          "SELECT CODBENFEITORIA,DESCBENFEITORIA FROM BENFEITORIA WHERE CODBENFEITORIA<>999 ORDER BY DESCBENFEITORIA; " & _
          "SELECT CODPEDOLOGIA,DESCPEDOLOGIA FROM PEDOLOGIA WHERE CODPEDOLOGIA<>999 ORDER BY DESCPEDOLOGIA; " & _
          "SELECT CODTOPOGRAFIA,DESCTOPOGRAFIA FROM TOPOGRAFIA WHERE CODTOPOGRAFIA<>999 ORDER BY DESCTOPOGRAFIA; " & _
          "SELECT CODUSOTERRENO,DESCUSOTERRENO FROM USOTERRENO WHERE CODUSOTERRENO<>999 ORDER BY DESCUSOTERRENO; " & _
          "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO INNER JOIN CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE INNER JOIN UF ON CIDADE.SIGLAUF = UF.SIGLAUF WHERE (UF.SIGLAUF = 'SP') AND (DESCCIDADE = 'JABOTICABAL') AND (CODBAIRRO <> 999) ORDER BY DESCBAIRRO; " & _
          "SELECT CODCATEGPROP,DESCCATEGPROP FROM CATEGPROP WHERE CODCATEGPROP<>999 ORDER BY DESCCATEGPROP"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    If .EOF Then GoTo 2
    txtCodCond.Text = !CD_CODIGO
    cmbCond.Clear
    Do Until .EOF
       cmbCond.AddItem !cd_nomecond
       cmbCond.ItemData(cmbCond.NewIndex) = !CD_CODIGO
      .MoveNext
    Loop
    If cmbCond.ListCount > 0 Then cmbCond.ListIndex = 0
2:
   .MoreResults
    Do Until .EOF
       cmbSit.AddItem !DescSituacao
       cmbSit.ItemData(cmbSit.NewIndex) = !Codsituacao
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbBen.AddItem !DescBenfeitoria
       cmbBen.ItemData(cmbBen.NewIndex) = !CODBENFEITORIA
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbPed.AddItem !DescPedologia
       cmbPed.ItemData(cmbPed.NewIndex) = !CODPEDOLOGIA
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbTop.AddItem !DescTopografia
       cmbTop.ItemData(cmbTop.NewIndex) = !CODTOPOGRAFIA
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbUso.AddItem !DescUsoTerreno
       cmbUso.ItemData(cmbUso.NewIndex) = !CODUSOTERRENO
      .MoveNext
    Loop
   .MoreResults
    cmbBairro.AddItem ""
    cmbBairro.ItemData(cmbBairro.NewIndex) = 999
    Do Until .EOF
       cmbBairro.AddItem !DescBairro
       cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
   .MoreResults
    Do Until .EOF
       cmbCat.AddItem !DescCategProp
       cmbCat.ItemData(cmbCat.NewIndex) = !CODCATEGPROP
      .MoveNext
    Loop
   .Close
End With
bExec = True
End Sub

Private Sub grdUnid_Click()

If grdUnid.Rows > 1 Then
     If grdUnid.Row > 0 Then
          txtSU.Text = grdUnid.TextMatrix(grdUnid.Row, 1)
     End If
End If
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   pnlProp.Visible = False
   cmdOutro.Value = False
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmdSair.Visible = True
   cmdBusca.Visible = False
   For Each Ct In frmCadCondominio
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
           Ct.Enabled = False
       End If
   Next
   cmdEditSU.Enabled = False
   cmdAddTestada.Enabled = False
   cmdDelTestada.Enabled = False
   cmdAddArea.Enabled = False
   cmdEditArea.Enabled = False
   cmdDelArea.Enabled = False
   
   cmbCond.Visible = True
   cmbCond.Enabled = True
   cmbCond.BackColor = vbWhite
   txtCond.Visible = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   If Evento = "Novo" Then
      cmdBusca.Visible = True
   End If
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   cmdSair.Visible = False
   For Each Ct In frmCadCondominio
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   cmdEditSU.Enabled = True
   cmdAddTestada.Enabled = True
   cmdDelTestada.Enabled = True
   cmdAddArea.Enabled = True
   cmdEditArea.Enabled = True
   cmdDelArea.Enabled = True
   txtLogr.BackColor = Kde
   txtLogr.Locked = True
   txtCodCond.BackColor = Kde
   txtCodCond.Locked = True
   If Evento = "Alterar" Then
      txtDist.Enabled = False
      txtSetor.Enabled = False
      txtQuadra.Enabled = False
      txtLote.Enabled = False
      txtSeq.Enabled = False
      txtDist.BackColor = Kde
      txtSetor.BackColor = Kde
      txtQuadra.BackColor = Kde
      txtLote.BackColor = Kde
      txtSeq.BackColor = Kde
   End If
End If

FormHagana

End Sub

Private Sub Grava()
Dim RdoAux2 As rdoResultset
Dim MaxCod As Integer, nCodCidade As Integer, nCodBairro As Integer

If Evento = "Novo" Then
    Sql = "SELECT MAX(CD_CODIGO) AS MAXIMO FROM CONDOMINIO WHERE CD_CODIGO<999"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!MAXIMO) Then
        MaxCod = 1
    Else
        MaxCod = RdoAux!MAXIMO + 1
    End If
    RdoAux.Close
Else
    MaxCod = Val(txtCodCond.Text)
End If
If cmbBairro.ListIndex > -1 Then
    nCodBairro = cmbBairro.ItemData(cmbBairro.ListIndex)
Else
    nCodBairro = 0
End If




If Evento = "Novo" Then
    Sql = "INSERT CONDOMINIO(CD_CODIGO,CD_NOMECOND,CD_DISTRITO,CD_SETOR,CD_QUADRA,CD_LOTE,CD_SEQ,CD_NUM,CD_COMPL,"
    Sql = Sql & "CD_UF,CD_CODCIDADE,CD_CODBAIRRO,CD_CEP,CD_QUADRAS,CD_LOTES,CD_AREATERRENO,CD_CODUSOTERRENO,"
    Sql = Sql & "CD_CODBENF,CD_CODTOPOG,CD_CODCATEGPROP,CD_CODSITUACAO,CD_CODPEDOL,CD_AREATOTCONSTR,CD_NUMUNID,"
    Sql = Sql & "CD_PROP,CD_FRACAO) VALUES(" & MaxCod & ",'" & Mask(txtCond.Text) & "'," & Val(txtDist.Text) & ","
    Sql = Sql & Val(txtSetor.Text) & "," & Val(txtQuadra.Text) & "," & Val(txtLote.Text) & "," & Val(txtSeq.Text) & ","
    Sql = Sql & Val(txtNum.Text) & ",'" & txtCompl.Text & "','" & "SP" & "'," & "413" & "," & IIf(nCodBairro > 0, nCodBairro, "Null") & ",'"
    Sql = Sql & lblCEP.Caption & "','" & Mask(txtQuadras.Text) & "','" & Mask(txtLotes.Text) & "'," & Virg2Ponto(RemovePonto(txtAreaTer.Text)) & ","
    Sql = Sql & cmbUso.ItemData(cmbUso.ListIndex) & "," & cmbBen.ItemData(cmbBen.ListIndex) & "," & cmbTop.ItemData(cmbTop.ListIndex) & ","
    Sql = Sql & cmbCat.ItemData(cmbCat.ListIndex) & "," & cmbSit.ItemData(cmbSit.ListIndex) & "," & cmbPed.ItemData(cmbPed.ListIndex) & ","
    Sql = Sql & Virg2Ponto(RemovePonto(txtAreaCon.Text)) & "," & Val(txtNumUnid.Text) & "," & Val(txtCodProp.Text) & "," & Virg2Ponto(RemovePonto(txtFracao.Text)) & ")"
Else
    Sql = "UPDATE CONDOMINIO SET CD_DISTRITO=" & Val(txtDist.Text) & ",CD_SETOR=" & Val(txtSetor.Text) & ",CD_QUADRA=" & Val(txtQuadra.Text) & ",CD_LOTE=" & Val(txtLote.Text) & ","
    Sql = Sql & "CD_SEQ=" & Val(txtSeq.Text) & ",CD_NOMECOND='" & Mask(txtCond.Text) & "',CD_NUM=" & Val(txtNum.Text) & ",CD_COMPL='" & txtCompl.Text & "',CD_UF='SP',"
    Sql = Sql & "CD_CODCIDADE=413,CD_CODBAIRRO=" & IIf(nCodBairro > 0, nCodBairro, "Null") & ",CD_CEP='" & lblCEP.Caption & "',CD_QUADRAS='" & Mask(txtQuadras.Text) & "',"
    Sql = Sql & "CD_LOTES='" & Mask(txtLotes.Text) & "',CD_AREATERRENO=" & Virg2Ponto(RemovePonto(txtAreaTer.Text)) & ",CD_CODUSOTERRENO=" & cmbUso.ItemData(cmbUso.ListIndex) & ","
    Sql = Sql & "CD_CODBENF=" & cmbBen.ItemData(cmbBen.ListIndex) & ",CD_CODTOPOG=" & cmbTop.ItemData(cmbTop.ListIndex) & ",CD_CODCATEGPROP=" & cmbCat.ItemData(cmbCat.ListIndex) & ",CD_CODSITUACAO=" & cmbSit.ItemData(cmbSit.ListIndex) & ","
    Sql = Sql & "CD_CODPEDOL=" & cmbPed.ItemData(cmbPed.ListIndex) & ",CD_AREATOTCONSTR=" & Virg2Ponto(RemovePonto(txtAreaCon.Text)) & ",CD_NUMUNID=" & Val(txtNumUnid.Text) & ",CD_PROP=" & Val(txtCodProp.Text) & ","
    Sql = Sql & "CD_FRACAO=" & Virg2Ponto(RemovePonto(txtFracao.Text)) & " WHERE CD_CODIGO=" & Val(txtCodCond.Text)
End If
cn.Execute Sql, rdExecDirect


Sql = "SELECT CD_CODIGO,CD_NOMECOND FROM CONDOMINIO ORDER BY CD_NOMECOND"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
bExec = False
cmbCond.Clear
Do Until RdoAux.EOF
   cmbCond.AddItem RdoAux!cd_nomecond
   cmbCond.ItemData(cmbCond.NewIndex) = RdoAux!CD_CODIGO
  RdoAux.MoveNext
Loop

If Evento = "Alterar" Then MaxCod = txtCodCond.Text
For x = 0 To cmbCond.ListCount - 1
    cmbCond.ListIndex = x
    If cmbCond.ItemData(cmbCond.ListIndex) = MaxCod Then
         Exit For
    End If
Next

'*******GRAVA AREA *******************************************

'TABELA CONDOMINIOAREA
Sql = "DELETE FROM CONDOMINIOAREA WHERE CODCONDOMINIO=" & MaxCod
cn.Execute Sql, rdExecDirect


For x = 1 To lvArea.ListItems.Count
    Sql = "INSERT CONDOMINIOAREA (CODCONDOMINIO,SEQAREA,TIPOAREA,DATAAPROVA,AREACONSTR,USOCONSTR,TIPOCONSTR,"
    Sql = Sql & "CATCONSTR,QTDEPAV) VALUES(" & MaxCod & "," & x & ",'" & "'," & IIf(IsDate(lvArea.ListItems(x).SubItems(2)), "'" & Format(lvArea.ListItems(x).SubItems(2), "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & Virg2Ponto(Left(lvArea.ListItems(x).SubItems(1), Len(lvArea.ListItems(x).SubItems(1)) - 3)) & "," & lvArea.ListItems(x).SubItems(3) & "," & lvArea.ListItems(x).SubItems(5) & ","
    Sql = Sql & lvArea.ListItems(x).SubItems(7) & "," & lvArea.ListItems(x).SubItems(9) & ")"
    cn.Execute Sql, rdExecDirect
Next


'*******GRAVA TESTADA *******************************************
Sql = "DELETE FROM TESTADACONDOMINIO WHERE CODCOND=" & MaxCod
cn.Execute Sql, rdExecDirect

For x = 1 To grdTestada.Rows - 1
    Sql = "INSERT TESTADACONDOMINIO(CODCOND,NUMFACE,AREATESTADA) VALUES("
    Sql = Sql & MaxCod & "," & Val(grdTestada.TextMatrix(x, 0)) & "," & Virg2Ponto(grdTestada.TextMatrix(x, 1)) & ")"
    cn.Execute Sql, rdExecDirect
Next

'*******GRAVA SUB UNIDADE *******************************************
Sql = "DELETE FROM CONDOMINIOUNIDADE WHERE CD_CODIGO=" & MaxCod
cn.Execute Sql, rdExecDirect

For x = 1 To grdUnid.Rows - 1
    Sql = "INSERT CONDOMINIOUNIDADE (CD_CODIGO,CD_UNIDADE,CD_SUBUNIDADES) VALUES("
    Sql = Sql & MaxCod & "," & grdUnid.TextMatrix(x, 0) & "," & grdUnid.TextMatrix(x, 1) & ")"
    cn.Execute Sql, rdExecDirect
Next

bExec = True
Limpa
Le


'*******ATUALIZA OS IMOVEIS DESTE CONDOMINIO******************
nCodCidade = 413
If Evento <> "Novo" Then
    Sql = "UPDATE CADIMOB SET DISTRITO=" & Val(txtDist.Text) & ",SETOR=" & Val(txtSetor.Text) & ",QUADRA=" & Val(txtQuadra.Text) & ",LOTE=" & Val(txtLote.Text) & ",SEQ=" & Val(txtSeq.Text) & ",LI_NUM=" & Val(txtNum.Text) & ", "
    Sql = Sql & "LI_COMPL='" & txtCompl.Text & "' , LI_CEP='" & lblCEP.Caption & "', LI_UF='SP', LI_CODCIDADE=" & nCodCidade & ","
    Sql = Sql & "LI_CODBAIRRO=" & IIf(nCodBairro = 0, Null, nCodBairro) & ",  LI_QUADRAS='" & Mask(Trim$(txtQuadras.Text)) & "', LI_LOTES='" & Mask(Trim$(txtLotes.Text)) & "',"
    Sql = Sql & "DT_AREATERRENO=" & Virg2Ponto(RemovePonto(txtAreaTer.Text)) & ", DT_CODUSOTERRENO=" & cmbUso.ItemData(cmbUso.ListIndex) & ", DT_CODBENF=" & cmbBen.ItemData(cmbBen.ListIndex) & ","
    Sql = Sql & "DT_CODTOPOG=" & cmbTop.ItemData(cmbTop.ListIndex) & ", DT_CODCATEGPROP=" & cmbCat.ItemData(cmbCat.ListIndex) & ", DT_CODSITUACAO=" & cmbSit.ItemData(cmbSit.ListIndex) & ","
    Sql = Sql & "Dt_CodPedol=" & cmbPed.ItemData(cmbPed.ListIndex) & " WHERE CodCondominio=" & cmbCond.ItemData(cmbCond.ListIndex)
    cn.Execute Sql, rdExecDirect

    '*******GRAVA TESTADA *******************************************
    Sql = "SELECT DISTINCT TESTADA.CODREDUZIDO FROM TESTADA INNER JOIN  CADIMOB ON  Testada.CODREDUZIDO = CADIMOB.CODREDUZIDO "
    Sql = Sql & "Where CADIMOB.CodCondominio =" & cmbCond.ItemData(cmbCond.ListIndex)
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Do Until .EOF
            Sql = "DELETE FROM TESTADA WHERE CODREDUZIDO=" & RdoAux2!CODREDUZIDO
            cn.Execute Sql, rdExecDirect
            For x = 1 To grdTestada.Rows - 1
                Sql = "INSERT TESTADA(CODREDUZIDO,NUMFACE,AREATESTADA) VALUES("
                Sql = Sql & RdoAux2!CODREDUZIDO & "," & Val(grdTestada.TextMatrix(x, 0)) & "," & Virg2Ponto(grdTestada.TextMatrix(x, 1)) & ")"
                cn.Execute Sql, rdExecDirect
            Next
           .MoveNext
        Loop
       .Close
    End With

End If

End Sub

Private Sub Le()
Dim itmX As ListItem, z As Long
If cmbCond.ListIndex = -1 Then Exit Sub
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)
With xImovel
     .CarregaCondominio CLng(cmbCond.ItemData(cmbCond.ListIndex))
     txtCodProp.Text = .CodProp
     lblNomeProp.Caption = .Proprietario
     txtFracao.Text = FormatNumber(.FracaoIdeal, 2)
     txtCond.Text = .NomeCondominio
     txtDist.Text = .Distrito
     txtSetor.Text = .Setor
     txtQuadra.Text = .Quadra
     txtLote.Text = .Lote
     txtSeq.Text = .Seq
     txtNum.Text = .Li_Num
     txtCompl.Text = .Li_Compl
     lblCEP.Caption = Format(.Li_Cep, "00000-000")
     If .Li_CodBairro <> 999 Then
        For x = 0 To cmbBairro.ListCount - 1
            cmbBairro.ListIndex = x
            If cmbBairro.ItemData(cmbBairro.ListIndex) = .Li_CodBairro Then
               Exit For
            End If
        Next
     End If
     txtQuadras.Text = .Li_Quadras
     txtLotes.Text = .Li_Lotes
     txtAreaTer.Text = FormatNumber(.Dt_AreaTerreno, 2)
     If Not IsNull(.Dt_CodUsoTerreno) Then
        For x = 0 To cmbUso.ListCount - 1
            cmbUso.ListIndex = x
            If cmbUso.ItemData(cmbUso.ListIndex) = .Dt_CodUsoTerreno Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodBenf) Then
        For x = 0 To cmbBen.ListCount - 1
            cmbBen.ListIndex = x
            If cmbBen.ItemData(cmbBen.ListIndex) = .Dt_CodBenf Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodTopog) Then
        For x = 0 To cmbTop.ListCount - 1
            cmbTop.ListIndex = x
            If cmbTop.ItemData(cmbTop.ListIndex) = .Dt_CodTopog Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodCategProp) Then
        For x = 0 To cmbCat.ListCount - 1
            cmbCat.ListIndex = x
            If cmbCat.ItemData(cmbCat.ListIndex) = .Dt_CodCategProp Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodSituacao) Then
        For x = 0 To cmbSit.ListCount - 1
            cmbSit.ListIndex = x
            If cmbSit.ItemData(cmbSit.ListIndex) = .Dt_CodSituacao Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodPedol) Then
        For x = 0 To cmbPed.ListCount - 1
            cmbPed.ListIndex = x
            If cmbPed.ItemData(cmbPed.ListIndex) = .Dt_CodPedol Then
               Exit For
            End If
        Next
     End If
     txtAreaCon.Text = FormatNumber(.AreaConstruida, 2)
     txtNumUnid.Text = .NumUnidades
    
    'testadas
    .CarregaTestadaCond Val(txtCodCond.Text)
     For x = 1 To .QtdeTestadaCond
       If Val(.TestadaCond(x, 1)) = Val(txtSeq.Text) Then
          grdTestada.AddItem Format(.TestadaCond(x, 1), "00") & Chr(9) & FormatNumber(.TestadaCond(x, 2), 2)
          Exit For
       End If
     Next
     For x = 1 To .QtdeTestadaCond
       If Val(.TestadaCond(x, 1)) <> Val(txtSeq.Text) Then
          grdTestada.AddItem Format(.TestadaCond(x, 1), "00") & Chr(9) & FormatNumber(.TestadaCond(x, 2), 2)
       End If
     Next
     
     If grdTestada.Rows > 1 Then
        txtFace.Text = Val(grdTestada.TextMatrix(1, 0)) + 1
     End If

    'subunidades
    Sql = "SELECT CD_UNIDADE,CD_SUBUNIDADES FROM CONDOMINIOUNIDADE "
    Sql = Sql & "WHERE CD_CODIGO=" & Val(txtCodCond.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
           grdUnid.AddItem !CD_UNIDADE & Chr(9) & !CD_SUBUNIDADES
          .MoveNext
        Loop
       .Close
    End With

End With


'CARREGA ÁREA


'Areas
Sql = "SELECT CONDOMINIOAREA.SEQAREA,CONDOMINIOAREA.QTDEPAV,CONDOMINIOAREA.TIPOAREA,CONDOMINIOAREA.DATAAPROVA,CONDOMINIOAREA.AREACONSTR,CONDOMINIOAREA.NUMPROCESSO,CONDOMINIOAREA.DATAPROCESSO,"
Sql = Sql & "CONDOMINIOAREA.USOCONSTR,USOCONSTR.DESCUSOCONSTR,CONDOMINIOAREA.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
Sql = Sql & "CONDOMINIOAREA.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM CONDOMINIOAREA INNER JOIN USOCONSTR ON "
Sql = Sql & "CONDOMINIOAREA.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
Sql = Sql & "CONDOMINIOAREA.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
Sql = Sql & "CONDOMINIOAREA.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
Sql = Sql & "WHERE CODCONDOMINIO=" & CLng(cmbCond.ItemData(cmbCond.ListIndex))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblQtdeEdif.Caption = .RowCount
    Do Until .EOF
        
       '****ListView
        Set itmX = lvArea.ListItems.Add(, "A" & Format(.AbsolutePosition, "00"), Format(.AbsolutePosition, "00"))
        itmX.SubItems(1) = FormatNumber(!AREACONSTR, 2) & " m²"
        itmX.SubItems(2) = IIf(IsNull(!DATAAPROVA), "", Format(!DATAAPROVA, "dd/mm/yyyy"))
        itmX.SubItems(3) = !USOCONSTR
        itmX.SubItems(4) = !DESCUSOCONSTR
        itmX.SubItems(5) = !TIPOCONSTR
        itmX.SubItems(6) = !DESCTIPOCONSTR
        itmX.SubItems(7) = !CATCONSTR
        itmX.SubItems(8) = !DESCCATEGCONSTR
        itmX.SubItems(9) = Val(SubNull(!QTDEPAV))
       
       
      .MoveNext
    Loop
   .Close
End With



End Sub

Private Sub Limpa()

txtCodProp.Text = ""
lblNomeProp.Caption = ""
txtFracao.Text = ""
txtDist.Text = ""
txtSetor.Text = ""
txtQuadra.Text = ""
txtLote.Text = ""
txtSeq.Text = ""
txtLogr.Text = ""
txtNum.Text = ""
txtCompl.Text = ""
cmbBairro.ListIndex = -1
lblCEP.Caption = ""
txtQuadras.Text = ""
txtLotes.Text = ""
txtAreaTer.Text = 0
cmbUso.ListIndex = -1
cmbBen.ListIndex = -1
cmbTop.ListIndex = -1
cmbCat.ListIndex = -1
cmbSit.ListIndex = -1
cmbPed.ListIndex = -1
txtAreaCon.Text = 0
txtNumUnid.Text = 0
grdTestada.Rows = 1
grdUnid.Rows = 1
txtSU.Text = 0
txtFace.Text = 0
txtTestada.Text = "0,00"

End Sub


Private Sub txtAreaCon_GotFocus()
txtAreaCon.SelStart = 0
txtAreaCon.SelLength = Len(txtAreaCon.Text)
End Sub

Private Sub txtAreaCon_KeyPress(KeyAscii As Integer)
Tweak txtAreaCon, KeyAscii, DecimalPositive
End Sub

Private Sub txtAreaTer_GotFocus()
txtAreaTer.SelStart = 0
txtAreaTer.SelLength = Len(txtAreaTer.Text)
End Sub

Private Sub txtAreaTer_KeyPress(KeyAscii As Integer)
Tweak txtAreaTer, KeyAscii, DecimalPositive
End Sub

Private Sub txtCodProp_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtCodProp_LostFocus
Else
   Tweak txtCodProp, KeyAscii, IntegerPositive
End If
End Sub

Private Sub txtCodProp_LostFocus()
Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCodProp.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Cidadão não cadastrado.", vbExclamation, "Atenção"
        lblNomeProp.Caption = ""
    Else
        lblNomeProp.Caption = !nomecidadao
    End If
   .Close
End With
End Sub

Private Sub txtDist_Change()
txtSeq_Change
End Sub

Private Sub txtDist_GotFocus()
txtDist.SelStart = 0
txtDist.SelLength = Len(txtDist.Text)
End Sub

Private Sub txtFracao_KeyPress(KeyAscii As Integer)
Tweak txtCodProp, KeyAscii, DecimalPositive
End Sub

Private Sub txtNum_LostFocus()
If Val(txtNum.Text) > 10000 Then
    MsgBox "Nº inválido.", vbExclamation, "Atenção"
    txtNum.SetFocus
    Exit Sub
End If

lblCEP.Caption = ""
If Val(Left$(txtLogr.Text, 4)) > 0 Then
     lblCEP.Caption = RetornaCEP(Val(Left$(txtLogr.Text, 4)), Val(txtNum.Text))
End If

End Sub

Private Sub txtNumUnid_LostFocus()

grdUnid.Rows = Val(txtNumUnid.Text) + 1
For x = 1 To grdUnid.Rows - 1
      grdUnid.TextMatrix(x, 0) = x
      If Val(grdUnid.TextMatrix(x, 1)) = 0 Then
           grdUnid.TextMatrix(x, 1) = 0
      End If
Next
If grdUnid.Rows > 2 Then grdUnid.Row = 1

End Sub

Private Sub txtSeq_Change()

If Val(txtDist.Text) > 0 And Val(txtSetor.Text) > 0 And Val(txtQuadra.Text) > 0 And Val(txtLote.Text) > 0 And Val(txtSeq.Text) > 0 Then
    Sql = "SELECT ABREVTIPOLOG, ABREVTITLOG, NOMELOGRADOURO, CODLOGR "
    Sql = Sql & "FROM vwFACEQUADRA "
    Sql = Sql & "WHERE CODDISTRITO=" & Val(txtDist.Text) & " AND "
    Sql = Sql & "CODSETOR=" & Val(txtSetor.Text) & " AND "
    Sql = Sql & "CODQUADRA=" & Val(txtQuadra.Text) & " AND "
    Sql = Sql & "CODFACE=" & Val(txtSeq.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
         txtLogr.Text = Format(RdoAux!CodLogr, "0000") & " - " & Trim$(RdoAux!AbrevTipoLog) & IIf(IsNull(RdoAux!AbrevTitLog), "", " " & Trim$(SubNull(RdoAux!AbrevTitLog))) & " " & RdoAux!NomeLogradouro
    Else
        MsgBox "Face de Quadra não cadastrada.", vbExclamation, "Atenção"
        txtLogr.Text = ""
    End If
Else
   txtLogr.Text = ""
End If

End Sub

Private Sub txtLote_Change()
txtSeq_Change
End Sub

Private Sub txtLote_GotFocus()
txtLote.SelStart = 0
txtLote.SelLength = Len(txtLote.Text)
End Sub

Private Sub txtNum_GotFocus()
txtNum.SelStart = 0
txtNum.SelLength = Len(txtNum.Text)
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
Tweak txtNum, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumUnid_GotFocus()
txtNumUnid.SelStart = 0
txtNumUnid.SelLength = Len(txtNumUnid.Text)
End Sub

Private Sub txtNumUnid_KeyPress(KeyAscii As Integer)
Tweak txtNumUnid, KeyAscii, IntegerPositive
End Sub

Private Sub txtQuadra_Change()
txtSeq_Change
End Sub

Private Sub txtQuadra_GotFocus()
txtQuadra.SelStart = 0
txtQuadra.SelLength = Len(txtQuadra.Text)
End Sub

Private Sub txtSetor_Change()
txtSeq_Change
End Sub

Private Sub txtSetor_GotFocus()
txtSetor.SelStart = 0
txtSetor.SelLength = Len(txtSetor.Text)
End Sub

Private Sub txtFace_GotFocus()
txtFace.SelStart = 0
txtFace.SelLength = Len(txtFace)
End Sub

Private Sub txtFace_KeyPress(KeyAscii As Integer)
Tweak txtFace, KeyAscii, IntegerPositive
End Sub

Private Sub txtSU_GotFocus()
txtSU.SelStart = 0
txtSU.SelLength = Len(txtSU.Text)
End Sub

Private Sub txtSU_KeyPress(KeyAscii As Integer)
Tweak txtSU, KeyAscii, IntegerPositive
End Sub

Private Sub txtTestada_GotFocus()
txtTestada.SelStart = 0
txtTestada.SelLength = Len(txtTestada)
End Sub

Private Sub txtTestada_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    If Val(txtTestada.Text) = 0 Then
       MsgBox "Digite a Área da Testada.", vbExclamation, "Atenção"
    Else
       grdTestada.AddItem grdTestada.Rows & Chr(9) & FormatNumber(txtTestada, 2)
    End If
End If

Tweak txtTestada, KeyAscii, DecimalPositive
End Sub

Private Sub cmdAddTestada_Click()
Dim Achou As Boolean
If txtFace.Enabled = False Then Exit Sub
If Val(txtFace.Text) = 0 Then
   MsgBox "Digite a Face da Testada.", vbExclamation, "Atenção"
   txtFace.SetFocus
   Exit Sub
End If

If Val(txtTestada.Text) = 0 Then
   MsgBox "Digite a Área da Testada.", vbExclamation, "Atenção"
   txtTestada.SetFocus
   Exit Sub
End If

If grdTestada.Rows = 1 Then
   If Val(txtFace.Text) <> Val(txtSeq.Text) Then
       MsgBox "A 1ª testada deve ser igual a face descrita na inscrição cadastral.", vbExclamation, "Atenção"
       txtFace.SetFocus
       Exit Sub
   End If
End If


Achou = False
For x = 1 To grdTestada.Rows - 1
   If Val(grdTestada.TextMatrix(x, 0)) = Val(txtFace.Text) Then
      Achou = True
      Exit For
   End If
Next
If Achou Then
   MsgBox "Face já cadastrada.", vbExclamation, "Atenção"
   txtFace.SetFocus
   Exit Sub
End If

grdTestada.AddItem Format(txtFace.Text, "00") & Chr(9) & FormatNumber(txtTestada, 2)
txtFace.Text = Val(grdTestada.TextMatrix(grdTestada.Rows - 1, 0)) + 1
txtTestada.Text = "0,00"
txtFace.SetFocus

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

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub
