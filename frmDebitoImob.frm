VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmDebitoImob 
   BackColor       =   &H00EEEEEE&
   Caption         =   "Consulta de Débitos"
   ClientHeight    =   5895
   ClientLeft      =   7875
   ClientTop       =   4635
   ClientWidth     =   11415
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   11415
   Begin VB.Frame pnlObs 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   5085
      Left            =   1140
      TabIndex        =   5
      Top             =   660
      Width           =   8925
      Begin prjChameleon.chameleonButton cmdCancelarObs 
         Height          =   315
         Left            =   7860
         TabIndex        =   13
         ToolTipText     =   "Cancelar Edição"
         Top             =   4710
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
         MICON           =   "frmDebitoImob.frx":0000
         PICN            =   "frmDebitoImob.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtObservacao 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2085
         Left            =   90
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2490
         Width           =   8805
      End
      Begin MSComctlLib.ListView lvObserv 
         Height          =   2325
         Left            =   90
         TabIndex        =   7
         Top             =   90
         Width           =   8835
         _ExtentX        =   15584
         _ExtentY        =   4101
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   12648447
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Seq"
            Object.Width           =   1060
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuário"
            Object.Width           =   3529
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Data"
            Object.Width           =   2294
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Obs"
            Object.Width           =   7938
         EndProperty
      End
      Begin prjChameleon.chameleonButton cmdSairObs 
         Height          =   315
         Left            =   7830
         TabIndex        =   8
         ToolTipText     =   "Sair da Tela de Observação"
         Top             =   4710
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Retornar"
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
         MICON           =   "frmDebitoImob.frx":0176
         PICN            =   "frmDebitoImob.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExcluirObs 
         Height          =   315
         Left            =   2280
         TabIndex        =   9
         ToolTipText     =   "Excluir Registro"
         Top             =   4740
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "E&xcluir"
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
         MICON           =   "frmDebitoImob.frx":02EC
         PICN            =   "frmDebitoImob.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAlterarObs 
         Height          =   315
         Left            =   1200
         TabIndex        =   10
         ToolTipText     =   "Editar Registro"
         Top             =   4710
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
         MICON           =   "frmDebitoImob.frx":03AA
         PICN            =   "frmDebitoImob.frx":03C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNovoObs 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Novo Registro"
         Top             =   4710
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
         MICON           =   "frmDebitoImob.frx":0520
         PICN            =   "frmDebitoImob.frx":053C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGravarObs 
         Height          =   315
         Left            =   6750
         TabIndex        =   12
         ToolTipText     =   "Gravar os Dados"
         Top             =   4710
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
         MICON           =   "frmDebitoImob.frx":0696
         PICN            =   "frmDebitoImob.frx":06B2
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
   Begin VB.Frame frReparc 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   585
      TabIndex        =   14
      Top             =   765
      Width           =   9315
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C000&
         Caption         =   "Cálculo de:"
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   300
         TabIndex        =   17
         Top             =   3420
         Width           =   1785
         Begin VB.CheckBox chkMulta 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "Multa"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   150
            TabIndex        =   19
            Top             =   270
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkJuros 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C000&
            Caption         =   "Juros"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   930
            TabIndex        =   18
            Top             =   270
            Value           =   1  'Checked
            Width           =   735
         End
      End
      Begin VB.ComboBox cmbProc 
         Height          =   315
         Left            =   1770
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   210
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid grdOrigem 
         Height          =   1965
         Left            =   4080
         TabIndex        =   15
         Top             =   390
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3466
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColorBkg    =   12632064
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "^Código     |^Ano     |^Lanc. |^Seq  |^Parc. |^Compl. |^Vencimento      "
      End
      Begin prjChameleon.chameleonButton cmdSairRep 
         Height          =   315
         Left            =   60
         TabIndex        =   20
         ToolTipText     =   "Sair da Tela de Reparcelamento"
         Top             =   4770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Retornar"
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
         MICON           =   "frmDebitoImob.frx":0A57
         PICN            =   "frmDebitoImob.frx":0A73
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid grdDestino 
         Height          =   2355
         Left            =   3480
         TabIndex        =   21
         Top             =   2730
         Width           =   5745
         _ExtentX        =   10134
         _ExtentY        =   4154
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedCols       =   0
         BackColorBkg    =   12632064
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "^Código     |^Ano     |^Lanc.|^Seq |^Parc.|^CP.|^Vencimento    |^Pagamento      "
      End
      Begin prjChameleon.chameleonButton cmdCancelReparc 
         Height          =   315
         Left            =   1200
         TabIndex        =   22
         ToolTipText     =   "Cancela Reparcelamento da SMAR"
         Top             =   4770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "frmDebitoImob.frx":0BCD
         PICN            =   "frmDebitoImob.frx":0BE9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCalculo 
         Height          =   315
         Left            =   2340
         TabIndex        =   79
         ToolTipText     =   "Exibe o cálculo do Parcelamento"
         Top             =   4770
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cálculo"
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
         MICON           =   "frmDebitoImob.frx":0CAC
         PICN            =   "frmDebitoImob.frx":0CC8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAno 
         Height          =   225
         Left            =   4320
         TabIndex        =   80
         Top             =   5160
         Visible         =   0   'False
         Width           =   2595
      End
      Begin VB.Label lblCancel 
         BackStyle       =   0  'Transparent
         Caption         =   "CANCELADO"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   585
         Left            =   120
         TabIndex        =   45
         Top             =   1650
         Width           =   3405
      End
      Begin VB.Label lblQtde 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   44
         Top             =   1380
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de Parcelas........:"
         Height          =   225
         Index           =   8
         Left            =   300
         TabIndex        =   43
         Top             =   1380
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cód.Responsável.......:"
         Height          =   225
         Index           =   11
         Left            =   300
         TabIndex        =   42
         Top             =   2700
         Width           =   1665
      End
      Begin VB.Label lblResp 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   41
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   270
         TabIndex        =   40
         Top             =   4425
         Width           =   3255
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário Responsável...:"
         Height          =   225
         Index           =   17
         Left            =   240
         TabIndex        =   39
         Top             =   4170
         Width           =   2085
      End
      Begin VB.Label lblDataProc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   38
         Top             =   750
         Width           =   1485
      End
      Begin VB.Label lblDataParc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   37
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label lbl1venc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2070
         TabIndex        =   36
         Top             =   1710
         Width           =   1485
      End
      Begin VB.Label lblValor 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   35
         Top             =   2040
         Width           =   1485
      End
      Begin VB.Label lblPerc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   34
         Top             =   2370
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Processo.......:"
         Height          =   225
         Index           =   10
         Left            =   300
         TabIndex        =   33
         Top             =   750
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Parcelamento:"
         Height          =   225
         Index           =   9
         Left            =   300
         TabIndex        =   32
         Top             =   1050
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do 1º Vencto......:"
         Height          =   225
         Index           =   5
         Left            =   300
         TabIndex        =   31
         Top             =   1710
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor de Entrada.........:"
         Height          =   225
         Index           =   6
         Left            =   300
         TabIndex        =   30
         Top             =   2040
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "% de Entrada..............:"
         Height          =   225
         Index           =   7
         Left            =   300
         TabIndex        =   29
         Top             =   2370
         Width           =   1665
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   285
         TabIndex        =   28
         Top             =   270
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas de Destino"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   27
         Top             =   2460
         Width           =   2325
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas de Origem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   4110
         TabIndex        =   26
         Top             =   120
         Width           =   2325
      End
      Begin VB.Label lblPago 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P A G O"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   585
         Left            =   240
         TabIndex        =   25
         Top             =   1620
         Width           =   3405
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Plano de Desconto.....:"
         Height          =   225
         Index           =   12
         Left            =   300
         TabIndex        =   24
         Top             =   3030
         Width           =   1665
      End
      Begin VB.Label lblPlano 
         BackStyle       =   0  'Transparent
         Caption         =   "Sem Plano"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2040
         TabIndex        =   23
         Top             =   3030
         Width           =   1395
      End
   End
   Begin VB.Frame pnlFilter 
      BackColor       =   &H00ADE7FF&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   60
      TabIndex        =   82
      Top             =   4890
      Visible         =   0   'False
      Width           =   10155
      Begin VB.CheckBox chkUnica 
         BackColor       =   &H00ADE7FF&
         Caption         =   "Exibir parcela única"
         Height          =   285
         Left            =   7920
         TabIndex        =   97
         Top             =   60
         Width           =   1905
      End
      Begin VB.ComboBox cmbAj 
         Height          =   315
         ItemData        =   "frmDebitoImob.frx":1075
         Left            =   8490
         List            =   "frmDebitoImob.frx":1082
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   360
         Width           =   1095
      End
      Begin VB.ComboBox cmbDA 
         Height          =   315
         ItemData        =   "frmDebitoImob.frx":1099
         Left            =   6690
         List            =   "frmDebitoImob.frx":10A6
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Top             =   390
         Width           =   1095
      End
      Begin VB.ComboBox cmbSeq 
         Height          =   315
         Left            =   6690
         Style           =   2  'Dropdown List
         TabIndex        =   91
         Top             =   30
         Width           =   1095
      End
      Begin VB.ComboBox cmbSit 
         Height          =   315
         Left            =   3150
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   390
         Width           =   2775
      End
      Begin VB.ComboBox cmbLanc 
         Height          =   315
         Left            =   3150
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   30
         Width           =   2775
      End
      Begin VB.ComboBox cmbAno2 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   390
         Width           =   1095
      End
      Begin VB.ComboBox cmbAno1 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   30
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuiz..:"
         Height          =   225
         Index           =   6
         Left            =   7920
         TabIndex        =   96
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Div.A..:"
         Height          =   225
         Index           =   5
         Left            =   6060
         TabIndex        =   94
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Seq....:"
         Height          =   225
         Index           =   4
         Left            =   6060
         TabIndex        =   93
         Top             =   90
         Width           =   705
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação........:"
         Height          =   225
         Index           =   3
         Left            =   2040
         TabIndex        =   90
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Lancamento..:"
         Height          =   225
         Index           =   2
         Left            =   2040
         TabIndex        =   89
         Top             =   90
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano Até.:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   86
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano de..:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   85
         Top             =   90
         Width           =   705
      End
   End
   Begin VB.Frame pnlInativo 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   8325
      TabIndex        =   77
      Top             =   -45
      Visible         =   0   'False
      Width           =   1335
      Begin VB.Label pnlInativo2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "INATIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   405
         Left            =   45
         TabIndex        =   78
         Top             =   90
         Width           =   1260
      End
   End
   Begin VB.Frame frTop 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   675
      Left            =   30
      TabIndex        =   67
      Top             =   0
      Width           =   11355
      Begin VB.CheckBox chkShowAll 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Todos os lançamentos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6300
         TabIndex        =   2
         Top             =   90
         Width           =   1980
      End
      Begin VB.CheckBox chkTodosAnos 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Todos os exercícios"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4410
         TabIndex        =   1
         Top             =   90
         Width           =   1755
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
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
         Height          =   195
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   98
         Top             =   405
         Width           =   4470
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1920
         MaxLength       =   6
         TabIndex        =   0
         Top             =   60
         Width           =   945
      End
      Begin prjChameleon.chameleonButton cmdCnsImovel 
         Height          =   345
         Left            =   2940
         TabIndex        =   68
         ToolTipText     =   "Consulta Imóvel"
         Top             =   30
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
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
         MICON           =   "frmDebitoImob.frx":10BD
         PICN            =   "frmDebitoImob.frx":10D9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRefresh 
         Height          =   345
         Left            =   3420
         TabIndex        =   69
         ToolTipText     =   "Atualizar Dados"
         Top             =   30
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "!"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":1233
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image imgSerasa 
         Height          =   480
         Left            =   9765
         Picture         =   "frmDebitoImob.frx":124F
         Stretch         =   -1  'True
         Top             =   90
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.Label lblDV 
         BackColor       =   &H00EEEEEE&
         Caption         =   "DV - 0"
         Height          =   285
         Left            =   3810
         TabIndex        =   81
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lblRS 
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
         Left            =   60
         TabIndex        =   75
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço.:"
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
         Index           =   6
         Left            =   5550
         TabIndex        =   74
         Top             =   390
         Width           =   795
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
         Left            =   990
         TabIndex        =   73
         Top             =   390
         Visible         =   0   'False
         Width           =   3765
      End
      Begin VB.Label lblRua 
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
         Height          =   195
         Left            =   6420
         TabIndex        =   72
         Top             =   390
         Width           =   3885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido/I.M...:"
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
         Left            =   60
         TabIndex        =   71
         Top             =   90
         Width           =   1785
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
         Left            =   6495
         TabIndex        =   70
         Top             =   90
         Visible         =   0   'False
         Width           =   3585
      End
   End
   Begin VB.Frame frBotao 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5265
      Left            =   10260
      TabIndex        =   57
      Top             =   720
      Width           =   1125
      Begin prjChameleon.chameleonButton cmdSair 
         Height          =   315
         Left            =   30
         TabIndex        =   58
         ToolTipText     =   "Sair da Tela"
         Top             =   4815
         Width           =   1050
         _ExtentX        =   1852
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
         MICON           =   "frmDebitoImob.frx":3369
         PICN            =   "frmDebitoImob.frx":3385
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdReparc 
         Height          =   315
         Left            =   45
         TabIndex        =   59
         ToolTipText     =   "Tela de Reparcelamento"
         Top             =   1850
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Reparcelam."
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
         MICON           =   "frmDebitoImob.frx":33F3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDoc 
         Height          =   315
         Left            =   45
         TabIndex        =   60
         ToolTipText     =   "Documento(s) da Parcela"
         Top             =   1498
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Documento"
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
         MICON           =   "frmDebitoImob.frx":340F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExtrato 
         Height          =   315
         Left            =   45
         TabIndex        =   61
         ToolTipText     =   "Emissão de Extrato"
         Top             =   1146
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Extrato"
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
         MICON           =   "frmDebitoImob.frx":342B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDetalhe 
         Height          =   315
         Left            =   45
         TabIndex        =   62
         ToolTipText     =   "Detalhes da Parcela"
         Top             =   794
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Detalhe"
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
         MICON           =   "frmDebitoImob.frx":3447
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdObs 
         Height          =   315
         Left            =   45
         TabIndex        =   63
         ToolTipText     =   "Observação"
         Top             =   442
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Observação"
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
         MICON           =   "frmDebitoImob.frx":3463
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdFilter 
         Height          =   315
         Left            =   45
         TabIndex        =   64
         ToolTipText     =   "Ativa e Desativa o Filtro"
         Top             =   90
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Filtro"
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
         MICON           =   "frmDebitoImob.frx":347F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDAM 
         Height          =   315
         Left            =   45
         TabIndex        =   65
         ToolTipText     =   "Emissão de DAM"
         Top             =   2202
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&D.A.M."
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
         MICON           =   "frmDebitoImob.frx":349B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAj 
         Height          =   315
         Left            =   45
         TabIndex        =   66
         ToolTipText     =   "Outras opções"
         Top             =   2910
         Width           =   1050
         _ExtentX        =   1852
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":34B7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdEF 
         Height          =   315
         Left            =   45
         TabIndex        =   102
         ToolTipText     =   "Execuções Fiscais"
         Top             =   2554
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Exec.Fiscal"
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
         MICON           =   "frmDebitoImob.frx":34D3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAjuiza 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Label4"
         Height          =   375
         Left            =   210
         TabIndex        =   76
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
   End
   Begin VB.Frame frDoc 
      BackColor       =   &H00000000&
      Caption         =   "Consulta Documentos"
      ForeColor       =   &H00C0FFFF&
      Height          =   2925
      Left            =   3390
      TabIndex        =   53
      Top             =   1710
      Width           =   3765
      Begin MSFlexGridLib.MSFlexGrid grdDoc 
         Height          =   1980
         Left            =   90
         TabIndex        =   54
         Top             =   330
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   3493
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "^Nº Documento  |^Data Doc.   |>Valor Pago     "
      End
      Begin prjChameleon.chameleonButton cmdSairDoc 
         Height          =   375
         Left            =   3285
         TabIndex        =   55
         ToolTipText     =   "Sair da Tela de Documento"
         Top             =   2400
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   661
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
         MICON           =   "frmDebitoImob.frx":34EF
         PICN            =   "frmDebitoImob.frx":350B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCnsDoc 
         Height          =   375
         Left            =   2835
         TabIndex        =   56
         ToolTipText     =   "Consulta Documento"
         Top             =   2400
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
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
         MICON           =   "frmDebitoImob.frx":3665
         PICN            =   "frmDebitoImob.frx":3681
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
   Begin VB.Frame frStatus 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   345
      Left            =   30
      TabIndex        =   46
      Top             =   5640
      Width           =   10185
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Débito (Parcelado):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   52
         Top             =   60
         Width           =   1785
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vencido (Parcelado):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   3540
         TabIndex        =   51
         Top             =   60
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecionado..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   7230
         TabIndex        =   50
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label lblDebito 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   2130
         TabIndex        =   49
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label lblVencer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   5490
         TabIndex        =   48
         Top             =   60
         Width           =   1275
      End
      Begin VB.Label lblSel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Left            =   8610
         TabIndex        =   47
         Top             =   60
         Width           =   1275
      End
   End
   Begin VB.Frame frEFiscal 
      BackColor       =   &H00ADE7FF&
      Height          =   5355
      Left            =   900
      TabIndex        =   99
      Top             =   315
      Visible         =   0   'False
      Width           =   8295
      Begin VB.Frame frEfObs 
         BackColor       =   &H00ADE7FF&
         Height          =   1995
         Left            =   90
         TabIndex        =   119
         Top             =   2880
         Width           =   1320
         Begin VB.TextBox txtDocEF 
            Appearance      =   0  'Flat
            Height          =   1560
            Left            =   90
            MaxLength       =   2000
            MultiLine       =   -1  'True
            TabIndex        =   121
            Top             =   360
            Width           =   1140
         End
         Begin prjChameleon.chameleonButton cmdEfObs 
            Height          =   240
            Left            =   45
            TabIndex        =   120
            Top             =   0
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   "Observação"
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmDebitoImob.frx":385B
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
      Begin VB.Frame frEFDoc 
         BackColor       =   &H00ADE7FF&
         Height          =   1995
         Left            =   1440
         TabIndex        =   116
         Top             =   2880
         Width           =   6765
         Begin MSComctlLib.ListView lvDoc 
            Height          =   1560
            Left            =   45
            TabIndex        =   117
            Top             =   360
            Width           =   6645
            _ExtentX        =   11721
            _ExtentY        =   2752
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Descricao"
               Object.Width           =   7586
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Qtde"
               Object.Width           =   970
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Situação"
               Object.Width           =   2540
            EndProperty
         End
         Begin prjChameleon.chameleonButton cmdEFDoc 
            Height          =   240
            Left            =   45
            TabIndex        =   118
            Top             =   0
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   "Documentos"
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   128
            FCOLO           =   128
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmDebitoImob.frx":3877
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdEfQtde 
            Height          =   285
            Left            =   5175
            TabIndex        =   122
            ToolTipText     =   "Alterar quantidade do item"
            Top             =   0
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   503
            BTYPE           =   14
            TX              =   "Qtde"
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   128
            FCOLO           =   128
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmDebitoImob.frx":3893
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdEFSit 
            Height          =   285
            Left            =   5760
            TabIndex        =   123
            ToolTipText     =   "Alterar situação do item"
            Top             =   0
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            BTYPE           =   14
            TX              =   "Situação"
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   128
            FCOLO           =   128
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmDebitoImob.frx":38AF
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
      Begin VB.TextBox txtEF 
         Height          =   330
         Left            =   4410
         MaxLength       =   25
         TabIndex        =   113
         Top             =   315
         Visible         =   0   'False
         Width           =   3390
      End
      Begin prjChameleon.chameleonButton cmdRetornar 
         Height          =   315
         Left            =   6960
         TabIndex        =   103
         ToolTipText     =   "Retornar a tela de débito"
         Top             =   4950
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
         MICON           =   "frmDebitoImob.frx":38CB
         PICN            =   "frmDebitoImob.frx":38E7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAlterarEF 
         Height          =   315
         Left            =   4785
         TabIndex        =   105
         ToolTipText     =   "Alterar execução fiscal"
         Top             =   4950
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":3955
         PICN            =   "frmDebitoImob.frx":3971
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNovoEF 
         Height          =   315
         Left            =   3690
         TabIndex        =   106
         ToolTipText     =   "Nova execução fiscal"
         Top             =   4950
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":3ACB
         PICN            =   "frmDebitoImob.frx":3AE7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExcluirEF 
         Height          =   315
         Left            =   5880
         TabIndex        =   107
         ToolTipText     =   "Excluir execução fiscal"
         Top             =   4950
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "E&xcluir"
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
         MICON           =   "frmDebitoImob.frx":3C41
         PICN            =   "frmDebitoImob.frx":3C5D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGravarEF 
         Height          =   315
         Left            =   5880
         TabIndex        =   108
         ToolTipText     =   "Gravar os Dados"
         Top             =   4950
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":3CFF
         PICN            =   "frmDebitoImob.frx":3D1B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvEFOrigem 
         Height          =   2100
         Left            =   90
         TabIndex        =   109
         Top             =   720
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   3704
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ano"
            Object.Width           =   1271
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Lc."
            Object.Width           =   795
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Sq."
            Object.Width           =   795
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Pc"
            Object.Width           =   707
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Cp"
            Object.Width           =   707
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Vencto"
            Object.Width           =   1854
         EndProperty
      End
      Begin MSComctlLib.ListView lvEFDest 
         Height          =   2100
         Left            =   4410
         TabIndex        =   110
         Top             =   720
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   3704
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ano"
            Object.Width           =   1024
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Lc."
            Object.Width           =   795
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Sq."
            Object.Width           =   795
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Pc"
            Object.Width           =   707
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Cp"
            Object.Width           =   707
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Vencto"
            Object.Width           =   1854
         EndProperty
      End
      Begin prjChameleon.chameleonButton cmdCancelarEF 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   6960
         TabIndex        =   104
         ToolTipText     =   "Cancelar Edição"
         Top             =   4950
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":40C0
         PICN            =   "frmDebitoImob.frx":40DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdC2 
         Height          =   285
         Left            =   3960
         TabIndex        =   111
         ToolTipText     =   "Remove centro de custos"
         Top             =   1455
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
         MICON           =   "frmDebitoImob.frx":4236
         PICN            =   "frmDebitoImob.frx":4252
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdC1 
         Height          =   285
         Left            =   3960
         TabIndex        =   112
         ToolTipText     =   "Adiciona centro de custos"
         Top             =   1125
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
         MICON           =   "frmDebitoImob.frx":43AC
         PICN            =   "frmDebitoImob.frx":43C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAllEfo 
         Height          =   225
         Left            =   3375
         TabIndex        =   114
         ToolTipText     =   "Seleciona todos"
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   14
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   32768
         FCOLO           =   32768
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":4522
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelEfo 
         Height          =   225
         Left            =   3645
         TabIndex        =   115
         ToolTipText     =   "Remove todos"
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   14
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDebitoImob.frx":453E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cmbEF 
         Height          =   315
         Left            =   4410
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   315
         Width           =   3630
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Débitos disponíveis para seleção"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   3675
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdExtrato 
      Height          =   4875
      Left            =   30
      TabIndex        =   4
      Top             =   720
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   8599
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   16777215
      HighlightForeColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
   End
   Begin VB.Label lblDataVencto 
      Caption         =   "Label5"
      Height          =   210
      Left            =   1950
      TabIndex        =   3
      Top             =   7050
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "frmDebitoImob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuContrib As cPopupMenu
Attribute m_cMenuContrib.VB_VarHelpID = -1
Public WithEvents m_cMenuOpcoes As cPopupMenu
Attribute m_cMenuOpcoes.VB_VarHelpID = -1
Public WithEvents m_cMenuExtrato As cPopupMenu
Attribute m_cMenuExtrato.VB_VarHelpID = -1
Public WithEvents m_cMenuInterno As cPopupMenu
Attribute m_cMenuInterno.VB_VarHelpID = -1

Private Type Debito
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNotificado As String
    sExFiscal As String
    nProt_certidao As Long
    nProt_dtremessa As Date
End Type

Private Type Multa
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nStatus As Integer
    bAchou As Boolean
End Type

Dim RdoAux As rdoResultset, Sql As String
Dim bSel As Boolean, nCodObs As Integer, bChangeStatus As Boolean, bExtrato As Boolean
Dim sRet As String, bCarregado As Boolean
Dim evDAM As Integer, evCND As Integer, evDAT As Integer, evAJU As Integer, evADO As Integer, evEDI As Integer, evSMA As Integer, evSMOV As Integer, evCOM As Integer, evRea As Integer, evReaJ As Integer, evRP As Integer, evEF As Integer, evSer As Integer
Dim bDam As Boolean, bCND As Boolean, bDAT As Boolean, bAJU As Boolean, bADO As Boolean, bEDI As Boolean, bSMA As Boolean, bSMOV As Boolean, bCOM As Boolean, bRea As Boolean, bReaJ As Boolean, bRP As Boolean, bEF As Boolean, bSer As Boolean
Dim xImovel As clsImovel, bNovoObs As Boolean, bObs As Boolean, bExecF As Boolean
Dim bFilterLoad As Boolean, nExtrato As Integer, sEventoEF As String

Private Sub chkUnica_Click()
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End Sub

Private Sub cmbAj_Click()
If cmbAj.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If
End Sub

Private Sub cmbAno1_Click()
If cmbAno1.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If
End Sub

Private Sub cmbAno2_Click()
If cmbAno2.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If
End Sub

Private Sub cmbDA_Click()
If cmbDA.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If
End Sub

Private Sub cmbEF_Click()
Dim x As Integer, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sEF As String, itmX As ListItem, sVencto As String
Dim Sql As String, RdoAux As rdoResultset, sNum As String, nNum As Integer

z = SendMessage(lvEFDest.hwnd, LVM_DELETEALLITEMS, 0, 0)
With grdExtrato
    
    For x = 1 To .Rows
        nAno = Val(.CellText(x, 1))
        nLanc = Val(Left$(.CellText(x, 2), 3))
        nSeq = Val(.CellText(x, 3))
        nParc = IIf(.CellText(x, 4) = "Unica", 0, Val(.CellText(x, 4)))
        nCompl = Val(.CellText(x, 5))
        sVencto = .CellText(x, 7)
        sEF = .CellText(x, 14)
        If sEF = cmbEF.Text Then
            Set itmX = lvEFDest.ListItems.Add(, "C" & Format(x, "0000"), nAno)
            itmX.SubItems(1) = Format(nLanc, "000")
            itmX.SubItems(2) = Format(nSeq, "000")
            itmX.SubItems(3) = Format(nParc, "00")
            itmX.SubItems(4) = Format(nCompl, "00")
            itmX.SubItems(5) = sVencto
        End If
    Next
    
End With

sNum = cmbEF.Text
'nNum = Val(Left$(sNum, InStr(1, sNum, "/", vbBinaryCompare) - 1))
'nAno = Val(Right$(sNum, 4))

Sql = "SELECT * FROM EXECUCAOFISCAL WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND processocnj='" & sNum & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtDocEF.Text = SubNull(!obs)
       .Close
    End If
End With

CarregaOrigemEF
End Sub

Private Sub cmbLanc_Click()
If cmbLanc.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If
End Sub

Private Sub cmbProc_Click()
Dim RdoAux2 As rdoResultset
Dim sprotocolo As String
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nComp As Integer
Dim sDataVencto As String, bTudoPago As Boolean

grdOrigem.Rows = 1
grdDestino.Rows = 1
If Right$(cmbProc.Text, 4) <> "SMAR" Then

    Sql = "SELECT * FROM vwCNSREPARCELAMENTOO WHERE  CODREDUZIDO=" & Val(txtCod.Text) & " AND  NUMPROCESSO='" & cmbProc.Text & "' "
    Sql = Sql & "ORDER BY CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMPARCELA,DATAVENCIMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
             lblDataProc.Caption = Format(!DATAPROCESSO, "dd/mm/yyyy")
             lblDataParc.Caption = Format(!datareparc, "dd/mm/yyyy")
             lblQtde.Caption = !qtdeparcela
             lblValor.Caption = FormatNumber(!VALORENTRADA, 2)
             lblPerc.Caption = FormatNumber(!PERCENTRADA, 2)
             chkMulta.value = IIf(!CalculaMulta, 1, 0)
             chkJuros.value = IIf(!CalculaJuros, 1, 0)
             lblResp.Caption = Format(!CODIGORESP, "0000000")
             'lblFunc.Caption = IIf(IsNull(!funcionario), "SMAR", !funcionario)
             lblFunc.Caption = IIf(IsNull(!NomeLogin), "SMAR", !NomeLogin)
             grdOrigem.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!numsequencia, "00") & Chr(9) & _
             Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy")
            .MoveNext
        Loop
    End With

    Sql = "SELECT numprocesso,numproc,anoproc,plano,nome FROM  processoreparc INNER JOIN plano ON processoreparc.plano = plano.codigo "
    Sql = Sql & "WHERE numprocesso='" & cmbProc.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            lblPlano.Caption = RdoAux!Nome
        Else
            lblPlano.Caption = "Sem plano"
        End If
       .Close
    End With
    
    Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & cmbProc.Text & "' AND CODREDUZIDO = " & Val(txtCod.Text)
    Sql = Sql & " ORDER BY ANOEXERCICIO,CODLANCAMENTO,NUMPARCELA,DATAVENCIMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        
        Do Until .EOF
             If .AbsolutePosition = 1 Then
                 lbl1venc.Caption = Format(!DataVencimento, "dd/mm/yyyy")
             End If
             grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!numsequencia, "00") & Chr(9) & _
             Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy")
            .MoveNext
        Loop
    End With
    If Right$(cmbProc.Text, 4) = "SMAR" Then
        Sql = "SELECT NUMprotocolo FROM PROCESSOREPARC WHERE NUMPROCESSO='" & cmbProc.Text & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sprotocolo = Trim(SubNull(!NUMprotocolo))
           .Close
        End With
        lblFunc.Caption = lblFunc.Caption & " PROT: " & sprotocolo
    End If
    Sql = "SELECT CANCELADO FROM PROCESSOREPARC WHERE NUMPROCESSO='" & cmbProc.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         If .RowCount = 0 Then
            lblCancel.Visible = False
         Else
            If !Cancelado Then
                lblCancel.Visible = True
            Else
                lblCancel.Visible = False
            End If
         End If
        .Close
    End With
Else
1:
    Sql = "SELECT CODREDUZD FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.Text) & " AND CODSEQD=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
       nCodReduz = RdoAux!CODREDUZD
    Else
        Exit Sub
    End If
    RdoAux.Close
       
    Sql = "SELECT * FROM REPARC2TMP WHERE CODREDUZ=" & nCodReduz & " AND CODSEQ=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        On Error Resume Next
        lblDataProc.Caption = ""
        lblDataParc.Caption = Format(!DataVencto, "dd/mm/yyyy")
        lblQtde.Caption = !PARCELAS
        lblValor.Caption = FormatNumber(!ValorPago, 2)
        lblPerc.Caption = "0,00"
        chkMulta.value = IIf(!TEMMULTA, 1, 0)
        chkJuros.value = IIf(!TEMJUROS, 1, 0)
        lblResp.Caption = Format(!CodReduz, "0000000")
        lblFunc.Caption = "SMAR PROT: " & !NUMprotocolo
       .Close
    End With

    Sql = "SELECT DISTINCT * FROM REPARCTMP WHERE CODREDUZD=" & nCodReduz & " AND CODSEQD=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
       If .RowCount > 0 Then
           If !CODSIT > 0 Then
               lblCancel.Visible = True
           Else
               lblCancel.Visible = False
           End If
       Else
           lblCancel.Visible = False
       End If
       Do Until .EOF
           Sql = "SELECT CODREDUZIDO,DATAVENCIMENTO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZO & " AND ANOEXERCICIO=" & !ANOEXERCO & " AND CODLANCAMENTO=" & !CODLANCO & " AND SEQLANCAMENTO=" & !CODSEQO & " AND NUMPARCELA=" & !NUMPARCO & " AND CODCOMPLEMENTO=" & !CODCOMPLO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux2
               If .RowCount > 0 Then
                    sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
                End If
               .Close
            End With
            grdOrigem.AddItem Format(!CODREDUZO, "0000000") & Chr(9) & !ANOEXERCO & Chr(9) & Format(!CODLANCO, "00") & Chr(9) & Format(!CODSEQO, "00") & Chr(9) & _
            Format(!NUMPARCO, "00") & Chr(9) & Format(!CODCOMPLO, "00") & Chr(9) & sDataVencto
           .MoveNext
        Loop
       .Close
    End With
    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND STATUSLANC<>5 AND SEQLANCAMENTO=" & Left$(cmbProc.Text, Len(cmbProc) - 5) & "  ORDER BY DATAVENCIMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
          grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & _
           Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy")
          .MoveNext
        Loop
       .Close
    End With
End If

bTudoPago = True
With grdDestino
    For x = 1 To .Rows - 1
        nCodReduz = .TextMatrix(x, 0)
        nAno = .TextMatrix(x, 1)
        nLanc = .TextMatrix(x, 2)
        nSeq = .TextMatrix(x, 3)
        nParc = .TextMatrix(x, 4)
        nComp = .TextMatrix(x, 5)
        
        Sql = "SELECT debitopago.codreduzido, debitopago.datapagamento, debitoparcela.statuslanc "
        Sql = Sql & "FROM  debitopago RIGHT OUTER JOIN debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND "
        Sql = Sql & "debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
        Sql = Sql & "debitopago.numparcela = debitoparcela.numparcela AND debitopago.codcomplemento = debitoparcela.codcomplemento WHERE debitoparcela.CODREDUZIDO=" & nCodReduz & " AND debitoparcela.ANOEXERCICIO=" & nAno
        Sql = Sql & " AND debitoparcela.CODLANCAMENTO=" & nLanc & " AND debitoparcela.SEQLANCAMENTO=" & nSeq & " AND debitoparcela.NUMPARCELA=" & nParc
        Sql = Sql & " AND debitoparcela.CODCOMPLEMENTO=" & nComp
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If !statuslanc = 6 Or !statuslanc = 28 Then
                        grdDestino.TextMatrix(x, 7) = "Compensado"
                Else
                    If Not IsNull(!DataPagamento) Then
                        grdDestino.TextMatrix(x, 7) = Format(!DataPagamento, "dd/mm/yyyy")
                    Else
                        If !statuslanc = 2 Or !statuslanc = 1 Or !statuslanc = 7 Then
                            grdDestino.TextMatrix(x, 7) = "Pago sem Data"
                        Else
                            bTudoPago = False
                        End If
                    End If
                End If
            Else
                Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
                Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
                Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO = " & nAno
                Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND NUMPARCELA=" & nParc & " AND SEQLANCAMENTO=" & nSeq
                Sql = Sql & " AND CODCOMPLEMENTO=" & nComp & " AND VALORPAGO>0"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount = 0 Then
                        bTudoPago = False
                    Else
                        grdDestino.TextMatrix(x, 7) = "Pago sem Data"
                    End If
                   .Close
                End With
            End If
           .Close
        End With
    Next
End With

If bTudoPago And lblCancel.Visible = False Then
    If grdDestino.Rows = 1 Then
       lblPago.Visible = False
    Else
       lblPago.Visible = True
    End If
Else
    lblPago.Visible = False
End If

End Sub

Private Sub cmbSeq_Click()
If cmbSeq.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If

End Sub

Private Sub cmbSit_Click()
If cmbSit.ListIndex > -1 And bExecF Then
    bCarregado = False
    CarregaDebito (Val(txtCod.Text))
End If
End Sub

Private Sub CMDAJ_Click()

lIndex = m_cMenuOpcoes.ShowPopupMenu(frBotao.Left + cmdAj.Left, cmdAj.Top, cmdAj.Left, cmdAj.Top, Me.ScaleWidth - cmdAj.Left - cmdAj.Width, cmdAj.Top + cmdAj.Height, False)

End Sub

Private Sub cmdAllEfo_Click()
Dim x As Integer
With lvEFOrigem
    For x = 1 To .ListItems.Count
        .ListItems(x).Checked = True
    Next
End With

End Sub

Private Sub cmdAlterarEF_Click()
If cmbEF.ListIndex = -1 Then
    MsgBox "Selecione uma execução fiscal.", vbCritical, "Atenção"
Else
    sEventoEF = "Alterar"
    EventosEF False
End If
End Sub

Private Sub cmdAlterarObs_Click()
Dim sNomeUser As String
If txtObservacao.Text = "" Then
    MsgBox "Selecione o registro que deseja alterar.", vbExclamation, "Atenção"
Else
    sNomeUser = UCase(lvObserv.SelectedItem.SubItems(1))
    If NomeDeLogin = "ROSE" Or NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "JOSEANE" Or NomeDeLogin = "RHENO.SOARES" Or NomeDeLogin = "DANIELE.SILVA" Or NomeDeLogin = "MATHEUS.BOTELHO" Or NomeDeLogin = "GISELE.ALMEIDA" Or NomeDeLogin = "WHICTOR.HOMEM" Then
            bNovoObs = False
            EventosObs False
    Else
        If sNomeUser <> UCase(NomeDeLogin) And Right(sNomeUser, Len(NomeDeLogin)) <> UCase(NomeDeLogin) Then
            MsgBox "Você não pode alterar uma observação criada por outro usuário.", vbCritical, "ALERTA DE SEGURANÇA"
        Else
            bNovoObs = False
            EventosObs False
        End If
    End If
End If

End Sub

Private Sub cmdC1_Click()
Dim itmX As ListItem, z As Long, sVencto As String, x As Integer, idx As Integer

With lvEFOrigem
Inicio:
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked Then
           
            Set itmX = lvEFDest.ListItems.Add(, .ListItems(x).Text & .ListItems(x).SubItems(1) & .ListItems(x).SubItems(2) & .ListItems(x).SubItems(3) & .ListItems(x).SubItems(4), .ListItems(x).Text)
            itmX.SubItems(1) = .ListItems(x).SubItems(1)
            itmX.SubItems(2) = .ListItems(x).SubItems(2)
            itmX.SubItems(3) = .ListItems(x).SubItems(3)
            itmX.SubItems(4) = .ListItems(x).SubItems(4)
            itmX.SubItems(5) = .ListItems(x).SubItems(5)
           .ListItems.Remove (x)
            GoTo Inicio
        End If
    Next
End With

End Sub

Private Sub cmdC2_Click()
Dim itmX As ListItem, z As Long, sVencto As String, x As Integer

With lvEFDest
Inicio:
    For x = 1 To .ListItems.Count
        If .ListItems(x).Selected Then
            Set itmX = lvEFOrigem.ListItems.Add(, .ListItems(x).Text & .ListItems(x).SubItems(1) & .ListItems(x).SubItems(2) & .ListItems(x).SubItems(3) & .ListItems(x).SubItems(4), .ListItems(x).Text)
            itmX.SubItems(1) = .ListItems(x).SubItems(1)
            itmX.SubItems(2) = .ListItems(x).SubItems(2)
            itmX.SubItems(3) = .ListItems(x).SubItems(3)
            itmX.SubItems(4) = .ListItems(x).SubItems(4)
            itmX.SubItems(5) = .ListItems(x).SubItems(5)
           .ListItems.Remove (x)
            GoTo Inicio
        End If
    Next
End With

End Sub

Private Sub cmdCalculo_Click()
Dim nSomaLiquido As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double, nSomaPrincipal As Double, nValorPrincipal As Double
Dim nNumParcela As Integer, dVencimento As Date, nJuros As Double, nValorParcela As Double, nSaldo As Double, nJurosMesPerc As Double, RdoAux2 As rdoResultset
Dim nJurosMesValor As Double, nValorHonorario As Double, nValorTotal As Double, x As Integer, aCalculo() As Debito, bJurosMulta As Boolean, sExercicio As String

Sql = "SELECT MIN(ANOEXERCICIO) AS MINIMO, MAX(ANOEXERCICIO) AS MAXIMO "
Sql = Sql & "From ORIGEMREPARC WHERE NUMPROCESSO = '" & cmbProc.Text & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    lblAno.Caption = !MINIMO & " - " & !maximo
   .Close
End With

ReDim aCalculo(0)
Sql = "SELECT * FROM DESTINOREPARC WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND NUMPROCESSO='" & cmbProc.Text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Cálculo não disponível para este parcelamento.", vbExclamation, "Atenção"
        Exit Sub
    Else
        If Val(SubNull(!VALORLIQUIDO)) = 0 Then
            MsgBox "Cálculo não disponível para este parcelamento.", vbExclamation, "Atenção"
            Exit Sub
        Else
            bJurosMulta = False
            Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento
            Sql = Sql & " AND SEQLANCAMENTO=" & !numsequencia & " AND NUMPARCELA=1  And CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO=113"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                bJurosMulta = True
            End If
            RdoAux2.Close
            
            Do Until .EOF
                ReDim Preserve aCalculo(UBound(aCalculo) + 1)
                aCalculo(UBound(aCalculo)).nParc = !NumParcela
                aCalculo(UBound(aCalculo)).nValorTributo = !ValorPrincipal
                aCalculo(UBound(aCalculo)).nValorJurApl = !jurosapl
                aCalculo(UBound(aCalculo)).nValorAtual = !Total
                aCalculo(UBound(aCalculo)).nSaldo = !Saldo
                aCalculo(UBound(aCalculo)).nValorMulta = !jurosperc  'REAPROVEITAMOS A VARIAVEL nVALORMULTA, MAIS SE APLICA AO CAMPO JUROSPERC
                aCalculo(UBound(aCalculo)).nValorJuros = !jurosvalor
                aCalculo(UBound(aCalculo)).nValorHon = !honorario
               .MoveNext
            Loop
        End If
    End If
   .Close
End With


Sql = "DELETE FROM CALCULOPARCELAMENTO WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

nSomaLiquido = 0
nSomaJuros = 0
nSomaMulta = 0
nSomaCorrecao = 0
nSomaPrincipal = 0

Sql = "SELECT * FROM ORIGEMREPARC WHERE NUMPROCESSO='" & cmbProc.Text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nSomaLiquido = nSomaLiquido + !principal
        If bJurosMulta Then
            nSomaJuros = nSomaJuros + !Juros
            nSomaMulta = nSomaMulta + !Multa
        End If
        nSomaCorrecao = nSomaCorrecao + !CORRECAO
       .MoveNext
    Loop
   .Close
End With
nValorTotal = nSomaLiquido + nSomaJuros + nSomaMulta + nSomaCorrecao

For x = 1 To UBound(aCalculo)
    With aCalculo(x)
        Sql = "INSERT CALCULOPARCELAMENTO(COMPUTER,CODREDUZIDO,NOME,PROCESSO,DATAPROCESSO,QTDEPARCELA,SOMALIQUIDO,SOMAJUROS,SOMAMULTA,SOMACORRECAO,SOMAPRINCIPAL,"
        Sql = Sql & "VALORPRINCIPAL,NUMPARCELA,VENCIMENTO,JUROS,VALORPARCELA,SALDO,JUROSMESPERC,JUROSMESVALOR,VALORHONORARIO,VALORTOTAL) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & Val(txtCod.Text) & ",'" & Left(lblProp.Caption, 50) & "','" & cmbProc.Text & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "',"
        Sql = Sql & Val(lblQtde.Caption) & "," & Virg2Ponto(RemovePonto(CStr(nSomaLiquido))) & "," & Virg2Ponto(RemovePonto(CStr(nSomaJuros))) & "," & Virg2Ponto(RemovePonto(CStr(nSomaMulta))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(nSomaCorrecao))) & "," & Virg2Ponto(RemovePonto(CStr(nValorTotal))) & "," & Virg2Ponto(RemovePonto(CStr(.nValorTributo))) & "," & .nParc & ",'"
        Sql = Sql & Format(grdDestino.TextMatrix(x, 6), "mm/dd/yyyy") & "'," & Virg2Ponto(RemovePonto(CStr(.nValorJurApl))) & "," & Virg2Ponto(RemovePonto(CStr(.nValorTributo))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.nSaldo))) & "," & Virg2Ponto(RemovePonto(CStr(.nValorMulta))) & "," & Virg2Ponto(RemovePonto(CStr(.nValorJuros))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.nValorHon))) & "," & Virg2Ponto(RemovePonto(CStr(.nValorAtual))) & ")"
        cn.Execute Sql, rdExecDirect
    End With
Next

'Exit Sub
If frmMdi.frTeste.Visible = True Then
    If frmMdi.frTeste.Caption = "ACESSANDO OS DADOS LOCAIS" Then
        frmReport.ShowReport "CALCULOPARCELAMENTO", frmMdi.hwnd, Me.hwnd
    Else
        frmReport.ShowReport "CALCULOPARCELAMENTOTMP", frmMdi.hwnd, Me.hwnd
    End If
Else
    frmReport.ShowReport "CALCULOPARCELAMENTO", frmMdi.hwnd, Me.hwnd
End If

Sql = "DELETE FROM CALCULOPARCELAMENTO WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdCancelarEF_Click()
CarregaOrigemEF
sEventoEF = ""
EventosEF True
End Sub

Private Sub cmdCancelarObs_Click()
EventosObs True
End Sub

Private Sub cmdCancelReparc_Click()
Dim bAchou As Boolean, nSeq As Integer

nSeq = grdDestino.TextMatrix(1, 3)

If lblCancel.Visible = True Then
    MsgBox "Este reparcelamento já foi cancelado.", vbExclamation, "Atenção"
Else
    If Right$(cmbProc.Text, 4) <> "SMAR" Then
        MsgBox "Apenas reparcelamentos feitos na SMAR podem ser cancelados por aqui.", vbExclamation, "Atenção"
    Else
        If MsgBox("Deseja CANCELAR este reparcelamento ? " & vbCrLf & "OBS: Caso houverem lancamentos ativos na tela de consulta estes não serão cancelados.", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            Sql = "UPDATE REPARCTMP SET CODSIT=1 WHERE CODREDUZD=" & Val(txtCod.Text) & " AND CODSEQD=" & nSeq
            cn.Execute Sql, rdExecDirect
            lblCancel.Visible = True
        End If
    End If
End If

End Sub

Private Sub cmdCnsDoc_Click()
frmDoc.show: frmDoc.ZOrder 0
frmDoc.txtNumDoc.Text = Val(grdDoc.TextMatrix(grdDoc.Row, 0))
frmDoc.txtNumDoc_KeyPress (vbKeyReturn)
End Sub

Private Sub cmdCnsImovel_Click()

lIndex = m_cMenuContrib.ShowPopupMenu(cmdCnsImovel.Left, cmdCnsImovel.Top, cmdCnsImovel.Left, cmdCnsImovel.Top, Me.ScaleWidth - cmdCnsImovel.Left - cmdCnsImovel.Width, cmdCnsImovel.Top + cmdCnsImovel.Height, False)

End Sub

Private Sub cmdDAM_Click()
Dim Achou As Boolean, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim vData As Variant, dData As Date, sStat As String, Achou2 As Boolean

'If bLocal Then
'    Exit Sub
'End If

ReDim aDocDAM(0)
Achou = False
With grdExtrato
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
            nAno = Val(.CellText(x, 1))
            nLanc = Val(Left$(.CellText(x, 2), 3))
            nSeq = Val(.CellText(x, 3))
            nParc = IIf(.CellText(x, 4) = "Unica", 0, Val(.CellText(x, 4)))
            nCompl = Val(.CellText(x, 5))
            sStat = .CellText(x, 8)
            If Val(Left(.CellText(x, 6), 2)) <> 3 And Val(Left(.CellText(x, 6), 2)) <> 38 And Val(Left(.CellText(x, 6), 2)) <> 39 Then
                MsgBox "Só é possível emitir DAM para lançamentos não pagos e protestados.", vbExclamation, "Atenção"
                Exit Sub
            End If
                                        
                                
            If bAnistia Then
                Sql = "SELECT * FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl & " AND NUMDOCUMENTO > 2000000 AND NUMDOCUMENTO< 2900000"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                       If nAno > Year(Now) - 1 Then
                            MsgBox "Não é possível emitir DAM para lançamento gerado pelo ISS Eletrônico de " & Year(Now) & "." & vbCrLf & "Por favor encaminhar o contribuinte ao Setor de Tributos de Movimentação Econômica, para emissão de segunda via.", vbExclamation, "Atenção"
                            Achou = True
                            Exit Sub
                        End If
                        ReDim Preserve aDocDAM(UBound(aDocDAM) + 1)
                        aDocDAM(UBound(aDocDAM)) = !NumDocumento
                    End If
                   .Close
                End With
            Else
'                If sStat = "N" Then
                If Val(Left(.CellText(x, 2), 3)) = 5 Then
                
'                    Sql = "SELECT * FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=5 AND "
'                    Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl & " AND NUMDOCUMENTO between 2000000 AND 2900000 "
'                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                    With RdoAux
 '                       If .RowCount > 0 Then
 '                           If !AnoExercicio > 2015 Then
 '                               MsgBox "Não é possível emitir DAM para lançamento gerado pelo sistema Giss Online." & vbCrLf & "O contribuinte deverá gerar a guia pelo sistema de ISS.", vbExclamation, "Atenção"
 '                               Achou = True
 '                               Exit Sub
 '                           End If
'                        End If
 '                       .Close
'                    End With
                End If
           End If
    
        End If
    Next
End With


Achou = False
With grdExtrato
    For x = 1 To .Rows
        If CDbl(.CellText(x, 10)) = 0 And .CellText(x, 12) = "S" Then
            Achou = True
            Exit For
        End If
    Next
End With

If Achou Then
    MsgBox "Não é possível emitir DAM para débitos com valor zerado.", vbCritical, "Atenção"
    Exit Sub
End If



If bAnistia Then
    Achou = False
    With grdExtrato
        For x = 1 To .Rows
            If (Left(.CellText(x, 2), 3) = 78 Or Left(.CellText(x, 2), 3) = 41) And .CellText(x, 12) = "S" Then
                Achou = True
                Exit For
            End If
        Next
    End With
    
    Achou2 = False
    With grdExtrato
        For x = 1 To .Rows
            If Left(.CellText(x, 2), 3) <> 78 And Left(.CellText(x, 2), 3) <> 41 And .CellText(x, 12) = "S" And Left(.CellText(x, 2), 3) <> 78 And Left(.CellText(x, 2), 3) <> 41 And .CellText(x, 1) >= 2015 Then
                Achou2 = True
                Exit For
            End If
        Next
    End With
        
  '  If Achou And Achou2 Then
 '       MsgBox "Não é possível emitir DAM para AR Digital e/ou Despesas Judiciais junto com outros débitos durante o REFIS.", vbCritical, "Atenção"
'        Exit Sub
'    End If
    
   
End If





Achou = False
With grdExtrato
    For x = 1 To .Rows
        If Left(.CellText(x, 2), 3) = 5 And .CellText(x, 12) = "S" Then
            Achou = True
            Exit For
        End If
    Next
End With

'If Achou Then
'    MsgBox "Não é possível emitir DAM para ISS Variável." & vbCrLf & "Utilize o sistema de ISS Eletrônico.", vbCritical, "Atenção"
'    Exit Sub
'End If

Achou = False
With grdExtrato
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
            Achou = True
            Exit For
        End If
    Next
End With

If Not Achou Then
    MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
Else
    If Not bAnistia Then
        If Not ValidaMI Then
            If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "RENATA" And NomeDeLogin <> "JOSEANE" Then
                Exit Sub
            End If
        End If
    End If
Inicio:
    vData = InputBox("Digite a Data de Vencimento da DAM.", "Atenção", Format(Now, "dd/mm/yyyy"))
    If vData <> "" Then
       If Len(vData) <> 10 Then
          MsgBox "Data inválida.", vbCritical, "Atenção"
          GoTo Inicio
       Else
          If Not IsDate(vData) Then
             MsgBox "Data inválida.", vbCritical, "Atenção"
             GoTo Inicio
          Else
             dData = CDate(vData)
             If dData < Format(Now, "dd/mm/yyyy") Then
                MsgBox "Data de vencimento não pode ser retroativa.", vbCritical, "Atenção"
                GoTo Inicio
             Else
                lblDataVencto.Caption = vData
                frmDAM.Honorarios = False
                frmDAM.CodigoDAM = CLng(txtCod.Text)
                frmDAM.VencimentoDAM = vData
               
                 frmDAM.show vbModal
             End If
          End If
       End If
    End If
End If

End Sub

Private Sub cmdDelEfo_Click()
Dim x As Integer
With lvEFOrigem
    For x = 1 To .ListItems.Count
        .ListItems(x).Checked = False
    Next
End With

End Sub

Private Sub cmdDetalhe_Click()
Dim sNumProc As String, nInicio As Integer
Dim nParcela As Integer, nAno As Integer, nLancamento As Integer, nSequencia As Integer, nComplemento As Integer

If grdExtrato.Rows = 0 Then
    MsgBox "Não existem débitos.", vbExclamation, "Atenção"
    Exit Sub
End If

With grdExtrato
    If .SelectedRow = 0 Then Exit Sub
    nAno = .CellText(.SelectedRow, 1)
    nLancamento = Val(Left$(.CellText(.SelectedRow, 2), 3))
    nSequencia = Val(.CellText(.SelectedRow, 3))
    nParcela = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
    nComplemento = Val(.CellText(.SelectedRow, 5))
End With

With grdExtrato
       If .SelectedRow > 0 Then
            Unload frmCnsParcela
            Set frm = frmCnsParcela
            frm.nParcela = nParcela
            frm.nAno = nAno
            frm.nLancamento = nLancamento
            frm.nSequencia = nSequencia
            frm.nComplemento = nComplemento
            If .CellBackColor(.SelectedRow, 1) = 0 Then
                frm.nCodRed = Val(txtCod.Text)
            Else
                nInicio = InStr(1, .CellText(.SelectedRow, 2), " (", vbBinaryCompare)
                sNumProc = Mid(.CellText(.SelectedRow, 2), nInicio + 2, Len(.CellText(.SelectedRow, 2)) - Val(nInicio) - 2)
                Sql = "SELECT CODIGORESP FROM PROCESSOREPARC WHERE  NUMPROCESSO='" & sNumProc & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        frm.nCodRed = !CODIGORESP
                        frm.nResponsavel = 1
                    Else
                        frm.nCodRed = Val(txtCod.Text)
                        frm.nResponsavel = 0
                    End If
                   .Close
                End With
            End If
            frm.show vbModal
       End If
End With
End Sub

Private Sub cmdDoc_Click()
Dim nDV As Integer
Dim nParcela As Integer, nAno As Integer, nLancamento As Integer, nSequencia As Integer, nComplemento As Integer

With grdExtrato
    If .SelectedRow = 0 Then Exit Sub
    nAno = .CellText(.SelectedRow, 1)
    nLancamento = Val(Left$(.CellText(.SelectedRow, 2), 3))
    nSequencia = Val(.CellText(.SelectedRow, 3))
    nParcela = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
    nComplemento = Val(.CellText(.SelectedRow, 5))
End With

If grdExtrato.Rows = 0 Then Exit Sub
frBotao.Enabled = False
frDoc.Visible = True
frDoc.ZOrder 0
frTop.Enabled = False
grdExtrato.Enabled = False
grdDoc.Rows = 1
Sql = "SELECT PARCELADOCUMENTO.NUMDOCUMENTO,NUMDOCUMENTO.DATADOCUMENTO,NumDocumento.CODBANCO,NUMDOCUMENTO.VALORPAGO "
Sql = Sql & "FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO Where "
Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno
Sql = Sql & " AND CODLANCAMENTO=" & nLancamento
Sql = Sql & " AND NUMPARCELA=" & nParcela & " AND SEQLANCAMENTO=" & nSequencia & " AND CODCOMPLEMENTO=" & nComplemento
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    Do Until .EOF
       nDV = RetornaDVNumDoc(!NumDocumento)
       grdDoc.AddItem Format(!NumDocumento, "00000000") & sTr(nDV) & Chr(9) & Format(!DATADOCUMENTO, "dd/mm/yyyy") & Chr(9) & IIf(IsNull(!ValorPago), 0, FormatNumber(!ValorPago, 2))
      .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub cmdEF_Click()
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDA As String, sAj As String, sEF As String
Dim nSit As Integer, itmX As ListItem, z As Long, sVencto As String, Sql As String, RdoAux As rdoResultset

If Val(txtCod.Text) = 0 Then
    MsgBox "Selecione um contribuinte.", vbExclamation, "Atenção"
    Exit Sub
End If

If bFilterLoad Then
    MsgBox "Destive o filtro para acessar as execuções fiscais", vbCritical, "Atenção"
    Exit Sub
End If
cmdEFDoc_Click
cmbEF.Clear

Sql = "SELECT distinct processocnj FROM debitoparcela WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If Not IsNull(!ProcessoCNJ) Then
            cmbEF.AddItem !ProcessoCNJ
        End If
       .MoveNext
    Loop
   .Close
End With
If cmbEF.ListCount > 0 Then
    cmbEF.ListIndex = 0
Else
    CarregaOrigemEF
End If

EventosEF (True)
frBotao.Enabled = False
frEFiscal.Visible = True
frEFiscal.ZOrder 0
frTop.Enabled = False
grdExtrato.Enabled = False

End Sub

Private Sub cmdEFDoc_Click()
frEFDoc.Width = 6765
frEfObs.Width = 1320
frEFDoc.Left = 1440
lvDoc.Width = 6645
txtDocEF.Width = 1140
End Sub

Private Sub cmdEfObs_Click()
frEFDoc.Width = 1320
frEfObs.Width = 6765
frEFDoc.Left = 6895
lvDoc.Width = 1220
txtDocEF.Width = 6590
End Sub

Private Sub cmdEfQtde_Click()
Dim z As Variant
If sEventoEF = "" Then Exit Sub
On Error Resume Next

z = InputBox("Digite a quantidade para: " & lvDoc.SelectedItem.Text, "Informação requerida", lvDoc.SelectedItem.SubItems(1))

If Not IsNumeric(z) Then Exit Sub
lvDoc.SelectedItem.SubItems(1) = Val(z)
If Val(z) = 0 Then
    lvDoc.SelectedItem.Checked = False
Else
    lvDoc.SelectedItem.Checked = True
End If

End Sub

Private Sub cmdEFSit_Click()
Dim z As Variant
If sEventoEF = "" Then Exit Sub
On Error Resume Next

z = InputBox("Digite a situação para: " & lvDoc.SelectedItem.Text, "Informação requerida", lvDoc.SelectedItem.SubItems(2))

If z = "" Then
    lvDoc.SelectedItem.SubItems(2) = ""
Else
    lvDoc.SelectedItem.SubItems(2) = Left(z, 20)
End If

End Sub

Private Sub cmdExcluirEF_Click()
Dim Sql As String, x As Integer, z As Long, sNum As String, nNum As Integer, nAno As Integer
If cmbEF.ListIndex = -1 Then
    MsgBox "Selecione uma execução fiscal.", vbCritical, "Atenção"
    Exit Sub
End If

z = SendMessage(lvEFDest.hwnd, LVM_DELETEALLITEMS, 0, 0)

If cmbEF.Visible And cmbEF.ListIndex > -1 Then
    If MsgBox("Excluir a execução fiscal n° " & cmbEF.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
    
    sNum = cmbEF.Text
   ' nNum = Val(Left$(sNum, InStr(1, sNum, "/", vbBinaryCompare) - 1))
   ' nAno = Val(Right$(sNum, 4))
    
    Sql = "DELETE FROM EXECUCAOFISCAL WHERE processocnj='" & sNum & "'"
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM EXECUCAOFISCALDOC WHERE processocnj='" & sNum & "'"
    cn.Execute Sql, rdExecDirect
    
    Sql = "UPDATE DEBITOPARCELA SET processocnj=null WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND processocnj='" & sNum & "'"
    cn.Execute Sql, rdExecDirect
    With grdExtrato
        For Y = 1 To .Rows
            If .CellText(Y, 14) = cmbEF.Text Then
                .CellText(Y, 14) = ""
            End If
        Next
    End With
    txtDocEF.Text = ""
    cmbEF.Clear
    
    Sql = "SELECT * FROM EXECUCAOFISCAL WHERE CODREDUZIDO=" & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            cmbEF.AddItem !ProcessoCNJ
           .MoveNext
        Loop
       .Close
    End With
    If cmbEF.ListCount > 0 Then
        cmbEF.ListIndex = 0
    Else
        CarregaOrigemEF
    End If
            
End If


End Sub

Private Sub cmdExcluirObs_Click()
Dim sNomeUser As String, nSeq As Integer
Dim nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer
If txtObservacao.Text = "" Then
    MsgBox "Selecione o registro que deseja excluir.", vbExclamation, "Atenção"
Else
    sNomeUser = UCase(lvObserv.SelectedItem.SubItems(1))
    If sNomeUser <> NomeDeLogin And Right(sNomeUser, Len(NomeDeLogin)) <> UCase(NomeDeLogin) Then
        MsgBox "Você não pode excluir uma observação criada por outro usuário.", vbCritical, "ALERTA DE SEGURANÇA"
    Else
        If MsgBox("Excluir esta observação ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            nSeq = Val(lvObserv.SelectedItem.Text)
            If bObs Then
                Sql = "DELETE FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND SEQ=" & nSeq
                cn.Execute Sql, rdExecDirect
                CarregaObs
            Else
                With grdExtrato
                    If .Rows = 0 Then
                        MsgBox "Não existem débitos.", vbExclamation, "Atenção"
                        Exit Sub
                    End If
                    If .SelectedRow = 0 Then Exit Sub
                    nAno = .CellText(.SelectedRow, 1)
                    nLanc = Val(Left$(.CellText(.SelectedRow, 2), 3))
                    nSeqLanc = Val(.CellText(.SelectedRow, 3))
                    nParc = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
                    nComp = Val(.CellText(.SelectedRow, 5))
                End With
                Sql = "DELETE FROM OBSPARCELA  WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno
                Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nParc
                Sql = Sql & " AND CODCOMPLEMENTO=" & nComp & " AND SEQ=" & nSeq
                cn.Execute Sql, rdExecDirect
                CarregaObsParcela
            End If
            If lvObserv.ListItems.Count > 0 Then
                lvObserv.ListItems(1).Selected = True
                lvObserv_Click
            End If
        End If
    End If
End If
Liberado
End Sub

Private Sub cmdExtrato_Click()
If Val(txtCod.Text) = 0 Then
    MsgBox "Selecione um contribuinte.", vbExclamation, "Atenção"
    Exit Sub
End If

lIndex = m_cMenuExtrato.ShowPopupMenu(cmdExtrato.Left + frBotao.Left, cmdExtrato.Top, cmdExtrato.Left, cmdExtrato.Top, Me.ScaleWidth - cmdExtrato.Left - cmdExtrato.Width, cmdExtrato.Top + cmdExtrato.Height, False)

End Sub

Private Sub GravaExtrato()

Dim aDebito() As Debito
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim Achou As Boolean, x As Integer
Dim nSomaDebito As Double, nEval As Integer, nValorCorrecao As Double
Dim nSomaVencer As Double, nValorAtualizado As Double
Dim bMulta As Boolean, bJuros As Boolean
Dim qd As New rdoQuery
Dim sComputer As String
Dim nSeq As Integer
Dim nCodReduz As Long
Dim sNumInsc As String
Dim sNomeProp As String
Dim sEnd As String
Dim nNumero As Integer
Dim sBairro As String
Dim nCodBanco As Integer

Dim nValorMulta As Double
Dim nValorJuros As Double
Dim nSaldo As Double
Dim sDA As String, sAj As String
Dim nAno As Integer

ReDim aDebito(0)
nSomaDebito = 0
nSomaVencer = 0
bSel = True

sComputer = NomeDeLogin
nCodReduz = Val(txtCod.Text)
sNomeProp = lblProp.Caption
sEnd = lblRua.Caption
sNumInsc = lblNumInsc.Caption

'MORREK ZMANI
Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT * FROM VWCNSLANCAMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODTRIBUTO<>3 "
If FiltroE > 0 Then
   Sql = Sql & " AND ANOEXERCICIO BETWEEN " & FiltroE & " AND " & FiltroE2
End If
If FiltroL > 0 Then
   Sql = Sql & " AND CODLANCAMENTO=" & FiltroL
End If
If FiltroS > 0 Then
'If frmFiltroDebito.Visible = True And frmFiltroDebito.cmbS.ListIndex > 0 Then
   Sql = Sql & " AND STATUSLANC=" & FiltroS
'End If
End If
If bFiltroSEQ Then
   Sql = Sql & " AND SEQLANCAMENTO=" & FiltroSEQ
End If
If FiltroLP <> "" Then
   Sql = Sql & " AND NUMPROCESSO='" & FiltroLP & "'"
End If
If FiltroD = "S" Then
   Sql = Sql & " AND DATAINSCRICAO IS NOT NULL"
ElseIf FiltroD = "N" Then
   Sql = Sql & " AND DATAINSCRICAO IS NULL"
End If
If FiltroA = "S" Then
   Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
ElseIf FiltroA = "N" Then
   Sql = Sql & " AND DATAAJUIZA IS NULL "
End If
Sql = Sql & " ORDER BY ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,YEAR(DATAVENCIMENTO),MONTH(DATAVENCIMENTO),DAY(DATAVENCIMENTO) "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
        nEval = UBound(aDebito)
        Do Until .EOF
          '  If !AnoExercicio = 2008 And !NumParcela = 12 Then MsgBox "teste"
        
        
            bJuros = False: bMulta = False
'            If !SeqLancamento = 6 Then MsgBox "teste"
            If !CodLancamento = 20 And !statuslanc = 5 Then GoTo proximo
            If !NumParcela = 0 And !statuslanc = 5 Then GoTo proximo
            If !statuslanc = 12 Or !statuslanc = 5 Then GoTo proximo
            'Carrega Matriz Debito
            nEval = UBound(aDebito)
            If !NumParcela = 0 And !statuslanc = 3 And DateDiff("d", !DataVencimento, Now) > 0 Then GoTo proximo
            If !AnoExercicio = 2003 And !CodLancamento = 1 And (!statuslanc <> 2 And !statuslanc <> 1) Then GoTo proximo
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aDebito(UBound(aDebito) + 1)
                nEval = UBound(aDebito)
                aDebito(nEval).nAno = !AnoExercicio
                aDebito(nEval).nLanc = !CodLancamento
                If !CodLancamento = 20 Then
                   If Not IsNull(!numprocesso) Then
                      aDebito(nEval).sLanc = !descreduz & " (" & !numprocesso & ")"
                   Else
                      aDebito(nEval).sLanc = !descreduz
                   End If
                Else
                   aDebito(nEval).sLanc = !descreduz
                End If
                aDebito(nEval).nSeq = !SeqLancamento
                aDebito(nEval).nParc = !NumParcela
                aDebito(nEval).nCompl = !CODCOMPLEMENTO
                aDebito(nEval).nSituacao = !statuslanc
                aDebito(nEval).sSituacao = !DescSituacao
                aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aDebito(nEval).nCodTributo = !CodTributo
                aDebito(nEval).nValorTributo = FormatNumber(!TOTALLANCADO, 2)
                If !CodTributo <> 3 Then
                    If !statuslanc = 3 Then
                        If DateDiff("d", !DataVencimento, Now) > 0 Then
                            nValorAtualizado = !TOTALLANCADO
                            nValorCorrecao = CalculaCorrecao(nValorAtualizado, !DataVencimento)
                            aDebito(nEval).nValorCorrecao = nValorCorrecao
                            Sql = "SELECT MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                If .RowCount > 0 Then
                                    bJuros = !Juros
                                    bMulta = !Multa
                                End If
                               .Close '
                            End With
                            If bJuros Then
                              nValorAtualizado = nValorAtualizado + CDbl(CalculaJuros(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                              aDebito(nEval).nValorJuros = CDbl(CalculaJuros(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                            End If
                            If bMulta Then
                               nValorAtualizado = nValorAtualizado + CDbl(CalculaMulta(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                               aDebito(nEval).nValorMulta = aDebito(nEval).nValorMulta + CDbl(CalculaMulta(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                            End If
                            nValorAtualizado = nValorAtualizado + nValorCorrecao
                        Else
                            nValorAtualizado = !TOTALLANCADO
                        End If
                       aDebito(nEval).nCodBanco = 0
                       aDebito(nEval).dDataPag = "01/01/1900"
                    Else
                        Sql = "SELECT * from DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND  CODLANCAMENTO=" & !CodLancamento
                        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
                        Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                aDebito(nEval).nValorJuros = RdoAux!ValorJuros
                                aDebito(nEval).nValorMulta = RdoAux!ValorMulta
                                aDebito(nEval).nValorCorrecao = RdoAux!ValorCorrecao
                                aDebito(nEval).nCodBanco = Val(SubNull(!CodBanco))
                                aDebito(nEval).dDataPag = Format(!DataPagamento, "dd/mm/yyyy")
                            Else
                                aDebito(nEval).nValorJuros = 0
                                aDebito(nEval).nValorMulta = 0
                                aDebito(nEval).nValorCorrecao = 0
                                aDebito(nEval).nCodBanco = 0
                                aDebito(nEval).dDataPag = "01/01/1900"
                            End If
                           .Close
                        End With
                        nValorAtualizado = !TOTALLANCADO + RdoAux!ValorJuros + RdoAux!ValorMulta + RdoAux!ValorCorrecao
                    End If
                Else
                    aDebito(nEval).nCodBanco = 0
                    aDebito(nEval).dDataPag = "01/01/1900"
                    nValorAtualizado = 0
                End If
               
                aDebito(nEval).nValorAtual = FormatNumber(nValorAtualizado, 2)
            Else
                If aDebito(x).nCodTributo = !CodTributo Then GoTo proximo
            
                aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !TOTALLANCADO, 2)
                If !CodTributo <> 3 Then
                    Sql = "SELECT MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            bJuros = !Juros
                            bMulta = !Multa
                        End If
                       .Close '
                    End With
                    If !statuslanc = 3 Then
                        If DateDiff("d", !DataVencimento, Now) > 0 Then
                            nValorAtualizado = !TOTALLANCADO
                            nValorCorrecao = CalculaCorrecao(!TOTALLANCADO, !DataVencimento)
                            aDebito(nEval).nValorCorrecao = aDebito(nEval).nValorCorrecao + nValorCorrecao
                            Sql = "SELECT MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                bJuros = !Juros
                                bMulta = !Multa
                               .Close
                            End With
                            If bJuros Then
                               nValorAtualizado = nValorAtualizado + CDbl(CalculaJuros(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                               aDebito(nEval).nValorJuros = aDebito(nEval).nValorJuros + CDbl(CalculaJuros(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                            End If
                            If bMulta Then
                               nValorAtualizado = nValorAtualizado + CDbl(CalculaMulta(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                               aDebito(nEval).nValorMulta = aDebito(nEval).nValorMulta + CDbl(CalculaMulta(!TOTALLANCADO + nValorCorrecao, !DataVencimento))
                            End If
                            nValorAtualizado = nValorAtualizado + nValorCorrecao
                        Else
                            nValorAtualizado = !TOTALLANCADO
                        End If
                        aDebito(nEval).nCodBanco = 0
                        aDebito(nEval).dDataPag = "01/01/1900"
                    ElseIf !statuslanc < 3 Or !statuslanc = 7 Then
                        Sql = "SELECT * from DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND  CODLANCAMENTO=" & !CodLancamento
                        Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
                        Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                aDebito(nEval).nCodBanco = Val(SubNull(!CodBanco))
                                aDebito(nEval).dDataPag = Format(!DataPagamento, "dd/mm/yyyy")
                            Else
                                aDebito(nEval).nCodBanco = 0
                                aDebito(nEval).dDataPag = "01/01/1900"
                            End If
                           .Close
                        End With
                        aDebito(nEval).nValorCorrecao = aDebito(nEval).nValorCorrecao + !ValorCorrecao
                        If bJuros Then
                            aDebito(nEval).nValorJuros = aDebito(nEval).nValorJuros + !ValorJuros
                        End If
                        If bMulta Then
                            aDebito(nEval).nValorMulta = aDebito(nEval).nValorMulta + !ValorMulta
                        End If
                        nValorAtualizado = !TOTALLANCADO + !ValorCorrecao + !ValorJuros + !ValorMulta
                    Else
                        aDebito(nEval).nValorCorrecao = aDebito(nEval).nValorCorrecao + !ValorCorrecao
                        If bJuros Then
                            aDebito(nEval).nValorJuros = aDebito(nEval).nValorJuros + !ValorJuros
                        End If
                        If bMulta Then
                            aDebito(nEval).nValorMulta = aDebito(nEval).nValorMulta + !ValorMulta
                        End If
                        aDebito(nEval).nCodBanco = 0
                        aDebito(nEval).dDataPag = "01/01/1900"
                        nValorAtualizado = !TOTALLANCADO + !ValorCorrecao + !ValorJuros + !ValorMulta
                    End If
                Else
                    If !statuslanc = 3 Then
                        nValorAtualizado = !TOTALLANCADO
                    Else
                        nValorAtualizado = 0
                    End If
                End If
                
                aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + nValorAtualizado, 2)
            End If
            If aDebito(nEval).dDataPag = "00:00:00" Then aDebito(nEval).dDataPag = "01/01/1900"
proximo:
            .MoveNext
        Loop
      End If
   .Close
End With

nSeq = 0
nSaldo = 0
Set qd.ActiveConnection = cn

For x = 1 To UBound(aDebito)
    With aDebito(x)
        nSeq = nSeq + 1
       'TAKLIT ET A NETUNIM
        Sql = "INSERT EXTRATOTMP(COMPUTER,SEQ,CODREDUZIDO,NOMEPROP,ENDERECO,NUMERO,BAIRRO,CODLANCAMENTO,DESCLANCAMENTO,"
        Sql = Sql & "ANOEXERCICIO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,DATAVENCIMENTO,CODBANCO,DATAPAGAMENTO,STATUSLANC,"
        Sql = Sql & "VALORLANCADO,VALORCORRECAO,VALORMULTA,VALORJUROS,VALORTOTAL,SALDO,DA,AJ) VALUES('" & NomeDeLogin & "',"
        Sql = Sql & nSeq & "," & Val(txtCod.Text) & ",'" & Mask(Left$(sNomeProp, 30)) & "','" & Left$(sEnd, 50) & "'," & nNumero & ",'" & sBairro & "',"
        Sql = Sql & .nLanc & ",'" & .sLanc & "'," & .nAno & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
        Sql = Sql & Format(.sVencto, "mm/dd/yyyy") & "'," & .nCodBanco & ",'" & Format(.dDataPag, "mm/dd/yyyy") & "','"
        Sql = Sql & Left$(Format(.nSituacao, "00") & " - " & .sSituacao, 30) & "'," & Virg2Ponto(CStr(.nValorTributo)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorCorrecao)) & "," & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorJuros)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorAtual)) & "," & Virg2Ponto(CStr(nSaldo)) & ",'" & .sDA & "','" & .sAj & "')"
        cn.Execute Sql, rdExecDirect
    End With
Next

'EXIBE RELATORIO
If InStr(1, cn.Connect, "Tributacao_", vbBinaryCompare) > 0 Then
    frmReport.ShowReport "ExtratoFull", frmMdi.hwnd, Me.hwnd
Else
    frmReport.ShowReport "Extrato", frmMdi.hwnd, Me.hwnd
End If
Liberado
'MORREK ZMANI
Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
   
End Sub

Private Sub GravaExtrato2(bForum As Boolean)

Dim aDebito() As Debito, z1 As Variant
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim Achou As Boolean, x As Integer, z As Integer
Dim nSomaDebito As Double, nEval As Integer, nValorCorrecao As Double
Dim nSomaVencer As Double, nValorAtualizado As Double
Dim bMulta As Boolean, bJuros As Boolean
Dim qd As New rdoQuery
Dim sComputer As String
Dim nSeq As Integer
Dim nCodReduz As Long
Dim sNumInsc As String
Dim sNomeProp As String
Dim sEnd As String
Dim nAno As Integer
Dim nLancamento As Integer
Dim nSequencia As Integer
Dim nParcela As Integer
Dim nComplemento As Integer

Dim nNumero As Integer
Dim sBairro As String
Dim nCodBanco As Integer

Dim nValorMulta As Double
Dim nValorJuros As Double
Dim nSaldo As Double
Dim sDA As String, sAj As String

ReDim aDebito(0)
nSomaDebito = 0
nSomaVencer = 0
bSel = True

sComputer = NomeDeLogin
nCodReduz = Val(txtCod.Text)
sNomeProp = lblProp.Caption
sEnd = lblRua.Caption
sNumInsc = lblNumInsc.Caption

z1 = InputBox("Digite a data de emissão do extrato", "Data do Extrato", Right$(frmMdi.Sbar.Panels(6).Text, 10))
If Not IsDate(z1) Then
    MsgBox "Data inválida", vbCritical, "Erro"
    Exit Sub
Else
    dDataAtualiza = CDate(z1)
End If

'MORREK ZMANI
Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = Val(txtCod.Text)
qd(1) = Val(txtCod.Text)
If cmbAno1.ListCount > 0 Then
    qd(2) = cmbAno1.Text: qd(3) = cmbAno2.Text
Else
    qd(2) = 1950: qd(3) = 2050
End If
If cmbLanc.ListIndex > 0 Then
    qd(4) = Val(Left(cmbLanc.Text, 3)): qd(5) = Val(Left(cmbLanc.Text, 3)) 'LANCAMENTO
Else
    qd(4) = 0: qd(5) = 99
End If
If cmbSeq.ListIndex > 0 Then
    qd(6) = Val(cmbSeq.Text): qd(7) = Val(cmbSeq.Text) 'SEQUENCIA
Else
    qd(6) = 0: qd(7) = 9999
End If
qd(8) = 0: qd(9) = 99
qd(10) = 0: qd(11) = 99
If cmbSit.ListIndex > 0 Then
    qd(12) = Val(Left(cmbSit.Text, 2)): qd(13) = Val(Left(cmbSit.Text, 2)) 'STATUSLANC
Else
    qd(12) = 0: qd(13) = 99
End If
qd(14) = Format(dDataAtualiza, "mm/dd/yyyy")
qd(15) = NomeDeLogin
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
        nEval = UBound(aDebito)
        Do Until .EOF
            
        
            bJuros = False: bMulta = False
            If cmbAj.ListIndex = 1 Then
                If IsNull(!dataajuiza) Then GoTo proximo
            End If
            If cmbAj.ListIndex = 2 Then
                If Not IsNull(!dataajuiza) Then GoTo proximo
            End If
            If cmbDA.ListIndex = 1 Then
                If IsNull(!datainscricao) Then GoTo proximo
            End If
            If cmbDA.ListIndex = 2 Then
                If Not IsNull(!datainscricao) Then GoTo proximo
            End If
            
'            If !CodLancamento = 20 And !statuslanc = 5 Then GoTo Proximo
'            If !NumParcela = 0 And !statuslanc = 5 Then GoTo Proximo
'            If !statuslanc = 12 Or !statuslanc = 5 Then GoTo Proximo
            
'            If !CodLancamento = 20 Then GoTo Proximo
            If !NumParcela = 0 And !statuslanc > 2 Then GoTo proximo
            If !statuslanc = 12 Then GoTo proximo
            
            
            If !NumParcela > 0 And !statuslanc = 1 Then GoTo proximo
            'Carrega Matriz Debito
            nEval = UBound(aDebito)
            If !NumParcela = 0 And !statuslanc = 3 And DateDiff("d", !DataVencimento, Now) > 0 Then GoTo proximo
            If !AnoExercicio = 2003 And !CodLancamento = 1 And (!statuslanc <> 2 And !statuslanc <> 1) Then GoTo proximo
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aDebito(UBound(aDebito) + 1)
                nEval = UBound(aDebito)
                aDebito(nEval).nAno = !AnoExercicio
                aDebito(nEval).nLanc = !CodLancamento
                If !CodLancamento = 20 Or !CodLancamento = 8 Then
                   If Not IsNull(!numprocesso) Then
                      If Val(Right$(!numprocesso, 4)) >= 2006 Then
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1) & "-" & RetornaDVProcesso(Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1)) & "/" & Right$(!numprocesso, 4) & ")"
                      Else
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & !numprocesso & ")"
                      End If
                   Else
                      aDebito(nEval).sLanc = !DESCLANCAMENTO
                   End If
                Else
                   aDebito(nEval).sLanc = !DESCLANCAMENTO
                End If
                aDebito(nEval).nSeq = !SeqLancamento
                aDebito(nEval).nParc = !NumParcela
                aDebito(nEval).nCompl = !CODCOMPLEMENTO
                aDebito(nEval).nSituacao = !statuslanc
                aDebito(nEval).sSituacao = !Situacao
                aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aDebito(nEval).nCodTributo = !CodTributo
                aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                
                
                If bForum Then
          '      If (!statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7) And Not bForum Then
          '          aDebito(nEval).nValorAtual = FormatNumber(!valortributo, 2)
          '      ElseIf (!statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7) And bForum Then
                    '***************************************************************
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,VALORTRIBUTO,ABREVTRIBUTO,DEBITOTRIBUTO.VALORJUROS,DEBITOTRIBUTO.VALORCORRECAO,DEBITOTRIBUTO.VALORMULTA "
                    Sql = Sql & "FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO "
                    Sql = Sql & "AND DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO "
                    Sql = Sql & "AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND DEBITOPARCELA.NumParcela = DEBITOTRIBUTO.NumParcela "
                    Sql = Sql & "AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO Inner Join TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO "
                    Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio & " AND DEBITOPARCELA.CODLANCAMENTO = " & !CodLancamento & " AND "
                    Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO = " & !SeqLancamento & " AND DEBITOPARCELA.NUMPARCELA = " & !NumParcela & " AND DEBITOPARCELA.CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO=" & !CodTributo
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
                        aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                        aDebito(nEval).nValorAtual = FormatNumber(!ValorTributo + !ValorJuros + !ValorMulta + !ValorCorrecao, 2)
                       .Close
                    End With
                    '***************************************************************
                Else
                    If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
                        If Not IsNull(!valorpagoreal) Then
                            aDebito(nEval).nValorAtual = FormatNumber(!valorpagoreal, 2)
                        Else
                            aDebito(nEval).nValorAtual = FormatNumber(0, 2)
                        End If
                    Else
                        aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                    End If
                    
                    aDebito(nEval).nValorJuros = IIf(IsNull(!ValorJuros), 0, FormatNumber(!ValorJuros, 2))
                    aDebito(nEval).nValorMulta = IIf(IsNull(!ValorMulta), 0, FormatNumber(!ValorMulta, 2))
                    aDebito(nEval).nValorCorrecao = IIf(IsNull(!ValorCorrecao), 0, FormatNumber(!ValorCorrecao, 2))
                End If
            Else
                If aDebito(nEval).nCodTributo = !CodTributo Then GoTo proximo
            
                aDebito(nEval).nValorTributo = FormatNumber(aDebito(nEval).nValorTributo + !ValorTributo, 2)
               If (!statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7) And Not bForum Then
                    'aDebito(nEval).nValorAtual = FormatNumber(aDebito(nEval).nValorAtual + !ValorTributo, 2)
                    aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                    If Not IsNull(!ValorMulta) Then
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                    Else
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + 0, 2)
                    End If
                    aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
               ElseIf (!statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7) And bForum Then
                    '***************************************************************
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,VALORTRIBUTO,ABREVTRIBUTO,DEBITOTRIBUTO.VALORJUROS,DEBITOTRIBUTO.VALORCORRECAO,DEBITOTRIBUTO.VALORMULTA "
                    Sql = Sql & "FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO "
                    Sql = Sql & "AND DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO "
                    Sql = Sql & "AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND DEBITOPARCELA.NumParcela = DEBITOTRIBUTO.NumParcela "
                    Sql = Sql & "AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO Inner Join TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO "
                    Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio & " AND DEBITOPARCELA.CODLANCAMENTO = " & !CodLancamento & " AND "
                    Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO = " & !SeqLancamento & " AND DEBITOPARCELA.NUMPARCELA = " & !NumParcela & " AND DEBITOPARCELA.CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO=" & !CodTributo
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
                        aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                        aDebito(nEval).nValorAtual = FormatNumber(aDebito(nEval).nValorAtual + !ValorTributo + !ValorJuros + !ValorMulta + !ValorCorrecao, 2)
                       .Close
                    End With
                    '***************************************************************
               Else
                    aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                    aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                    aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
                    aDebito(nEval).nValorAtual = FormatNumber(aDebito(nEval).nValorAtual + !ValorTotal, 2)
               End If
            End If
            If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
                Sql = "SELECT * FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND RESTITUIDO IS NULL"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                         aDebito(nEval).dDataPag = Format(!DataPagamento, "dd/mm/yyyy")
                         aDebito(nEval).nCodBanco = Val(SubNull(!CodBanco))
                    End If
                End With
            End If
proximo:
            .MoveNext
        Loop
      End If
   .Close
End With

nSeq = 0
nSaldo = 0
'Set qd.ActiveConnection = cn

For x = 1 To UBound(aDebito)
    With aDebito(x)
        If bExtrato Then
            Achou = False
            For z = 1 To grdExtrato.Rows
                If grdExtrato.CellBackColor(z, 2) <> &HC0FFC0 And nExtrato = 2 Then GoTo ProximoEx
                nAno = grdExtrato.CellText(z, 1)
                nLancamento = Val(Left$(grdExtrato.CellText(z, 2), 3))
                nSequencia = Val(grdExtrato.CellText(z, 3))
                nParcela = IIf(grdExtrato.CellText(z, 4) = "Unica", 0, grdExtrato.CellText(z, 4))
                nComplemento = Val(grdExtrato.CellText(z, 5))
                If nAno = .nAno And nLancamento = .nLanc And nSequencia = .nSeq And nParcela = .nParc And nComplemento = .nCompl Then
                    Achou = True
                    Exit For
                End If
ProximoEx:
            Next
            If Not Achou Then GoTo PROXIMO2
        End If
    
    
        If .dDataPag = "00:00:00" Then .dDataPag = "01/01/1900"
        nSeq = nSeq + 1
       'TAKLIT ET A NETUNIM
        If InStr(1, cn.Connect, "Tributacao_", vbBinaryCompare) > 0 Then
            Sql = "INSERT Tributacao_FULL..EXTRATOTMP"
        Else
            Sql = "INSERT Tributacao..EXTRATOTMP"
        End If
        Sql = Sql & "(COMPUTER,SEQ,CODREDUZIDO,NOMEPROP,ENDERECO,NUMERO,BAIRRO,CODLANCAMENTO,DESCLANCAMENTO,"
        Sql = Sql & "ANOEXERCICIO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,DATAVENCIMENTO,CODBANCO,DATAPAGAMENTO,STATUSLANC,"
        Sql = Sql & "VALORLANCADO,VALORCORRECAO,VALORMULTA,VALORJUROS,VALORTOTAL,SALDO,DA,AJ) VALUES('" & NomeDeLogin & "',"
        Sql = Sql & nSeq & "," & Val(txtCod.Text) & ",'" & Mask(Left$(sNomeProp, 30)) & "','" & Left$(sEnd, 50) & "'," & nNumero & ",'" & sBairro & "',"
        Sql = Sql & .nLanc & ",'" & .sLanc & "'," & .nAno & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
        Sql = Sql & Format(.sVencto, "mm/dd/yyyy") & "'," & .nCodBanco & ",'" & Format(.dDataPag, "mm/dd/yyyy") & "','"
        Sql = Sql & Left$(Format(.nSituacao, "00") & " - " & .sSituacao, 30) & "'," & Virg2Ponto(CStr(.nValorTributo)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorCorrecao)) & "," & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorJuros)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorAtual)) & "," & Virg2Ponto(CStr(nSaldo)) & ",'" & .sDA & "','" & .sAj & "')"
        cn.Execute Sql, rdExecDirect
    End With
PROXIMO2:
Next

'EXIBE RELATORIO
If InStr(1, cn.Connect, "Tributacao_", vbBinaryCompare) > 0 Then
    frmReport.ShowReport "ExtratoFull", frmMdi.hwnd, Me.hwnd
'MORREK ZMANI
    Sql = "DELETE FROM Tributacao_full..EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
Else
    If bForum Then
        frmReport.ShowReport "ExtratoForum", frmMdi.hwnd, Me.hwnd
    Else
        frmReport.ShowReport "Extrato", frmMdi.hwnd, Me.hwnd
    End If
    'MORREK ZMANI
    Sql = "DELETE FROM Tributacao..EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
End If
Liberado

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
   
End Sub



Private Sub cmdFilter_Click()

'bCarregado = False
'If grdExtrato.Rows = 0 Then
'   cmdFilter.Value = False
'   Exit Sub
'End If
'If cmdFilter.Value = True Then
'   frmFiltroDebito.show vbModeless, frmMdi
'Else
'   Unload frmFiltroDebito
'End If


If cmdFilter.value = False Then
    pnlFilter.Visible = False
    grdExtrato.Height = 4875
Else
    CarregaFiltro
    pnlFilter.Visible = True
    grdExtrato.Height = 4125
End If

End Sub


Private Sub CarregaFiltro()
Dim x As Integer, Y As Integer, nAno As Integer, bAchou As Boolean, aAno() As Integer, aLanc() As String, aStatus() As String, aSeq() As String
Ocupado

If bFilterLoad Then GoTo fim
ReDim aAno(0): cmbAno1.Clear: cmbAno2.Clear: cmbLanc.Clear: cmbSit.Clear
ReDim aLanc(1): ReDim aStatus(1): ReDim aSeq(1)
With grdExtrato
    For x = 1 To .Rows
        nAno = Val(.CellText(x, 1))
        bAchou = False
        For Y = 1 To UBound(aAno)
            If nAno = aAno(Y) Then
                bAchou = True
            End If
        Next
        'Next
        If Not bAchou Then
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = nAno
        End If
    
        aLanc(1) = "{Todos}"
        bAchou = False
        For Y = 1 To UBound(aLanc)
            If Val(Left(.CellText(x, 2), 3)) = 20 Then
                If "20-REPARCELAMENTO" = aLanc(Y) Then
                    bAchou = True
                End If
            Else
                If .CellText(x, 2) = aLanc(Y) Then
                    bAchou = True
                End If
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aLanc(UBound(aLanc) + 1)
            If Val(Left(.CellText(x, 2), 3)) = 20 Then
                aLanc(UBound(aLanc)) = "20-REPARCELAMENTO"
            Else
                aLanc(UBound(aLanc)) = .CellText(x, 2)
            End If
        End If
    
        aStatus(1) = "{Todos}"
        bAchou = False
        For Y = 1 To UBound(aStatus)
            If .CellText(x, 6) = aStatus(Y) Then
                bAchou = True
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aStatus(UBound(aStatus) + 1)
            aStatus(UBound(aStatus)) = .CellText(x, 6)
        End If
    
        aSeq(1) = "{Todos}"
        bAchou = False
        For Y = 2 To UBound(aSeq)
            If .CellText(x, 3) = aSeq(Y) Then
                bAchou = True
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aSeq(UBound(aSeq) + 1)
            aSeq(UBound(aSeq)) = .CellText(x, 3)
        End If

    
    Next
    
    

End With

bExecF = False
For x = 1 To UBound(aAno)
    cmbAno1.AddItem aAno(x)
    cmbAno2.AddItem aAno(x)
Next

For x = 1 To UBound(aLanc)
    cmbLanc.AddItem (aLanc(x))
Next

For x = 1 To UBound(aStatus)
    cmbSit.AddItem (aStatus(x))
Next

For x = 1 To UBound(aSeq)
    cmbSeq.AddItem (aSeq(x))
Next

If cmbAno1.ListCount > 0 Then cmbAno1.ListIndex = 0
If cmbAno2.ListCount > 0 Then cmbAno2.ListIndex = cmbAno2.ListCount - 1
If cmbLanc.ListCount > 0 Then cmbLanc.ListIndex = 0
If cmbSit.ListCount > 0 Then cmbSit.ListIndex = 0
If cmbSeq.ListCount > 0 Then cmbSeq.ListIndex = 0
cmbDA.ListIndex = 0
cmbAj.ListIndex = 0
bFilterLoad = True
fim:
bExecF = True
Liberado

End Sub

Private Sub cmdGravarEF_Click()
Dim nExercicio As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim sNum As String, nAno As Integer, nNum As Integer
Dim RdoAux As rdoResultset, Sql As String, x As Integer, Y As Integer

If lvEFDest.ListItems.Count = 0 Then
    MsgBox "Nenhum débito selecionado.", vbCritical, "Erro"
    Exit Sub
End If
If Trim(txtEF.Text) = "" And sEventoEF = "Novo" Then
    MsgBox "Digite o nº da Exec.fiscal.", vbCritical, "Erro"
    Exit Sub
End If

If sEventoEF = "Novo" Then
    sNum = Trim(txtEF.Text)
Else
    sNum = cmbEF.Text
End If
'nNum = Val(Left$(sNum, InStr(1, sNum, "/", vbBinaryCompare) - 1))
'nAno = Val(Right$(sNum, 4))

If sNum = "" Then
    MsgBox "Nº da exec.fiscal inválido.", vbCritical, "Erro"
    Exit Sub
End If
'If nAno < 1990 Or nAno > 2020 Then
'    MsgBox "Nº da exec.fiscal inválido.", vbCritical, "Erro"
'    Exit Sub
'End If

If sEventoEF = "Novo" Then
    Sql = "SELECT * FROM EXECUCAOFISCAL WHERE processocnj='" & sNum & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            nSeq = !Seq
             MsgBox "Execução fiscal já cadastrada para o código " & !CODREDUZIDO, vbCritical, "Erro"
          '  .Close
'            Exit Sub
        End If
       .Close
    End With
    
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM EXECUCAOFISCAL WHERE processocnj='" & sNum & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq = 0
    Else
        nSeq = RdoAux!maximo + 1
    End If
    RdoAux.Close
    
    Sql = "INSERT EXECUCAOFISCAL(NUMEXEC,ANOEXEC,processocnj,SEQ,CODREDUZIDO,OBS) VALUES("
    Sql = Sql & 0 & "," & 0 & ",'" & sNum & "'," & nSeq & "," & Val(txtCod.Text) & ",'" & Mask(txtDocEF.Text) & "')"
    cn.Execute Sql, rdExecDirect
Else
    Sql = "UPDATE EXECUCAOFISCAL SET OBS='" & Mask(txtDocEF.Text) & "' WHERE processocnj='" & sNum & "' AND SEQ=" & nSeq & " AND CODREDUZIDO=" & Val(txtCod.Text)
    cn.Execute Sql, rdExecDirect
End If

'apagamos todas os lancamentos com esta execução fical e gravamos de novo
If sEventoEF = "Alterar" Then
    Sql = "UPDATE DEBITOPARCELA SET processocnj=NULL WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND processocnj='" & sNum & "'"
    cn.Execute Sql, rdExecDirect
    With grdExtrato
        For Y = 1 To .Rows
            If .CellText(Y, 14) = cmbEF.Text Then
                .CellText(Y, 14) = ""
            End If
        Next
    End With
End If

With lvEFDest
    For x = 1 To .ListItems.Count
        
        nExercicio = .ListItems(x).Text
        nLanc = .ListItems(x).SubItems(1)
        nSeq = .ListItems(x).SubItems(2)
        nParc = .ListItems(x).SubItems(3)
        nCompl = .ListItems(x).SubItems(4)
        
        With grdExtrato
            For Y = 1 To .Rows
                If Val(.CellText(Y, 1)) = nExercicio And Val(Left(.CellText(Y, 2), 3)) = nLanc And Val(.CellText(Y, 3)) = nSeq And Val(.CellText(Y, 4)) = nParc And Val(.CellText(Y, 5)) = nCompl Then
                    .CellText(Y, 14) = Format(nNum, "00000") & "/" & nAno
                    Sql = "UPDATE DEBITOPARCELA SET processocnj='" & sNum & "' WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
                    Sql = Sql & " ANOEXERCICIO=" & nExercicio & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                    cn.Execute Sql, rdExecDirect
                    
                End If
            Next
        End With
    
    Next
End With

Sql = "DELETE FROM EXECUCAOFISCALDOC WHERE processocnj='" & sNum & "'"
cn.Execute Sql, rdExecDirect

With lvDoc
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked Then
            Sql = "INSERT EXECUCAOFISCALDOC(ANOEXEC,NUMEXEC,PROCESSOCNJ,NUMDOC,QTDE,SITUACAO) VALUES(" & 0 & "," & 0 & ",'" & sNum & "'," & Val(Right(.ListItems(x).Key, 3)) & "," & .ListItems(x).SubItems(1) & ",'" & Mask(.ListItems(x).SubItems(2)) & "')"
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With


If sEventoEF = "Novo" Then
    cmbEF.AddItem sNum
    cmbEF.ListIndex = cmbEF.ListCount - 1
End If
sEventoEF = ""
EventosEF True
End Sub

Private Sub cmdGravarObs_Click()
Dim nSeq As Integer, sData As String, i As Integer
Dim nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer

If txtObservacao.Text = "" Then
    MsgBox "Digite a observação", vbExclamation, "Atenção"
Else
    If bObs Then 'SE FOR OBSERVACAO GERAL
        If bNovoObs Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & Val(txtCod.Text)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USUARIO,DATAOBS,OBS) VALUES(" & Val(txtCod.Text) & "," & nSeq & ",'"
'            Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "','" & Mask(txtObservacao.Text) & "')"
            Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USERID,DATAOBS,OBS) VALUES(" & Val(txtCod.Text) & "," & nSeq & ","
            Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "','" & Mask(txtObservacao.Text) & "')"
            cn.Execute Sql, rdExecDirect
            CarregaObs
            lvObserv.ListItems(lvObserv.ListItems.Count).Selected = True
        Else
            nSeq = Val(lvObserv.SelectedItem.Text)
            sData = lvObserv.SelectedItem.SubItems(2)
            Sql = "UPDATE DEBITOOBSERVACAO SET OBS='" & Mask(txtObservacao.Text) & "' WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND SEQ=" & nSeq
            cn.Execute Sql, rdExecDirect
            CarregaObs
            For i = 1 To lvObserv.ListItems.Count
                If Val(lvObserv.ListItems(i).Text) = nSeq Then
                    lvObserv.ListItems(i).Selected = True
                    Exit For
                End If
            Next
        End If
    Else
        'SE FOR OBSERVACAO DA PARCELA
        With grdExtrato
            If .Rows = 0 Then
                MsgBox "Não existem débitos.", vbExclamation, "Atenção"
                Exit Sub
            End If
            If .SelectedRow = 0 Then Exit Sub
            nAno = .CellText(.SelectedRow, 1)
            nLanc = Val(Left$(.CellText(.SelectedRow, 2), 3))
            nSeqLanc = Val(.CellText(.SelectedRow, 3))
            nParc = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
            nComp = Val(.CellText(.SelectedRow, 5))
        End With
        
        If bNovoObs Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno
            Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nParc
            Sql = Sql & " AND CODCOMPLEMENTO=" & nComp
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & Val(txtCod.Text) & "," & nAno & ","
'            Sql = Sql & nLanc & "," & nSeqLanc & "," & nParc & "," & nComp & "," & nSeq & ",'" & Mask(txtObservacao.Text) & "','" & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
            Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & Val(txtCod.Text) & "," & nAno & ","
            Sql = Sql & nLanc & "," & nSeqLanc & "," & nParc & "," & nComp & "," & nSeq & ",'" & Mask(txtObservacao.Text) & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
            CarregaObsParcela
            If lvObserv.ListItems.Count > 0 Then
                lvObserv.ListItems(lvObserv.ListItems.Count).Selected = True
            End If
        Else
            nSeq = Val(lvObserv.SelectedItem.Text)
            sData = lvObserv.SelectedItem.SubItems(2)
            Sql = "UPDATE OBSPARCELA SET OBS='" & Mask(txtObservacao.Text) & "' WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno
            Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nParc
            Sql = Sql & " AND CODCOMPLEMENTO=" & nComp & " AND SEQ=" & nSeq
            cn.Execute Sql, rdExecDirect
            CarregaObsParcela
            For i = 1 To lvObserv.ListItems.Count
                If Val(lvObserv.ListItems(i).Text) = nSeq Then
                    lvObserv.ListItems(i).Selected = True
                    Exit For
                End If
            Next
        End If
    End If
    EventosObs True
End If
Liberado

End Sub

Private Sub cmdNovoEF_Click()
Dim z As Long

If lvEFOrigem.ListItems.Count = 0 Then
    MsgBox "Não existem débitos para ser incluidos na execução fiscal.", vbCritical, "Atenção"
    Exit Sub
End If
For x = 1 To lvDoc.ListItems.Count
    lvDoc.ListItems(x).Checked = False
    lvDoc.ListItems(x).SubItems(1) = "0"
    lvDoc.ListItems(x).SubItems(2) = ""
Next
z = SendMessage(lvEFDest.hwnd, LVM_DELETEALLITEMS, 0, 0)
sEventoEF = "Novo"
txtEF.Visible = True
cmbEF.Visible = False
EventosEF False
txtEF.Text = ""
txtDocEF.Text = ""
txtEF.SetFocus
End Sub

Private Sub cmdNovoObs_Click()
bNovoObs = True
EventosObs False
txtObservacao.Text = ""
txtObservacao.SetFocus
End Sub

Private Sub cmdObs_Click()
Dim itmX As ListItem, i As Integer
Dim z As Long

bObs = True
'If NomeDoComputador = "BOJUTSU" Then
    NovaObs
'    Exit Sub
'End If

'z = SendMessage(lvObs.hwnd, LVM_DELETEALLITEMS, 0, 0)
'txtDesc.text = ""
'Ocupado
'i = 0
'Sql = "SELECT LOGPARCELA.ANOEXERCICIO, LOGPARCELA.CODLANCAMENTO, LANCAMENTO.DESCREDUZ, LOGPARCELA.SEQLANCAMENTO, "
'Sql = Sql & "LOGPARCELA.NumParcela , LOGPARCELA.CODCOMPLEMENTO, LOGPARCELA.DATALOG, LOGPARCELA.USUARIO, LOGPARCELA.Texto "
'Sql = Sql & "FROM LOGPARCELA INNER JOIN LANCAMENTO ON LOGPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
'Sql = Sql & "Where LOGPARCELA.CODREDUZIDO = " & Val(txtCod.text)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'
'        Set itmX = lvObs.ListItems.Add(, "O" & Format(i, "0000"), !AnoExercicio)
'        itmX.SubItems(1) = Format(!CodLancamento, "000") & "-" & !DESCREDUZ
'        itmX.SubItems(2) = Format(!SeqLancamento, "00")
'        itmX.SubItems(3) = Format(!NumParcela, "00")
'        itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
'        itmX.SubItems(5) = !USUARIO
'        itmX.SubItems(6) = Format(!DATALOG, "dd/mm/yyyy")
'        itmX.SubItems(7) = SubNull(!Texto)
'        i = i + 1
'       .MoveNext
'    Loop
'   .Close
'End With
'
'Sql = "SELECT DEBITOCANCEL.NUMPROCESSO, DEBITOCANCEL.ANOEXERCICIO, DEBITOCANCEL.CODLANCAMENTO, LANCAMENTO.DESCREDUZ, "
'Sql = Sql & "DEBITOCANCEL.SEQLANCAMENTO, DEBITOCANCEL.NUMPARCELA, DEBITOCANCEL.CODCOMPLEMENTO, DEBITOCANCEL.USUARIO, "
'Sql = Sql & "DEBITOCANCEL.DATACANCEL , DEBITOCANCEL.MOTIVO FROM  DEBITOCANCEL INNER JOIN "
'Sql = Sql & "LANCAMENTO ON DEBITOCANCEL.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
'Sql = Sql & "Where DEBITOCANCEL.CODREDUZIDO = " & Val(txtCod.text)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        Set itmX = lvObs.ListItems.Add(, "O" & Format(i, "0000"), !AnoExercicio)
'        itmX.SubItems(1) = Format(!CodLancamento, "000") & "-" & !DESCREDUZ
'        itmX.SubItems(2) = Format(!SeqLancamento, "00")
'        itmX.SubItems(3) = Format(!NumParcela, "00")
'        itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
'        itmX.SubItems(5) = !USUARIO
'        itmX.SubItems(6) = Format(!DATACANCEL, "dd/mm/yyyy")
'        itmX.SubItems(7) = SubNull(!MOTIVO)
'        i = i + 1
'       .MoveNext
'    Loop
'   .Close
'End With
'
'On Error Resume Next
'Sql = "SELECT CODREDUZIDO,OBSDEBITO FROM DEBITOOBS WHERE CODREDUZIDO=" & Val(txtCod.text)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'     If .RowCount > 0 Then
'         txtObs.text = !OBSDEBITO
'     Else
'        txtObs.text = ""
'     End If
'    .Close
'End With
'
'Liberado
'fim:
frBotao.Enabled = False
'frObs.Visible = True
'frObs.ZOrder 0
frTop.Enabled = False
'grdExtrato.Enabled = False
'nCodObs = 0
End Sub

Private Sub cmdRefresh_Click()
bCarregado = False
txtCod_LostFocus
End Sub

Private Sub cmdReparc_Click()
If Val(txtCod.Text) = 0 Then
    MsgBox "Selecione um contribuinte.", vbExclamation, "Atenção"
    Exit Sub
End If

cmbProc.Clear
Sql = "SELECT DISTINCT NUMPROCESSO FROM ORIGEMREPARC WHERE "
Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      Do Until .EOF
            cmbProc.AddItem !numprocesso
           .MoveNext
      Loop
     .Close
End With
grdOrigem.Rows = 1
grdDestino.Rows = 1
lblCancel.Visible = False
lblResp.Caption = 0
lblFunc.Caption = ""
lblValor.Caption = "0,00"
lblQtde.Caption = "0"
lblDataParc.Caption = ""
    
    'VERIFICA OS REPARCELAMENTOS DA SMAR
    Sql = "SELECT DISTINCT(CODSEQD) From REPARCTMP Where CODREDUZD =" & Val(txtCod.Text) & " Or CODREDUZO = " & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then

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
    Exit Sub
Else
    cmbProc.ListIndex = 0
End If


frBotao.Enabled = False
frReparc.Visible = True
frReparc.ZOrder 0
frTop.Enabled = False
grdExtrato.Enabled = False


End Sub

Private Sub cmdSair_Click()
txtCod.Text = ""
Unload Me
End Sub

Private Sub cmdSairDoc_Click()
frBotao.Enabled = True
frDoc.Visible = False
frTop.Enabled = True
grdExtrato.Enabled = True

End Sub

Private Sub cmdRetornar_Click()
frBotao.Enabled = True
frEFiscal.Visible = False
frTop.Enabled = True
grdExtrato.Enabled = True
End Sub

Private Sub cmdSairObs_Click()
frBotao.Enabled = True
pnlObs.Visible = False
pnlObs.ZOrder 0
frTop.Enabled = True
grdExtrato.Enabled = True

End Sub

Private Sub cmdSairRep_Click()
frBotao.Enabled = True
frReparc.Visible = False
frTop.Enabled = True
grdExtrato.Enabled = True
End Sub

Private Sub Form_Activate()


bCarregado = False
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 122 Then
    frmCadMob.ZOrder 0
End If
End Sub

Private Sub Form_Load()
Dim nIndex As Long
Ocupado
MontaMenu
If NomeDeLogin <> "RENATA" And NomeDeLogin <> "PAULOT" And NomeDeLogin <> "SOLANGE" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "LUIZH" And NomeDeLogin <> "JOSIANE" And _
    NomeDeLogin <> "NOELI" And NomeDeLogin <> "EDUARDO" And NomeDeLogin <> "LEANDRO" And NomeDeLogin <> "ANA" And NomeDeLogin <> "ROSANGELA" And NomeDeLogin <> "PAULA" And _
    NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "PAULO" And NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "RHENO.SOARES" And NomeDeLogin <> "VALQUIRIA.FELIPE" Then
    
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuSuspenso")) = False
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuMultaInfracao")) = False
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuExtratoForum")) = False
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuBuscaDoc")) = False
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuRetidoTomador")) = False
End If

Set xImovel = New clsImovel
Liberado
bChangeStatus = False: bExtrato = False
GridHeader (False)
bExecF = False
frBotao.Enabled = True
frDoc.Visible = False
frReparc.Visible = False
pnlObs.Visible = False
frTop.Enabled = True
grdExtrato.Enabled = True
sRet = RetEventUserForm(Me.Name)
FormHagana
sEventoEF = ""



Me.Height = Val(GetSetting("GTI", "WINDOW", "DEBITO_HEIGHT"))
If Me.Height < 600 Then
   SaveSetting "GTI", "WINDOW", "DEBITO_HEIGHT", 6585
   Me.Height = GetSetting("GTI", "WINDOW", "DEBITO_HEIGHT")
End If
Me.Width = Val(GetSetting("GTI", "WINDOW", "DEBITO_WIDTH"))
If Me.Width < 2000 Then
   SaveSetting "GTI", "WINDOW", "DEBITO_WIDTH", 11655
   Me.Width = GetSetting("GTI", "WINDOW", "DEBITO_WIDTH")
End If

Centraliza Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

SaveSetting "GTI", "WINDOW", "DEBITO_HEIGHT", Me.Height
SaveSetting "GTI", "WINDOW", "DEBITO_WIDTH", Me.Width

Set m_cMenuContrib = Nothing
Set m_cMenuOpcoes = Nothing
Set m_cMenuExtrato = Nothing
Set m_cMenuInterno = Nothing
Set frm = Nothing
Set xImovel = Nothing
End Sub

Private Sub Form_Resize()
If Me.Width < 1500 Or Me.Height < 1700 Then Exit Sub
frBotao.Left = Me.Width - 1400
grdExtrato.Width = Me.Width - 1500
grdExtrato.Height = Me.Height - 1700
frStatus.Top = Me.Height - 950
End Sub

Private Sub frDoc_DragDrop(Source As Control, x As Single, Y As Single)
frDoc.Left = x
frDoc.Top = Y
End Sub

Private Sub grdExtrato_ColumnClick(ByVal lcol As Long)

Dim sTag As String
Dim iSortIndex As Long
      
   With grdExtrato.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lcol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lcol
      End If
   
      ' Determine which sort order to apply:
      sTag = grdExtrato.ColumnTag(lcol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdExtrato.ColumnTag(lcol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdExtrato.ColumnSortType(lcol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdExtrato.Sort
   Screen.MousePointer = vbDefault

End Sub

Private Sub grdExtrato_DblClick(ByVal lrow As Long, ByVal lcol As Long)
Dim nLancamento As Integer, nStatus As Integer, z As Variant, sData As String, sValor As String, nAno As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "ISRAEL" Or NomeDeLogin = "IORIO" Or _
    NomeDeLogin = "SOLANGE" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "LEANDRO" Or NomeDeLogin = "LUIZH" Or _
    NomeDeLogin = "LEONARDODINIZ" Or NomeDeLogin = "ROSANGELA" Or NomeDeLogin = "PAULA" Or NomeDeLogin = "NOELI" Or _
    NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "PAULO" Or NomeDeLogin = "MARILIA" Or IsAtendente Then
    nAno = Val(grdExtrato.CellText(lrow, 1))
    nLancamento = Val(Left$(grdExtrato.CellText(lrow, 2), 3))
    nSeq = Val(grdExtrato.CellText(lrow, 3))
    nParc = Val(grdExtrato.CellText(lrow, 4))
    nCompl = Val(grdExtrato.CellText(lrow, 5))
    nStatus = Val(Left$(grdExtrato.CellText(lrow, 6), 2))
    sData = grdExtrato.CellText(lrow, 7)
    sValor = grdExtrato.CellText(lrow, 10)
    If (nLancamento = 36 Or nLancamento = 41 Or nLancamento = 11 Or nLancamento = 59 Or nLancamento = 52) And nStatus = 3 Then
        If lcol = 7 Then
INIDATA:
            z = InputBox("Digite a nova data de vencimento", "DATA DE VENCIMENTO", sData)
            If z = "" Then Exit Sub
            If Not IsDate(z) Then
                MsgBox "Data inválida", vbCritical, "Atenção"
                GoTo INIDATA
            Else
                If z <> sData Then
                    Sql = "UPDATE DEBITOPARCELA SET DATAVENCIMENTO='" & Format(z, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & nLancamento & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                    cn.Execute Sql, rdExecDirect
                    grdExtrato.CellText(lrow, 7) = z
                End If
            End If
        ElseIf lcol = 10 Then
INIVALOR:
            z = InputBox("Digite o novo valor", "VALOR R$", sValor)
            If z = "" Then Exit Sub
            If Val(z) = 0 Then
                MsgBox "Valor inválido", vbCritical, "Atenção"
                GoTo INIVALOR
            Else
                If z <> sValor Then
                    Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=" & Virg2Ponto(RemovePonto(CStr(z))) & " WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & nLancamento & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO<>3"
                    cn.Execute Sql, rdExecDirect
                    grdExtrato.CellText(lrow, 10) = z
                End If
            End If
        End If
    End If
End If

End Sub

Private Sub grdExtrato_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim x As Integer, Y As Integer, nStatus As Integer, sStatus As String, nOldStatus As Integer
Dim OldRow As Integer, sTexto As String
Dim sNumProc As String, nInicio As Integer
Dim nParcela As Integer, nAno As Integer, nLancamento As Integer, nSequencia As Integer, nComplemento As Integer


With grdExtrato
    If .SelectedRow = 0 Then Exit Sub
    nAno = .CellText(.SelectedRow, 1)
    nLancamento = Val(Left$(.CellText(.SelectedRow, 2), 3))
    nSequencia = Val(.CellText(.SelectedRow, 3))
    nParcela = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
    nComplemento = Val(.CellText(.SelectedRow, 5))
End With


If bExtrato = True Then
    
    If KeyCode = vbKeyReturn Then
        With grdExtrato
            If .CellBackColor(.SelectedRow, 2) = &HC0FFC0 Then
                .CellText(.SelectedRow, 12) = ""
                For x = 1 To .Columns
                    .CellBackColor(.SelectedRow, x) = Branco
                    If x = 6 Then
                       .CellForeColor(.SelectedRow, x) = &HDC&
                    Else
                       .CellForeColor(.SelectedRow, x) = vbBlack
                    End If
                Next
            Else
                .CellText(.SelectedRow, 12) = "S"
                For x = 1 To .Columns
                    .CellBackColor(.SelectedRow, x) = &HC0FFC0
                    .CellForeColor(.SelectedRow, x) = vbBlack
                Next
            End If
        End With
    ElseIf KeyCode = vbKeyF6 Then
        GravaExtrato2 (False)
        bExtrato = False
    End If
    
    Exit Sub
End If



If bChangeStatus = True Then
    If KeyCode = vbKeyReturn Then
        With grdExtrato
            If .CellBackColor(.SelectedRow, 2) = &HC0FFC0 Then
                .CellText(.SelectedRow, 12) = ""
                For x = 1 To .Columns
                    .CellBackColor(.SelectedRow, x) = Branco
                    If x = 6 Then
                       .CellForeColor(.SelectedRow, x) = &HDC&
                    Else
                       .CellForeColor(.SelectedRow, x) = vbBlack
                    End If
                Next
            Else
                .CellText(.SelectedRow, 12) = "S"
                For x = 1 To .Columns
                    .CellBackColor(.SelectedRow, x) = &HC0FFC0
                    .CellForeColor(.SelectedRow, x) = vbBlack
                Next
            End If
        End With
    ElseIf KeyCode = vbKeyF12 Then
        With grdExtrato
            nStatus = Val(InputBox("Digite o valor do novo status.", "Alteração de Status", 0))
            If Val(nStatus) = 0 Then
                MsgBox "Valor Inválido", vbCritical, "Atenção"
                Exit Sub
            End If
            Sql = "SELECT DESCSITUACAO FROM SITUACAOLANCAMENTO WHERE CODSITUACAO=" & nStatus
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount = 0 Then
                MsgBox "Status não existe.", vbCritical, "Atenção"
            Else
                sStatus = Format(nStatus, "00") & "-" & RdoAux!DescSituacao
                If MsgBox("Deseja alterar o status dos lançamentos selecionados para --> " & sStatus, vbQuestion + vbYesNo, "CONFIRMAÇÃO") = vbYes Then
                    For x = 1 To .Rows
                         If .CellBackColor(x, 11) = &HC0FFC0 Then
                            nAno = .CellText(x, 1)
                            nLancamento = Val(Left$(.CellText(x, 2), 3))
                            nSequencia = Val(.CellText(x, 3))
                            nParcela = IIf(.CellText(x, 4) = "Unica", 0, .CellText(x, 4))
                            nComplemento = Val(.CellText(x, 5))
                            nOldStatus = Val(.CellText(x, 6))
                            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " WHERE CODREDUZIDO=" & Val(txtCod.Text) & "AND ANOEXERCICIO=" & nAno & " AND "
                            Sql = Sql & "CODLANCAMENTO=" & nLancamento & " AND SEQLANCAMENTO=" & nSequencia & " AND NUMPARCELA=" & nParcela & " AND CODCOMPLEMENTO=" & nComplemento
                            cn.Execute Sql, rdExecDirect
                            Log Form, Me.Name, Alteração, "Alterado Status de " & CStr(nOldStatus) & " para " & CStr(nStatus) & " no Código:" & txtCod.Text & " Ano:" & nAno & " Lc:" & nLancamento & " Sq:" & nSequencia & " Pc:" & nParcela & " Cp:" & nComplemento
                                                                                            
                           .CellText(x, 6) = sStatus
                         End If
                    Next
                End If
            End If
        End With
    End If
    Exit Sub
End If

OldRow = grdExtrato.SelectedRow
If KeyCode = vbKeyReturn Then
     With grdExtrato
        .Redraw = False
         
         If .SelectedRow > 0 Then
              If Val(Left$(.CellText(.SelectedRow, 6), 2)) = 3 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 19 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 20 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 40 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 38 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 39 Then
                 If .CellBackColor(.SelectedRow, 2) = &HC0FFC0 Then
                     .CellText(.SelectedRow, 12) = ""
                     For x = 1 To .Columns
                         .CellBackColor(.SelectedRow, x) = Branco
                         If x = 6 Then
                            If Val(Left$(.CellText(.SelectedRow, 6), 2)) = 40 Then
                                .CellForeColor(.SelectedRow, x) = Roxo
                            ElseIf Val(Left$(.CellText(.SelectedRow, 6), 2)) = 38 Then
                                .CellBackColor(.SelectedRow, x) = vbRed
                                .CellForeColor(.SelectedRow, x) = vbYellow
                            ElseIf Val(Left$(.CellText(.SelectedRow, 6), 2)) = 39 Then
                                .CellBackColor(.SelectedRow, x) = vbRed
                                .CellForeColor(.SelectedRow, x) = vbWhite
                            Else
                               .CellForeColor(.SelectedRow, x) = &HDC&
                            End If
                         Else
                            .CellForeColor(.SelectedRow, x) = vbBlack
                         End If
                     Next
                 Else
                     .CellText(.SelectedRow, 12) = "S"
                     For x = 1 To .Columns
                         .CellBackColor(.SelectedRow, x) = &HC0FFC0
                         .CellForeColor(.SelectedRow, x) = vbBlack
                     Next
                 End If
              Else
              ' Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 19
                 If Val(Left$(.CellText(.SelectedRow, 6), 2)) = 6 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 5 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 8 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 10 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 5 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 12 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 14 Or Val(Left$(.CellText(.SelectedRow, 6), 2)) = 28 Then
                    Sql = "SELECT debitocancel.*,Usuario.NomeLogin FROM  debitocancel INNER JOIN usuario ON debitocancel.userid = usuario.Id WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & nSequencia & " AND NUMPARCELA=" & nParcela & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & nComplemento
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        If .RowCount > 0 Then
                           sTexto = "Cancelado por: " & SubNull(!NomeLogin) & vbCrLf
                           sTexto = sTexto & "Data: " & Format(!DataCancel, "dd/mm/yyyy") & vbCrLf
                           sTexto = sTexto & "Processo: " & !numprocesso & vbCrLf
                           sTexto = sTexto & "Motivo: " & !MOTIVO
                           MsgBox sTexto, vbInformation, "Dados do Cancelamento."
                        Else
                           MsgBox "Não localizado dados sobre este cancelamento.", vbExclamation, "Atenção"
                        End If
                       .Close
                    End With
                 Else
                    If .CellBackColor(.SelectedRow, 2) = &HC0FFC0 Then
                        .CellText(.SelectedRow, 12) = ""
                        For x = 1 To .Columns
                            .CellBackColor(.SelectedRow, x) = Branco
                            If x = 6 Then
                               .CellForeColor(.SelectedRow, x) = &HDC&
                            Else
                               .CellForeColor(.SelectedRow, x) = vbBlack
                            End If
                        Next
                    End If
                 End If
              End If
          End If
         .Redraw = True
     End With
ElseIf KeyCode = vbKeyF8 Then
     With grdExtrato
         .Redraw = False
          For Y = 1 To .Rows
             If Not bSel Then
                If Val(Left$(.CellText(Y, 6), 2)) = 3 And .CellBackColor(Y, 1) <> &H9FFFC0 Then
                    .CellText(Y, 12) = ""
                     For x = 1 To .Columns
                        .CellBackColor(Y, x) = Branco
                        If x = 6 Then
                          .CellForeColor(Y, x) = &HDC&
                        Else
                          .CellForeColor(Y, x) = vbBlack
                        End If
                     Next
                        lblSel.Caption = FormatNumber(CDbl(lblSel.Caption) - CDbl(.CellText(Y, 11)), 2)
                 End If
             Else
                 If Val(Left$(.CellText(Y, 6), 2)) = 3 And .CellBackColor(Y, 1) <> &H9FFFC0 Then
                    .CellText(Y, 12) = "S"
                    For x = 1 To .Columns
                        .CellBackColor(Y, x) = &HC0FFC0
                        .CellForeColor(Y, x) = vbBlack
                    Next
                    lblSel.Caption = FormatNumber(CDbl(lblSel.Caption) + CDbl(.CellText(Y, 11)), 2)
                 End If
            End If
         Next
        .Redraw = True
     End With
ElseIf KeyCode = vbKeyF3 Then
     With grdExtrato
         .Redraw = False
          For Y = 1 To .Rows
             If Not bSel Then
                If Val(Left$(.CellText(Y, 6), 2)) = 19 And .CellBackColor(Y, 1) <> &H9FFFC0 Then
                    .CellText(Y, 12) = ""
                     For x = 1 To .Columns
                        .CellBackColor(Y, x) = Branco
                        If x = 6 Then
                          .CellForeColor(Y, x) = &HDC&
                        Else
                          .CellForeColor(Y, x) = vbBlack
                        End If
                     Next
                        lblSel.Caption = FormatNumber(CDbl(lblSel.Caption) - CDbl(.CellText(Y, 11)), 2)
                 End If
             Else
                 If Val(Left$(.CellText(Y, 6), 2)) = 19 And .CellBackColor(Y, 1) <> &H9FFFC0 Then
                    .CellText(Y, 12) = "S"
                    For x = 1 To .Columns
                        .CellBackColor(Y, x) = &HC0FFC0
                        .CellForeColor(Y, x) = vbBlack
                    Next
                    lblSel.Caption = FormatNumber(CDbl(lblSel.Caption) + CDbl(.CellText(Y, 11)), 2)
                 End If
            End If
         Next
        .Redraw = True
     End With
ElseIf KeyCode = vbKeyF9 Then
     With grdExtrato
          If .SelectedRow > 0 Then
               Unload frmCnsParcela
               Set frm = frmCnsParcela
               frm.nParcela = nParcela
               frm.nAno = nAno
               frm.nLancamento = nLancamento
               frm.nSequencia = nSequencia
               frm.nComplemento = nComplemento
               If .CellBackColor(.SelectedRow, .SelectedCol) = 0 Then
                  frm.nCodRed = Val(txtCod.Text)
               Else
                  nInicio = InStr(1, .CellText(.SelectedRow, 2), " (", vbBinaryCompare)
                  sNumProc = Mid(.CellText(.SelectedRow, 2), nInicio + 2, Len(.CellText(.SelectedRow, 2)) - Val(nInicio) - 2)
                  Sql = "SELECT CODIGORESP FROM PROCESSOREPARC WHERE  NUMPROCESSO='" & sNumProc & "'"
                  Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                  With RdoAux
                       If .RowCount > 0 Then
                          frm.nCodRed = !CODIGORESP
                          frm.nResponsavel = 1
                       Else
                          frm.nCodRed = Val(txtCod.Text)
                          frm.nResponsavel = 0
                       End If
                      .Close
                 End With
               End If
               frm.show vbModal
          End If
     End With
ElseIf KeyCode = vbKeyF5 Then
    With grdExtrato
         If Val(.CellText(.SelectedRow, 3)) < 100 Or (Val(Left$(.CellText(.SelectedRow, 2), 3)) <> 11 And Val(.CellText(.SelectedRow, 1)) > 2000) Then
            MsgBox "Apenas Sequencias >=100, ou Taxas Diversas até 2.001 podem ser canceladas por duplicidade", vbExclamation, "Atenção"
            Exit Sub
         End If
    
         If MsgBox("Cancelar este lancamento por duplicidade?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
           .CellText(.SelectedRow, 6) = "12-CANCELADO POR DUPLICIDADE"
           Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=12 WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND "
           Sql = Sql & "CODLANCAMENTO=" & nLancamento & " AND SEQLANCAMENTO=" & nSequencia & " AND NUMPARCELA=" & nParcela & " AND "
           Sql = Sql & "CODCOMPLEMENTO=" & nComplemento
           cn.Execute Sql, rdExecDirect
         End If
    End With
End If
bSel = Not bSel
lblSel.Caption = "0,00"
With grdExtrato
    If .Rows = 0 Then Exit Sub
    .Redraw = False
    For x = 1 To .Rows
         If .CellBackColor(x, 11) = &HC0FFC0 Then
             lblSel.Caption = FormatNumber(CDbl(lblSel.Caption) + CDbl(.CellText(x, 11)), 2)
         End If
    Next
   .SelectedRow = OldRow
   .SetFocus
   .Redraw = True
End With

End Sub

Private Sub grdExtrato_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single, bDoDefault As Boolean)

If Button = vbRightButton Then
    nIndex = m_cMenuOpcoes.IndexForKey("mnuAjuiza")
    If m_cMenuOpcoes.Enabled(nIndex) = True Then
        nIndex = m_cMenuInterno.IndexForKey("mnuCancelAjuiza")
        m_cMenuInterno.Enabled(nIndex) = True
    Else
        nIndex = m_cMenuInterno.IndexForKey("mnuCancelAjuiza")
        m_cMenuInterno.Enabled(nIndex) = False
    End If

    nIndex = m_cMenuOpcoes.IndexForKey("mnuEditaParcela")
    If m_cMenuOpcoes.Enabled(nIndex) = True Then
        nIndex = m_cMenuInterno.IndexForKey("mnuEditParcela")
        m_cMenuInterno.Enabled(nIndex) = True
    Else
        nIndex = m_cMenuInterno.IndexForKey("mnuEditParcela")
        m_cMenuInterno.Enabled(nIndex) = False
    End If

    lIndex = m_cMenuInterno.ShowPopupMenu(x, Y + 800, x, Y, Me.ScaleWidth - x, Y, False)
End If

End Sub

Private Sub lvDoc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If sEventoEF = "" Then
    lvDoc.ListItems(Item.Index).Checked = Not lvDoc.ListItems(Item.Index).Checked
Else
    If lvDoc.ListItems(Item.Index).Checked Then
        lvDoc.ListItems(Item.Index).SubItems(1) = 1
    Else
        lvDoc.ListItems(Item.Index).SubItems(1) = 0
    End If
End If
End Sub

Private Sub lvObserv_Click()
Dim sObs As String
If lvObserv.ListItems.Count = 0 Then
    txtObservacao.Text = ""
    Exit Sub
End If
sObs = lvObserv.SelectedItem.SubItems(3)
txtObservacao.Text = sObs
End Sub

Private Sub lvObserv_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sObs As String
If lvObserv.ListItems.Count = 0 Then
    txtObservacao.Text = ""
    Exit Sub
End If
sObs = lvObserv.SelectedItem.SubItems(3)
txtObservacao.Text = sObs
End Sub

Private Sub m_cMenuContrib_Click(ItemNumber As Long)

Select Case m_cMenuContrib.ItemKey(ItemNumber)
    Case "mnuMob"
        sFormMob = "DI2"
        frmCnsMob.show
        frmCnsMob.ZOrder 0
    Case "mnuImob"
        sForm = "DI"
        frmCnsImovel.show
        frmCnsImovel.ZOrder 0
    Case "mnuOutros"
        Set frm = frmCnsCidadao
        frm.sForm = "frmDebitoImob"
        frm.show
        frm.ZOrder 0
End Select

End Sub

Private Sub m_cMenuExtrato_Click(ItemNumber As Long)
Select Case m_cMenuExtrato.ItemKey(ItemNumber)
    Case "mnuExtratoCompleto"
        If Val(txtCod.Text) = 0 Then Exit Sub
        If cGetInputState() <> 0 Then DoEvents
        nExtrato = 1
        GravaExtrato2 (False)
    Case "mnuExtratoFiltro"
        If Val(txtCod.Text) = 0 Then Exit Sub
        MsgBox "Selecione as parcelas que deseja incluir no extrato e pressione F6 para concluir.", vbInformation, "Filtro de Extrato"
        bExtrato = True
        nExtrato = 2
        grdExtrato.SetFocus
End Select

End Sub

Private Sub m_cMenuInterno_Click(ItemNumber As Long)
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nComp As Integer, nTipo As Integer
Dim nCodReduz As Long, nParcela As Integer, nLancamento As Integer, nSequencia As Integer, nComplemento As Integer
Dim z As Variant

Select Case m_cMenuInterno.ItemKey(ItemNumber)
    Case "mnuCancelAjuiza"
        With grdExtrato
            If .SelectedRow = 0 Then
                MsgBox "Selecione o Lançamento.", vbExclamation, "Atenção"
                Exit Sub
            End If
            If .CellText(.SelectedRow, 9) = "N" Then
                MsgBox "Lançamento não ajuizado.", vbExclamation, "Atenção"
                Exit Sub
            Else
                If MsgBox("Cancelar este ajuizamento?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                
                    nAno = .CellText(.SelectedRow, 1)
                    nLanc = Left$(.CellText(.SelectedRow, 2), 3)
                    nSeq = .CellText(.SelectedRow, 3)
                    nParc = IIf(.CellText(.SelectedRow, 4) = "Unica", "00", .CellText(.SelectedRow, 4))
                    nComp = .CellText(.SelectedRow, 5)
                    
                    Sql = "UPDATE DEBITOPARCELA SET DATAAJUIZA=NULL WHERE CODREDUZIDO=" & Val(txtCod.Text)
                    Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
                    Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nComp
                    cn.Execute Sql, rdExecDirect
                    .CellText(.SelectedRow, 9) = "N"
                End If
            End If
        End With
    Case "mnuCancelNotifISS"
        With grdExtrato
            If .SelectedRow = 0 Then
                MsgBox "Selecione o Lançamento.", vbExclamation, "Atenção"
                Exit Sub
            End If
            If .CellText(.SelectedRow, 13) = "N" Then
                MsgBox "Lançamento não notificado.", vbExclamation, "Atenção"
                Exit Sub
            Else
                If MsgBox("Cancelar esta notificação?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                
                    nAno = .CellText(.SelectedRow, 1)
                    nLanc = Left$(.CellText(.SelectedRow, 2), 3)
                    nSeq = .CellText(.SelectedRow, 3)
                    nParc = IIf(.CellText(.SelectedRow, 4) = "Unica", "00", .CellText(.SelectedRow, 4))
                    nComp = .CellText(.SelectedRow, 5)
                    
                    Sql = "UPDATE DEBITOPARCELA SET NOTIFICADO=NULL WHERE CODREDUZIDO=" & Val(txtCod.Text)
                    Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
                    Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nComp
                    cn.Execute Sql, rdExecDirect
                    .CellText(.SelectedRow, 13) = "N"
                End If
            End If
        End With
    Case "mnuEditParcela"
        If grdExtrato.SelectedRow > 0 Then
           If IsAtendente Then
              nLanc = Val(Left(grdExtrato.CellText(grdExtrato.SelectedRow, 2), 3))
              Select Case nLanc
                  Case 11, 36, 48
                     frmParcela.show vbModal
                  Case Else
                     MsgBox "Você não esta autorizado a alterar este lançamento.", vbCritical, "Atenção"
              End Select
            Else
                If bEDI Then
                    frmParcela.show vbModal
                End If
            End If
        End If
    Case "mnuReativar"
        If Val(Left$(grdExtrato.CellText(grdExtrato.SelectedRow, 6), 2)) <> 19 Then
            MsgBox "Apenas lancamentos suspensos podem ser reativados.", vbCritical, "Atenção"
        Else
            With grdExtrato
                nCodReduz = Val(txtCod.Text)
                For x = 1 To .Rows
                    If .CellText(x, 12) = "S" Then
                        nAno = .CellText(x, 1)
                        nLancamento = Val(Left$(.CellText(x, 2), 3))
                        nSequencia = .CellText(x, 3)
                        nParcela = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
                        nComplemento = .CellText(x, 5)
                           
                        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC =" & 3 & " WHERE CODREDUZIDO=" & nCodReduz
                        Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLancamento & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & nSequencia & " AND NUMPARCELA=" & nParcela & " AND "
                        Sql = Sql & "CODCOMPLEMENTO=" & nComplemento
                        cn.Execute Sql, rdExecDirect
                        grdExtrato.CellText(x, 6) = "03-NÃO PAGO"
                        grdExtrato.CellForeColor(x, 6) = vbRed
                    End If
                Next
            End With
        End If
    Case "mnuReativarJ"
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) = 20 And .CellText(x, 12) = "S" Then
                    nCodReduz = Val(txtCod.Text)
                    nAno = .CellText(x, 1)
                    nLancamento = Val(Left$(.CellText(x, 2), 3))
                    nSequencia = Val(.CellText(x, 3))
                    nParcela = IIf(.CellText(x, 4) = "Unica", 0, .CellText(x, 4))
                    nComplemento = Val(.CellText(x, 5))
                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC =" & 3 & " WHERE CODREDUZIDO=" & nCodReduz
                    Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLancamento & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & nSequencia & " AND NUMPARCELA=" & nParcela & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & nComplemento
                    cn.Execute Sql, rdExecDirect
                    grdExtrato.CellText(x, 6) = "03-NÃO PAGO"
                    grdExtrato.CellForeColor(x, 6) = vbRed
                End If
            Next
        End With
    Case "mnuVerObs"
        With grdExtrato
            If .SelectedRow = 0 Then Exit Sub
            nAno = .CellText(.SelectedRow, 1)
            nLancamento = Val(Left$(.CellText(.SelectedRow, 2), 3))
            nSequencia = Val(.CellText(.SelectedRow, 3))
            nParcela = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
            nComplemento = Val(.CellText(.SelectedRow, 5))
        End With
        
        Sql = "SELECT * FROM DEBITOCANCEL LEFT OUTER JOIN USUARIO ON DEBITOCANCEL.USERID=USUARIO.ID "
        Sql = Sql & "Where CODREDUZIDO = " & Val(txtCod.Text) & " And "
        Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLancamento & " AND "
        Sql = Sql & "SEQLANCAMENTO=" & nSequencia & " AND NUMPARCELA=" & nParcela & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nComplemento
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               sTexto = "Cancelado/Suspenso por: " & SubNull(!NomeLogin) & vbCrLf
               sTexto = sTexto & "Data: " & Format(!DataCancel, "dd/mm/yyyy") & vbCrLf
               sTexto = sTexto & "Processo: " & !numprocesso & vbCrLf
               sTexto = sTexto & "Motivo: " & !MOTIVO
               MsgBox sTexto, vbInformation, "Dados do Cancelamento."
            Else
               MsgBox "Não localizado dados sobre este cancelamento.", vbExclamation, "Atenção"
            End If
           .Close
        End With
    Case "mnuVinculoMI"
        On Error Resume Next
        With grdExtrato
            If Val(Left(.CellText(.SelectedRow, 2), 3)) = 69 Or Val(Left(.CellText(.SelectedRow, 2), 3)) = 49 Or Right(.CellText(.SelectedRow, 2), 4) = "(MI)" Then
                If Val(Left(.CellText(.SelectedRow, 2), 3)) = 69 Then
                    nTipo = 2 'VIEW MULTA
                Else
                    nTipo = 3 'VIEW LANC
                End If
            Else
               MsgBox "Lançamento selecionado não é multa de infração, nem esta vinculada a ela.", vbExclamation, "Atenção"
               Exit Sub
            End If
        End With
        
        frmMulta.nTipo = nTipo
        frmMulta.show vbModal
    Case "mnu2viaAuto"
        With grdExtrato
            If Val(Left(.CellText(.SelectedRow, 2), 3)) = 16 Then
                frmReport.ShowReport2 "MULTAINF2", frmMdi.hwnd, Me.hwnd
            Else
               MsgBox "Lançamento selecionado não é multa de infração.", vbExclamation, "Atenção"
               Exit Sub
            End If
        End With
        
    Case "mnuAgrupa"
        m_cMenuInterno.Checked(ItemNumber) = Not m_cMenuInterno.Checked(ItemNumber)
        grdExtrato.AllowGrouping = m_cMenuInterno.Checked(ItemNumber)
        If grdExtrato.AllowGrouping = False Then GridHeader (True)
End Select

End Sub

Private Sub m_cMenuOpcoes_Click(ItemNumber As Long)
Dim Achou As Boolean, bSim As Boolean, bNao As Boolean, nConta As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nComp As Integer
Dim vData As Variant, dData As Date, x As Integer, z As Variant
Dim Sql As String, RdoAux As rdoResultset, nAnoproc As Integer, nNumproc As Long, RdoAux2 As rdoResultset, nStatus As Integer, RdoAux3 As rdoResultset

Select Case m_cMenuOpcoes.ItemKey(ItemNumber)
    Case "mnuCancelDebito"
        
        
             For x = 1 To grdExtrato.Rows
                If grdExtrato.CellText(x, 12) = "S" Then
                    If Val(Left$(grdExtrato.CellText(x, 2), 3)) = 3 Or Val(Left$(grdExtrato.CellText(x, 2), 3)) = 5 Or Val(Left$(grdExtrato.CellText(x, 2), 3)) = 65 Or Val(Left$(grdExtrato.CellText(x, 2), 3)) = 14 Then
                        If NomeDeLogin <> "LUIZH" And NomeDeLogin <> "ROSE" And NomeDeLogin <> "RITA" And NomeDeLogin <> "RENATA" And NomeDeLogin <> "EDUARDO" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "NOELI" And NomeDeLogin <> "ANA" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "RHENO.SOARES" And NomeDeLogin <> "VALQUIRIAFELIPE." Then
                           MsgBox "Você não possui permissão para cancelar/suspender lançamentos de ISS.", vbExclamation, "Atenção"
                           Exit Sub
                        End If
                      
                    End If
                    If Val(Left$(grdExtrato.CellText(x, 2), 3)) <> 14 And Val(Left$(grdExtrato.CellText(x, 2), 3)) <> 13 And Val(Left$(grdExtrato.CellText(x, 2), 3)) <> 6 And Val(Left$(grdExtrato.CellText(x, 2), 3)) <> 2 And Val(Left$(grdExtrato.CellText(x, 2), 3)) And (NomeDeLogin = "RITA" Or NomeDeLogin = "DANIELAR") Then
                         MsgBox "Você não possui permissão para cancelar/suspender este tipo de lançamento.", vbExclamation, "Atenção"
                         Exit Sub
                    End If
                
                End If
            Next
        
        Select Case NomeDeLogin
            Case "EDUARDO", "RITA", "GLEISE", "RENATA", "SCHWARTZ", "ROSE", "JOSEANE", "NOELI", "ANA", "JOSIANE", "LEANDRO", "LUIZH", "MARILIA", "PAULA", "DANIELAR", "SOLANGE", "PAULOT", "RHENO.SOARES", "VALQUIRIA.FELIPE"
                Achou = False
                With grdExtrato
                    For x = 1 To .Rows
                        If .CellText(x, 12) = "S" Then
                            Achou = True
                            Exit For
                        End If
                    Next
                End With
                
                If Not Achou Then
                    MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
                Else
                    frmCancelDebito.show vbModal
                End If
            Case Else
                MsgBox "Você não esta autorizado a cancelar débitos no GTI." & vbCrLf & "Somente os responsáveis pelo setor de tributação e do Sistema Prático de Atendimento tem esta permissão.", vbCritical, "Atenção"
        End Select
    Case "mnuAnexaDoc"
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser anexados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
            frmAnexaDoc.show vbModal
        End If
    Case "mnuDA"
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 8) = "S" And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não inscritos na divida ativa podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If Val(.CellText(x, 1)) > Year(Now) And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos maiores que o ano atual podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
            Ocupado
            frmDivAtivaManual.show vbModal
        End If
    Case "mnuAjuiza"
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 8) = "N" And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos inscritos na divida ativa podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            
            bSim = False
            bNao = False
            For x = 1 To .Rows
                If .CellText(x, 9) = "S" And .CellText(x, 12) = "S" Then
                   bSim = True
                   lblAjuiza.Caption = "S"
                   Exit For
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 9) = "N" And .CellText(x, 12) = "S" Then
                   bNao = True
                   lblAjuiza.Caption = "N"
                   Exit For
                End If
            Next
            
            If bSim And bNao Then
                MsgBox "Você não pode selecionar parcelas ajuizadas e não ajuizadas junto.", vbExclamation, "atenção"
                Exit Sub
            End If
           
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
            Ocupado
            frmAjuizamento.show 1
        End If
    Case "mnuObs"
        bObs = False
        NovaObsParcela
        frBotao.Enabled = False
        frTop.Enabled = False
    Case "mnuEditaParcela"
        If grdExtrato.SelectedRow > 0 Then
           If IsAtendente Then
              nLanc = Val(Left(grdExtrato.CellText(grdExtrato.SelectedRow, 2), 3))
              Select Case nLanc
                  Case 11, 36, 48
                     frmParcela.show vbModal
                  Case Else
                     MsgBox "Você não esta autorizado a alterar este lançamento.", vbCritical, "Atenção"
              End Select
            Else
                If bEDI Then
                    frmParcela.show vbModal
                End If
            End If
       End If
    Case "mnuSmar"
'        If grdExtrato.Rows > 1 Then
'            frmReparcOld.show vbModal
'        End If
    Case "mnuSemMov"
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If (Val(Left(.CellText(x, 2), 3)) <> 5 Or Val(.CellText(x, 10)) > 0) And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos de ISS Variavel com Valor Zero podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
            frmCancelDebito.show
            frmCancelDebito.cmbTipo.ListIndex = 2
            frmCancelDebito.cmbTipo.Enabled = False
        End If
    Case "mnuCompensa"
        With grdExtrato
            If .Rows = 0 Then Exit Sub
            If .SelectedRow = 0 Then Exit Sub
            nAno = .CellText(.SelectedRow, 1)
            nLanc = Val(Left$(.CellText(.SelectedRow, 2), 3))
            nSeq = Val(.CellText(.SelectedRow, 3))
            nParc = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
            nComp = Val(.CellText(.SelectedRow, 5))
            nConta = 0
        
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                   nConta = nConta + 1
                End If
            Next
            
        End With
        
        If nConta <> 1 Then
            MsgBox "Selecione apenas uma parcela.", vbExclamation, "atenção"
        Else
            If MsgBox("Voce deseja compensar o débito selecionado ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=6 WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
                Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nComp
                cn.Execute Sql, rdExecDirect
                grdExtrato.CellText(grdExtrato.SelectedRow, 6) = "06-COMPENSADO"
                grdExtrato.CellForeColor(grdExtrato.SelectedRow, 6) = vbBlue
            End If
        End If
    Case "mnuSuspenso"
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
        
            For x = 1 To grdExtrato.Rows
                If grdExtrato.CellText(x, 12) = "S" Then
                    If Val(Left$(grdExtrato.CellText(x, 2), 3)) = 3 Or Val(Left$(grdExtrato.CellText(x, 2), 3)) = 5 Or Val(Left$(grdExtrato.CellText(x, 2), 3)) = 65 Or Val(Left$(grdExtrato.CellText(x, 2), 3)) = 14 Then
                        If NomeDeLogin <> "LUIZH" And NomeDeLogin <> "EDUARDO" Then
                           MsgBox "Você não possui permissão para Suspender lançamentos de ISS.", vbExclamation, "Atenção"
                           Exit Sub
                        End If
                       
                    End If
                End If
            Next
        
        
        
            frmCancelDebito.show
            frmCancelDebito.cmbTipo.ListIndex = 3
            frmCancelDebito.cmbTipo.Enabled = False
        End If
    Case "mnuRetidoTomador"
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
            frmCancelDebito.show
            frmCancelDebito.cmbTipo.ListIndex = 5
            frmCancelDebito.cmbTipo.Enabled = False
        End If
    Case "mnuDAMH"
        If NomeDeLogin <> "ROSE" And NomeDeLogin <> "JOSEANE" And Not IsAtendente And NomeDeLogin <> "CARMELINO" And NomeDeLogin <> "RHENO.SOARES" And NomeDeLogin <> "VALQUIRIA.FELIPE" And NomeDeLogin <> "WHICTOR.HOMEM" Then
            MsgBox "Acesso negado!", vbCritical, "ERRO"
            Exit Sub
        End If
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If CDbl(.CellText(x, 10)) = 0 And .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Achou Then
            MsgBox "Não é possível emitir DAM para débitos com valor zerado.", vbCritical, "Atenção"
            Exit Sub
        End If
        
        Achou = False
        With grdExtrato
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        End With
        
        If Not Achou Then
            MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
        Else
Inicio:
            vData = InputBox("Digite a Data de Vencimento da DAM.", "Atenção", Format(Now, "dd/mm/yyyy"))
            If vData <> "" Then
               If Len(vData) <> 10 Then
                  MsgBox "Data inválida.", vbCritical, "Atenção"
                  GoTo Inicio
               Else
                  If Not IsDate(vData) Then
                     MsgBox "Data inválida.", vbCritical, "Atenção"
                     GoTo Inicio
                  Else
                     dData = CDate(vData)
                     If dData < Format(Now, "dd/mm/yyyy") Then
                        MsgBox "Data de vencimento não pode ser retroativa.", vbCritical, "Atenção"
                        GoTo Inicio
                     Else
                        lblDataVencto.Caption = vData
                       frmDAM.Honorarios = True
                       frmDAM.VencimentoDAM = vData
                       frmDAM.CodigoDAM = txtCod.Text
                       frmDAM.show vbModal
                     End If
                  End If
               End If
            End If
        End If
    Case "mnuReativaParc"
        nCodReduz = Val(txtCod.Text)
        With grdExtrato
            x = .SelectedRow
            If x = 0 Then
                MsgBox "Selecione uma parcela do parcelamento a ser cancelado.", vbCritical, "Atenção"
                Exit Sub
            End If
            If Val(Left(.CellText(x, 2), 3)) = 20 Then
                nAno = Val(.CellText(x, 1))
                nLanc = Val(Left$(.CellText(x, 2), 3))
                nSeq = Val(.CellText(x, 3))
                nParc = Val(.CellText(x, 4))
                nCompl = Val(.CellText(x, 5))
                Sql = "SELECT NUMPROCESSO FROM DESTINOREPARC WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND "
                Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND NUMSEQUENCIA=" & nSeq & " AND CODCOMPLEMENTO=" & nCompl
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        sNumProc = !numprocesso
                        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
                        nAnoproc = Val(Right$(sNumProc, 4))
                        .Close
                        If MsgBox("Deseja REATIVAR o Parcelamento: " & CStr(nNumproc) & "-" & RetornaDVProcesso(nNumproc) & "/" & CStr(nAnoproc), vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                           '*** REATIVAREMOS PRIMEIRO O DESTINO DO PARCELAMENTO***
                           'AS PARCELAS QUE ESTIVEREM CANCELADAS(5) SE TORNAM NÃO PAGAS(3)
                           'MAS DEVEMOS VERIFICAR ANTES SE NÃO HOUVE ERRO NO CANCELAMENTO E CASO
                           'ACUSAR PAGAMENTO EM ALGUMA PARCELA COLOCAREMOS O STATUS DE PAGO(2)
                            
                            Sql = "SELECT * From destinoreparc WHERE numprocesso = '" & sNumProc & "'"
                            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux
                                Do Until .EOF
                                    nAno = !AnoExercicio
                                    nLanc = !CodLancamento
                                    nSeq = !numsequencia
                                    nParc = !NumParcela
                                    nCompl = !CODCOMPLEMENTO
                                    Sql = "SELECT * FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
                                    Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl & " AND RESTITUIDO IS NULL"
                                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                    With RdoAux2
                                        If .RowCount > 0 Then
                                            nStatus = 2
                                        Else
                                            nStatus = 3
                                        End If
                                       .Close
                                    End With
                                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
                                    Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                                    cn.Execute Sql, rdExecDirect
                                   .MoveNext
                                Loop
                               .Close
                            End With
                            
                            '*** REATIVAREMOS AGORA A ORIGEM DO PARCELAMENTO***
                            'AS PARCELAS QUE ESTIVEREM COMPENSADAS(6) SE TORNAM REPARCELADAS(4)
                            'AS PARCELAS QUE ESTIVEREM NÃO PAGAS(3) SE TORNAM REPARCELADAS TAMBEM(4)
                            
                            Sql = "SELECT * From origemreparc WHERE numprocesso = '" & sNumProc & "'"
                            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux
                                Do Until .EOF
                                    nAno = !AnoExercicio
                                    nLanc = !CodLancamento
                                    nSeq = !numsequencia
                                    nParc = !NumParcela
                                    nCompl = !CODCOMPLEMENTO
                                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=4 WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
                                    Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                                    cn.Execute Sql, rdExecDirect
                                   .MoveNext
                                Loop
                               .Close
                            End With
                            
                            '*** REATIVAREMOS AGORA O PARCELAMENTO EM SI***
                            Sql = "UPDATE PROCESSOREPARC SET CANCELADO=0,DATACANCEL=NULL,FUNCIONARIOCANCEL=NULL WHERE NUMPROCESSO='" & sNumProc & "'"
                            cn.Execute Sql, rdExecDirect
                            
                            Log Form, Me.Caption, Alteração, "Reativado processo nº '" & CStr(nNumproc) & "-" & RetornaDVProcesso(nNumproc) & "/" & CStr(nAnoproc) & "'"
                            
                             'GRAVA NA TABELA ACORDOSTATUS
                             If frmMdi.frTeste.Visible = False Then
                                ConectaIntegrativa
                                Sql = "insert acordostatus(idacordo,anoacordo,dtocorrencia,ocorrencia,dtgeracao) values("
                                Sql = Sql & nNumproc & "," & nAnoproc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & "PARCELAMENTO EM DIA" & "','" & Format(Now, "mm/dd/yyyy") & "')"
                                cnInt.Execute Sql, rdExecDirect
                                cnInt.Close
                            End If
                            MsgBox "O Parcelamento: " & CStr(nNumproc) & "-" & RetornaDVProcesso(nNumproc) & "/" & CStr(nAnoproc) & " foi reativado." & vbCrLf & "Atualize os dados para ver as alterações." & vbCrLf & vbCrLf & "IMPORTANTE: O Complemento do parcelamento deve ser excluido MANUALMENTE!", vbInformation, "Informação"
                        Else
                            MsgBox "Reativação cancelada.", vbExclamation, "Atenção"
                        End If
                    Else
                        MsgBox "Este parcelamento não pode ser reativado automaticamente.", vbCritical, "Atenção"
                    End If
                End With
                
            Else
                MsgBox "Parcela selecionada não é de um parcelamento.", vbCritical, "Atenção"
            End If
        End With
    Case "mnuMultaInfracao"
        With grdExtrato
            For x = 1 To .Rows
                If Val(Left$(.CellText(x, 6), 2)) <> 3 And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos não pagos podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If ((Val(Left(.CellText(x, 2), 3)) <> 5 And Val(Left(.CellText(x, 2), 3)) <> 49) Or Val(.CellText(x, 10)) = 0) And .CellText(x, 12) = "S" Then
                   MsgBox "Apenas lançamentos de ISS Variavel com Valor maior que Zero podem ser selecionados.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If Right(.CellText(x, 2), 4) = "(MI)" And .CellText(x, 12) = "S" Then
                   MsgBox "Não selecione ISS Variavel que já possua multa de infração.", vbExclamation, "Atenção"
                   Exit Sub
                End If
            Next
            For x = 1 To .Rows
                If .CellText(x, 12) = "S" Then
                    Achou = True
                    Exit For
                End If
            Next
        
            If Not Achou Then
                MsgBox "Selecione ao menos uma parcela.", vbExclamation, "atenção"
            Else
                frmMulta.nTipo = 1 'lanc normal
                frmMulta.show vbModal
            End If
            
        End With
    Case "mnuExtratoForum"
        If Val(txtCod.Text) = 0 Then Exit Sub
        GravaExtrato2 (True)
    Case "mnuSerasa"
        RetiraSerasa
    Case "mnuChangeStatus"
        MsgBox "Selecione as parcelas que deseja alterar e pressione F12 para concluir.", vbInformation, "Alteração de Status"
        bChangeStatus = True
    Case "mnuBuscaDoc"
        z = InputBox("Digite o número do documento", "Informação requerida")
        If Val(z) > 0 Then
            Sql = "SELECT * FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND NUMDOCUMENTO=" & Val(z)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                    MsgBox "O Documento se refere a parcela: " & vbCrLf & vbCrLf & "Ano: " & !AnoExercicio & vbCrLf & "Lancamento: " & !CodLancamento & vbCrLf & "Sequencia: " & !SeqLancamento & vbCrLf & _
                    "Parcela: " & !NumParcela & vbCrLf & "Complemento: " & !CODCOMPLEMENTO & vbCrLf & "Valor R$: " & FormatNumber(RdoAux2!soma, 2), vbInformation, "Resultado"
                    For x = 1 To grdExtrato.Rows
                        If grdExtrato.CellText(x, 1) = !AnoExercicio And Val(Left(grdExtrato.CellText(x, 2), 3)) = !CodLancamento And Val(grdExtrato.CellText(x, 3)) = !SeqLancamento And _
                        Val(grdExtrato.CellText(x, 4)) = !NumParcela And Val(grdExtrato.CellText(x, 5)) = !CODCOMPLEMENTO Then
                           grdExtrato.SelectedRow = x
                           Exit For
                        End If
                    Next
                    
                Else
                    MsgBox "Documento não encontrado para esta inscrição.", vbCritical, "Atenção"
                End If
               .Close
            End With
        End If
End Select

End Sub

Private Sub txtCod_Change()
bCarregado = False
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
Dim nCodImovel As Long
If Val(txtCod.Text) = 0 Then Exit Sub
nCodImovel = Val(txtCod.Text)
Ocupado
ZeraFiltro
pnlInativo.Visible = False
bFilterLoad = False
Sql = "SELECT * FROM vwfullimovel2 WHERE CODREDUZIDO=" & txtCod.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
'        lblNum.Caption = "Nº de Inscrição Cadastral"
        lblRS.Caption = "Proprietário"
        pnlInativo.Visible = !Inativo
        'CarregaImovel nCodImovel
        txtProp.Text = SubNull(!nomecidadao)
        lblProp.Caption = SubNull(!nomecidadao)
        lblRua.Caption = SubNull(!Logradouro) & ", " & SubNull(!Li_Num)
        DoEvents
        If txtCod.Text = "" Then Exit Sub
        CarregaDebito txtCod.Text
    Else
       .Close
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & txtCod.Text
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
            '   lblNum.Caption = "Nº de Inscrição Estadual"
               lblNumInsc.Caption = SubNull(!INSCESTADUAL)
               lblRS.Caption = "Raz.Social"
               lblProp.Caption = !RazaoSocial
               txtProp.Text = !RazaoSocial
               lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " nº " & SubNull(!Numero)
               CarregaDebito txtCod.Text
            Else
              .Close
               Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                   If .RowCount > 0 Then
                      If Val(txtCod.Text) < 500000 Then
                        '  lblNum.Caption = ""
                        lblNumInsc.Caption = ""
                        lblProp.Caption = !nomecidadao
                        txtProp.Text = ""
                        lblRua.Caption = ""
                        grdExtrato.Clear
                         MsgBox "Não existe débito para este código.", vbExclamation, "Atenção"
                      Else
                       '  lblNum.Caption = ""
                        lblNumInsc.Caption = ""
                        lblRS.Caption = "Proprietário"
                        lblProp.Caption = !nomecidadao
                        txtProp.Text = !nomecidadao
                        lblRua.Caption = ""
                        CarregaDebito txtCod.Text
                      End If
                   Else
                     MsgBox "Código não cadastrado.", vbCritical, "Atenção"
                     grdExtrato.Clear
                     grdExtrato.SetFocus
                   End If
                  .Close
               End With
            End If
        End With
    End If
End With
Liberado
'grdExtrato.SetFocus
End Sub

Private Sub CarregaImovel(nCodigoImovel As Long)

Ocupado
With xImovel
    .CarregaImovel nCodigoImovel
    lblDV.Caption = Left$(lblDV.Caption, Len(lblDV.Caption) - 1) & xImovel.DV
    If .Inativo = True Then
        pnlInativo.Visible = True
        
        
        cmdCancelReparc.Enabled = False
        'If Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) <> "ROSE" Then
        If NomeDeLogin <> "ROSE" And NomeDeLogin <> "SCHWARTZ" Then
            cmdAj.Enabled = False
           cmdDAM.Enabled = False
'           cmdGravar.Enabled = False
        Else
            cmdAj.Enabled = True
        End If
    Else
        pnlInativo.Visible = False
        cmdAj.Enabled = True
        cmdCancelReparc.Enabled = True
        If InStr(1, sRet, Format(evDAM, "000"), vbBinaryCompare) > 0 Then
           cmdDAM.Enabled = True
        Else
           cmdDAM.Enabled = False
        End If
'        cmdGravar.Enabled = True
    End If
    
    If .CodigoImovel > 0 Then
        lblNumInsc.Caption = .Inscricao
        lblRua.Caption = .EnderecoCompleto
    Else
        grdExtrato.Clear
        lblDebito.Caption = "0,00"
        lblVencer.Caption = "0,00"
'        lblDebitoUnica.Caption = "0,00"
        'lblVencerUnica.Caption = "0,00"
        lblNumInsc.Caption = ""
        lblRua.Caption = ""
        lblProp.Caption = ""
        txtProp.Text = ""
        MsgBox "Imóvel não cadastrado.", vbCritical, "Atenção"
        GoTo fim
    End If
    lblProp.Caption = .NomePropPrincipal
    txtProp.Text = .NomePropPrincipal
End With
grdExtrato.Redraw = False
CarregaDebito nCodigoImovel
grdExtrato.Redraw = True
fim:
Liberado

End Sub


Private Sub ZeraFiltro()
'Unload frmFiltroDebito
'FiltroE = 0
'FiltroL = 0
'FiltroS = 99
'FiltroA = "X"
'FiltroD = "X"
'FiltroLP = ""
If cmdFilter.value = True Then
    cmdFilter.value = False
    pnlFilter.Visible = False
    grdExtrato.Height = 4875
End If
bFilterLoad = False
bExecF = False
cmbAno1.Clear
cmbAno2.Clear
cmbLanc.Clear
cmbSit.Clear
cmbSeq.Clear
cmbDA.ListIndex = 0
cmbAj.ListIndex = 0
bExecF = True
End Sub

Private Sub FormHagana()
Dim nIndex As Integer

evEDI = 3
evDAM = 5
evCND = 6
evDAT = 7
evAJU = 8
evADO = 9
evSMA = 10
evSMOV = 12
evCOM = 13
evRea = 18
evReaJ = 20
evRP = 22
evEF = 23
evSer = 24

bDam = False: bCND = False: bDAT = False: bAJU = False: bADO = False: bEDI = False: bSMA = False: bSMOV = False: bCOM = False: bRea = False: bReaJ = False: bRP = False: bEF = False: bSer = False

If InStr(1, sRet, Format(evDAM, "000"), vbBinaryCompare) > 0 Then bDam = True
If InStr(1, sRet, Format(evCND, "000"), vbBinaryCompare) > 0 Then bCND = True
If InStr(1, sRet, Format(evDAT, "000"), vbBinaryCompare) > 0 Then bDAT = True
If InStr(1, sRet, Format(evAJU, "000"), vbBinaryCompare) > 0 Then bAJU = True
If InStr(1, sRet, Format(evADO, "000"), vbBinaryCompare) > 0 Then bADO = True
If InStr(1, sRet, Format(evEDI, "000"), vbBinaryCompare) > 0 Then bEDI = True
If InStr(1, sRet, Format(evSMA, "000"), vbBinaryCompare) > 0 Then bSMA = True
If InStr(1, sRet, Format(evSMOV, "000"), vbBinaryCompare) > 0 Then bSMOV = True
If InStr(1, sRet, Format(evCOM, "000"), vbBinaryCompare) > 0 Then bCOM = True
If InStr(1, sRet, Format(evRea, "000"), vbBinaryCompare) > 0 Then bRea = True
If InStr(1, sRet, Format(evReaJ, "000"), vbBinaryCompare) > 0 Then bReaJ = True
If InStr(1, sRet, Format(evRP, "000"), vbBinaryCompare) > 0 Then bRP = True
If InStr(1, sRet, Format(evEF, "000"), vbBinaryCompare) > 0 Then bEF = True
If InStr(1, sRet, Format(evSer, "000"), vbBinaryCompare) > 0 Then bSer = True

'On Error Resume Next
If Not bEF Then
    cmdNovoEF.Enabled = False
    cmdAlterarEF.Enabled = False
    cmdGravarEF.Enabled = False
    cmdExcluirEF.Enabled = False
    cmdAlterarEF.Enabled = False
    cmdAllEfo.Enabled = False
    cmdDelEfo.Enabled = False
End If
If Not bDam Then cmdDAM.Enabled = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuCancelDebito")
If Not bCND Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuDA")
If Not bDAT Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuAjuiza")
If Not bAJU Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuAnexaDoc")
If Not bADO Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuEditaParcela")
If Not bEDI Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuSmar")
If Not bSMA Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuSemMov")
If Not bSMOV Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuCompensa")
If Not bCOM Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuInterno.IndexForKey("mnuReativar")
If Not bReaJ Then m_cMenuInterno.Enabled(nIndex) = False
nIndex = m_cMenuInterno.IndexForKey("mnuReativarJ")
If Not bReaJ Then m_cMenuInterno.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuReativaParc")
If Not bRP Then m_cMenuOpcoes.Enabled(nIndex) = False
nIndex = m_cMenuOpcoes.IndexForKey("mnuSerasa")
If Not bSer Then m_cMenuOpcoes.Enabled(nIndex) = False

m_cMenuInterno.Enabled(m_cMenuInterno.IndexForKey("mnuCancelNotifISS")) = frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuNotificaISS"))

If NomeDeLogin <> "RENATA" And NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "ROSE" And NomeDeLogin <> "RITA" And NomeDeLogin <> "JOSEANE" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "SOLANGE" And NomeDeLogin <> "ANA" Then
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuChangeStatus")) = False
Else
    m_cMenuInterno.Enabled(m_cMenuInterno.IndexForKey("mnuReativar")) = True
    m_cMenuOpcoes.Enabled(m_cMenuOpcoes.IndexForKey("mnuChangeStatus")) = True
End If

End Sub

Public Sub CarregaDebito(nCodImovel As Long)
Dim aDebito() As Debito
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoS As rdoResultset, RdoP As rdoResultset
Dim nValorDebito As Double, Achou As Boolean, x As Integer, sExecFiscal As String
Dim nSomaDebito As Double, nEval As Integer, nValorCorrecao As Double, bFind As Boolean, k As Integer
Dim nSomaVencer As Double, nSomaDebitoUnica As Double, nSomaVencerUnica As Double
Dim sDescReduz As String, nValorAtualizado As Double, nSomaValorTributo As Double
Dim bAjuiza As Boolean, bDA As Boolean, qd As New rdoQuery, bIsentoMJ As Boolean

lblDebito.Caption = "0,00"
lblVencer.Caption = "0,00"
lblSel.Caption = "0,00"
dDataAtualiza = CDate(Right$(frmMdi.Sbar.Panels(6).Text, 10))

'dDataAtualiza = CDate("10/25/2010")
m_cMenuInterno.Checked(m_cMenuInterno.IndexForKey("mnuAgrupa")) = False
grdExtrato.AllowGrouping = False
GridHeader (True)
'CorrigeUnica
bChangeStatus = False
Achou = False
ReDim aDebito(0)

For x = 0 To Forms.Count - 1
    If Forms(x).Name = "frmFiltroDebito" Then
         Achou = True
         Exit For
    End If
Next

If Not Achou And bCarregado Then Exit Sub
modLg "Consulta de débito código: " & nCodImovel & " - " & txtProp.Text
Achou = False: bAjuiza = False: bDA = False
grdExtrato.Redraw = False
grdExtrato.Clear
grdExtrato.Redraw = True
grdExtrato.Redraw = False
lblDebito.Caption = "0,00"
lblVencer.Caption = "0,00"
nSomaDebito = 0
nSomaVencer = 0
bSel = True


'CARREGA O EXTRATO
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = nCodImovel
qd(1) = nCodImovel
If cmbAno1.ListCount > 0 Then
    qd(2) = cmbAno1.Text: qd(3) = cmbAno2.Text
Else
    qd(2) = 1950: qd(3) = 2050
End If

'If chkTodosAnos.Value = vbUnchecked Then
'    qd(2) = Year(Now) - 5
'End If

If cmbLanc.ListIndex > 0 Then
    qd(4) = Val(Left(cmbLanc.Text, 3)): qd(5) = Val(Left(cmbLanc.Text, 3)) 'LANCAMENTO
Else
    qd(4) = 0: qd(5) = 99
End If
If cmbSeq.ListIndex > 0 Then
    qd(6) = Val(cmbSeq.Text): qd(7) = Val(cmbSeq.Text) 'SEQUENCIA
Else
    qd(6) = 0: qd(7) = 9999
End If
qd(8) = 0: qd(9) = 999
qd(10) = 0: qd(11) = 99

If cmbSit.ListIndex > 0 Then
    qd(12) = Val(Left(cmbSit.Text, 2)): qd(13) = Val(Left(cmbSit.Text, 2)) 'STATUSLANC
Else
    qd(12) = 0: qd(13) = 99
End If
'qd(14) = Format("11/01/2010", "mm/dd/yyyy")
qd(14) = Format(dDataAtualiza, "mm/dd/yyyy")
qd(15) = NomeDeLogin
Set RdoAux = qd.OpenResultset(rdOpenKeyset)

With RdoAux
    If RdoAux.RowCount > 0 Then
        ReDim Preserve aDebito(UBound(aDebito) + 1)
        nEval = UBound(aDebito)
        Do Until .EOF
          '  If aDebito(nEval).nAno = 2016 And aDebito(nEval).nParc = 1 And aDebito(nEval).nLanc = 13 Then MsgBox "teste"
            If chkTodosAnos.value = vbUnchecked Then
                'If !AnoExercicio < Year(Now) - 5 And !statuslanc <> 3 And !statuslanc <> 19 And !statuslanc <> 20 And !statuslanc <> 25 Then
                If !AnoExercicio < Year(Now) - 5 And !statuslanc <> 3 And !statuslanc <> 19 And !statuslanc <> 25 And !statuslanc <> 20 And !statuslanc <> 40 Then
                    GoTo proximo
                End If
            End If
        
'            If !SeqLancamento = 9 And !NumParcela = 21 Then MsgBox "tste"
            bJuros = False: bMulta = False: bIsentoMJ = False
            If Not IsNull(!NumDocumento) Then
                Sql = "SELECT NUMDOCUMENTO,ISENTOMJ FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & RdoAux!NumDocumento
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux2.RowCount > 0 Then
                    If Val(SubNull(RdoAux2!isentomj)) > 0 Then
                        bIsentoMJ = True
                    End If
                End If
                RdoAux2.Close
            End If
       '     If !AnoExercicio = 2017 Then MsgBox "teste"
            If cmbAj.ListIndex = 1 Then
                If IsNull(!dataajuiza) Then GoTo proximo
            End If
            If cmbAj.ListIndex = 2 Then
                If Not IsNull(!dataajuiza) Then GoTo proximo
            End If
            If cmbDA.ListIndex = 1 Then
                If IsNull(!datainscricao) Then GoTo proximo
            End If
            If cmbDA.ListIndex = 2 Then
                If Not IsNull(!datainscricao) Then GoTo proximo
            End If

            If chkShowAll.value = vbUnchecked Then
                If !statuslanc = 5 Then GoTo proximo
                If !NumParcela = 0 And !statuslanc = 5 Then GoTo proximo
                If !NumParcela > 0 And !statuslanc = 1 And chkUnica.value = 0 Then GoTo proximo
                If !NumParcela = 0 And !statuslanc = 3 And !CODCOMPLEMENTO = 0 And DateDiff("d", !DataVencimento, Now) > 0 Then GoTo proximo
                If !AnoExercicio = 2003 And !CodLancamento = 1 And (!statuslanc <> 2 And !statuslanc <> 1) Then GoTo proximo
            End If
            
            nEval = UBound(aDebito)
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aDebito(UBound(aDebito) + 1)
                nEval = UBound(aDebito)
                aDebito(nEval).nAno = !AnoExercicio
                aDebito(nEval).nLanc = !CodLancamento
                If !CodLancamento = 20 Or !CodLancamento = 8 Then
                   If Not IsNull(!numprocesso) Then
                      If Val(Right$(!numprocesso, 4)) >= 2006 Then
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1) & "-" & RetornaDVProcesso(Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1)) & "/" & Right$(!numprocesso, 4) & ")"
                      Else
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & !numprocesso & ")"
                      End If
                   Else
                      aDebito(nEval).sLanc = !DESCLANCAMENTO
                   End If
                Else
                   aDebito(nEval).sLanc = !DESCLANCAMENTO
                End If
                aDebito(nEval).nSeq = !SeqLancamento
                aDebito(nEval).nParc = !NumParcela
                aDebito(nEval).nCompl = !CODCOMPLEMENTO
                aDebito(nEval).nSituacao = !statuslanc
                aDebito(nEval).sSituacao = !Situacao
                aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aDebito(nEval).nCodTributo = !CodTributo
                aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
                    If Not IsNull(!valorpagoreal) Then
                        aDebito(nEval).nValorAtual = FormatNumber(!valorpagoreal, 2)
                    Else
                        aDebito(nEval).nValorAtual = FormatNumber(0, 2)
                    End If
                Else
                    If bIsentoMJ Then
                        aDebito(nEval).nValorAtual = FormatNumber(!ValorTributo + !ValorCorrecao, 2)
                    Else
                        aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                    End If
                End If
                If IsNull(!notificado) Then
                    aDebito(nEval).sNotificado = "N"
                Else
                    aDebito(nEval).sNotificado = IIf(!notificado = True, "S", "N")
                End If
                
                sExecFiscal = ""
                If Not IsNull(!ProcessoCNJ) Then
                    sExecFiscal = !ProcessoCNJ
                Else
                    If Not IsNull(!anoexecfiscal) Then
                        sExecFiscal = Format(!numexecfiscal, "00000") & "/" & !anoexecfiscal
                    End If
                End If
                aDebito(nEval).sExFiscal = sExecFiscal
                aDebito(nEval).nProt_certidao = Val(SubNull(!prot_certidao))
                If IsNull(!prot_dtremessa) Then
                    aDebito(nEval).nProt_dtremessa = CDate("01/01/1900")
                Else
                    aDebito(nEval).nProt_dtremessa = Format(!prot_dtremessa, "dd/mm/yyyy")
                End If
            Else
                bFind = False
                For k = 1 To UBound(aDebito)
                    If aDebito(k).nAno = !AnoExercicio And aDebito(k).nLanc = !CodLancamento And _
                       aDebito(k).nSeq = !SeqLancamento And aDebito(k).nParc = !NumParcela And _
                       aDebito(k).nCompl = !CODCOMPLEMENTO And aDebito(k).nCodTributo = !CodTributo Then
                       bFind = True
                       Exit For
                    End If
                Next
                
                If Not bFind Then
                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                    If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
'                         aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTributo, 2)
                    Else
                        If bIsentoMJ Then
                            aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTributo + !ValorCorrecao, 2)
                        Else
                            aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTotal, 2)
                        End If
                    End If
                End If
            End If
proximo:
            .MoveNext
        Loop
      End If
   .Close
End With

'************************************************
'CÁLCULO DE REPARCELAMENTOS MULTIPLOS DA IDIOTA DA SMAR
GoTo CORRECAO
Dim nNumproc As Long, nAno As Integer, nNumSeq As Integer, nNovaSeq As Integer

Sql = "SELECT * From REPARC2TMP ORDER BY NUMPROC, DATAVENCTO, NUMSEQ"
Set RdoP = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoP
    nNovaSeq = 0
    Do Until .EOF
'        If !NUMPROC = 7893 Then msgbox "aqui"
        If !NumProc = nNumproc And Year(!DataVencto) = nAno Then
            Sql = "UPDATE REPARC2TMP SET SEQPROC=" & nNovaSeq & " WHERE "
            Sql = Sql & "NUMPROC=" & nNumproc & "AND NUMSEQ=" & !NumSeq & " AND CODREDUZ=" & !CodReduz
            cn.Execute Sql, rdExecDirect
            
            Sql = "UPDATE REPARCTMP SET SEQPROC=" & nNovaSeq & " WHERE "
            Sql = Sql & "NUMPROC=" & nNumproc & " AND CODSEQD=" & !NumSeq & " AND CODREDUZD=" & !CodReduz
            cn.Execute Sql, rdExecDirect
            
            nNovaSeq = nNovaSeq + 1
        Else
            nNovaSeq = 0
            nNumproc = !NumProc
            nAno = Year(!DataVencto)
            nNumSeq = !NumSeq
            
            Sql = "UPDATE REPARC2TMP SET SEQPROC=" & nNovaSeq & " WHERE "
            Sql = Sql & "NUMPROC=" & nNumproc & "AND NUMSEQ=" & !NumSeq & " AND CODREDUZ=" & !CodReduz
            cn.Execute Sql, rdExecDirect
            
            Sql = "UPDATE REPARCTMP SET SEQPROC=" & nNovaSeq & " WHERE "
            Sql = Sql & "NUMPROC=" & nNumproc & " AND ANOEXERCD=" & !ANOEXERC & "  AND CODSEQD=" & !NumSeq & " AND CODREDUZD=" & !CodReduz
            cn.Execute Sql, rdExecDirect
            nNovaSeq = nNovaSeq + 1
        End If
       .MoveNext
    Loop
End With

CORRECAO:

GoTo CORRECAO2

Sql = "SELECT * FROM REPARC2TMP ORDER BY NUMPROC, SEQPROC"
Set RdoP = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoP
    Do Until .EOF
        Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,SEQPROC,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
        Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CODIGORESP) VALUES('"
        Sql = Sql & "S-" & Format(!NumProc, "0000000") & "/" & Year(!DataVencto) & "'," & !SEQPROC & ",'" & Format(!DataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',"
        Sql = Sql & !PARCELAS & "," & 0 & "," & 0 & ","
        Sql = Sql & IIf(!TEMMULTA, 1, 0) & "," & IIf(!TEMJUROS, 1, 0) & ","
        Sql = Sql & !CodReduz & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA ORIGEM
        Sql = "SELECT * From REPARCTMP Where CODREDUZD=" & !CodReduz & " AND CODSEQD=" & !NumSeq
        Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoS
            Do Until .EOF
                Sql = "INSERT ORIGEMREPARC (NUMPROCESSO,SEQPROC,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO) VALUES('" & "S-" & Format(RdoP!NumProc, "0000000") & "/" & Year(RdoP!DataVencto) & "'," & RdoP!SEQPROC & "," & !CODREDUZO & ","
                Sql = Sql & !ANOEXERCO & "," & !CODLANCO & "," & !CODSEQO & "," & !NUMPARCO & ","
                Sql = Sql & !CODCOMPLO & ")"
                cn.Execute Sql, rdExecDirect
              .MoveNext
           Loop
           .Close
        End With
   
      'GRAVA DESTINO
       Sql = "SELECT DISTINCT CODREDUZO From REPARCTMP Where CODREDUZD =" & !CodReduz
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       Do Until RdoAux2.EOF
            Sql = "UPDATE DEBITOPARCELA SET NUMPROCESSO='" & "S-" & Format(RdoP!NumProc, "0000000") & "/" & Year(RdoP!DataVencto) & "',SEQPROC=" & RdoP!SEQPROC & " WHERE CODREDUZIDO=" & RdoAux2!CODREDUZO & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & !NumSeq
            cn.Execute Sql, rdExecDirect
           RdoAux2.MoveNext
       Loop
       
        Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CodReduz & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & !NumSeq
        Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoS
           Do Until .EOF
               Sql = "INSERT DESTINOREPARC (NUMPROCESSO,SEQPROC,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
               Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO) VALUES('" & "S-" & Format(RdoP!NumProc, "0000000") & "/" & Year(RdoP!DataVencto) & "'," & RdoP!SEQPROC & "," & !CODREDUZIDO & ","
               Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & ","
               Sql = Sql & !CODCOMPLEMENTO & ")"
               cn.Execute Sql, rdExecDirect
              .MoveNext
           Loop
       End With
      .MoveNext
   Loop
   .Close
End With

'************************************************
CORRECAO2:

GoTo CORRECAO3

Dim aTrib() As Debito, nCodTributo As Integer, bAchou As Boolean

Sql = "SELECT NUMPROCESSO,SEQPROC,CODIGORESP FROM PROCESSOREPARC WHERE NUMPROCESSO='S-0015658/2002' ORDER BY NUMPROCESSO,SEQPROC"
Set RdoP = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoP
    Do Until .EOF
            ReDim aTrib(0)
            Sql = "SELECT DEBITOPARCELA.*, DEBITOTRIBUTO.CODTRIBUTO, DEBITOTRIBUTO.VALORTRIBUTO "
            Sql = Sql & "FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO AND "
            Sql = Sql & "DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO AND "
            Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND DEBITOPARCELA.NUMPARCELA = DEBITOTRIBUTO.NUMPARCELA AND "
            Sql = Sql & "DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO WHERE NUMPROCESSO ='" & !numprocesso & "' AND SEQPROC=" & !SEQPROC
            Sql = Sql & " AND CODTRIBUTO<>3 "
            Sql = Sql & "ORDER BY DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.NUMPARCELA,CODTRIBUTO"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                    nCodTributo = !CodTributo
                    If nCodTributo = 2 Then nCodTributo = 1
                    bAchou = False
                    For x = 1 To UBound(aTrib)
                        If aTrib(x).nAno = !AnoExercicio And aTrib(x).nSeq = !SeqLancamento And aTrib(x).nParc = !NumParcela And aTrib(x).nCompl = !CODCOMPLEMENTO And aTrib(x).nCodTributo = nCodTributo Then
                            bAchou = True: Exit For
                        End If
                    Next
                    If bAchou Then
                        aTrib(x).nValorTributo = aTrib(x).nValorTributo + !ValorTributo
                    Else
                        ReDim Preserve aTrib(UBound(aTrib) + 1)
                        aTrib(UBound(aTrib)).nAno = !AnoExercicio
                        aTrib(UBound(aTrib)).nSeq = !SeqLancamento
                        aTrib(UBound(aTrib)).nParc = !NumParcela
                        aTrib(UBound(aTrib)).nCompl = !CODCOMPLEMENTO
                        aTrib(UBound(aTrib)).nCodTributo = nCodTributo
                        aTrib(UBound(aTrib)).nValorTributo = !ValorTributo
                    End If
                   .MoveNext
                Loop
               .Close
            End With
        .MoveNext
    Loop
   .Close
End With

CORRECAO3:
'************************************************
For x = 2 To UBound(aDebito)
    With aDebito(x)
        grdExtrato.AddRow
        grdExtrato.CellDetails grdExtrato.Rows, 1, .nAno, DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 2, Format(.nLanc, "000") & " - " & .sLanc
        grdExtrato.CellDetails grdExtrato.Rows, 3, Format(.nSeq, "00"), DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 4, Format(.nParc, "00"), DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 5, Format(.nCompl, "00"), DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 6, Format(.nSituacao, "00") & "-" & .sSituacao
        grdExtrato.CellDetails grdExtrato.Rows, 7, .sVencto, DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 8, .sDA, DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 9, .sAj, DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 10, FormatNumber(.nValorTributo, 2), DT_RIGHT
        If .nSituacao = 1 Or .nSituacao = 2 Then
            grdExtrato.CellDetails grdExtrato.Rows, 11, FormatNumber(.nValorTributo, 2), DT_RIGHT
        Else
            grdExtrato.CellDetails grdExtrato.Rows, 11, FormatNumber(.nValorAtual, 2), DT_RIGHT
        End If
        grdExtrato.CellDetails grdExtrato.Rows, 13, .sNotificado, DT_CENTER
        grdExtrato.CellDetails grdExtrato.Rows, 14, .sExFiscal, DT_LEFT
        If .nProt_certidao > 0 Then
            grdExtrato.CellDetails grdExtrato.Rows, 15, Format(.nProt_certidao, "000000000"), DT_LEFT
        End If
        If .nProt_dtremessa <> "01/01/1900" Then
            grdExtrato.CellDetails grdExtrato.Rows, 16, .nProt_dtremessa, DT_CENTER
        End If
            
        If DateDiff("d", CDate(.sVencto), Now) > 0 And .nSituacao = 3 Then
           nSomaVencer = nSomaVencer + .nValorAtual
        End If
        If .nSituacao = 3 And .nParc = 0 Then
        Else
            If .nSituacao = 3 And .nParc > 0 Then
                nSomaDebito = nSomaDebito + .nValorAtual
            End If
        End If
        With grdExtrato
           If Left$(.CellText(.Rows, 6), 2) = "01" Or Left$(.CellText(.Rows, 6), 2) = "02" Or Left$(.CellText(.Rows, 6), 2) = "04" Then
              .CellForeColor(.Rows, 6) = &H3F810C
            ElseIf Left$(.CellText(.Rows, 6), 2) = "03" Then
               .CellForeColor(.Rows, 6) = &HDC&
            ElseIf Left$(.CellText(.Rows, 6), 2) = "25" Then
               .CellForeColor(.Rows, 6) = Roxo
            ElseIf Left$(.CellText(.Rows, 6), 2) = "40" Then
               .CellForeColor(.Rows, 6) = Roxo
            ElseIf Left$(.CellText(.Rows, 6), 2) = "41" Then
               .CellForeColor(.Rows, 6) = Roxo
            ElseIf Left$(.CellText(.Rows, 6), 2) = "38" Then
               .CellBackColor(.Rows, 6) = vbRed
               .CellForeColor(.Rows, 6) = vbYellow
            ElseIf Left$(.CellText(.Rows, 6), 2) = "39" Then
               .CellBackColor(.Rows, 6) = vbRed
               .CellForeColor(.Rows, 6) = vbWhite
            Else
               .CellForeColor(.Rows, 6) = vbBlue
            End If
        End With
    End With
Next
grdExtrato.Redraw = True

'SERASA
imgSerasa.Visible = InSerasa(nCodImovel)

With grdExtrato
        If .Rows = 0 Then
             lblDebito.Caption = "0,00"
             lblVencer.Caption = "0,00"
             Liberado
             MsgBox "Não existem Débitos.", vbInformation, "Atenção"
             .Clear
        End If
End With




GoTo fim
'MAALE TASHLUMIM LO KAIAMIM
Sql = "SELECT DISTINCT(NUMPROCESSO) FROM VWCNSREPARCELAMENTOO WHERE "
Sql = Sql & "CODREDUZIDO=" & nCodImovel & " AND CODIGORESP<>" & nCodImovel
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "SELECT * FROM VWCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & !numprocesso & "'"
        If FiltroE > 0 Then
           Sql = Sql & " AND ANOEXERCICIO=" & FiltroE
        End If
        If FiltroL > 0 Then
           Sql = Sql & " AND CODLANCAMENTO=" & FiltroL
        End If
        If FiltroS > 0 Then
           Sql = Sql & " AND STATUSLANC=" & FiltroS
        End If
        If FiltroD = "S" Then
           Sql = Sql & " AND DATAINSCRICAO IS NOT NULL"
        ElseIf FiltroD = "N" Then
           Sql = Sql & " AND DATAINSCRICAO IS NULL"
        End If
        If FiltroA = "S" Then
           Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
        ElseIf FiltroA = "N" Then
           Sql = Sql & " AND DATAAJUIZA IS NULL"
        End If
        Sql = Sql & " ORDER BY ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,NUMPARCELA "
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                 If !CodLancamento = 20 And !statuslanc = 5 Then GoTo proximo3
                 Sql = "SELECT DESCREDUZ FROM LANCAMENTO WHERE CODLANCAMENTO=" & !CodLancamento
                 Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 sDescReduz = RdoS!descreduz & " (" & !numprocesso & ")"
                 Sql = "SELECT DESCSITUACAO FROM SITUACAOLANCAMENTO WHERE CODSITUACAO=" & !statuslanc
                 Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 sDescSituacao = RdoS!DescSituacao
                 
                'BUSCA VALOR LANÇADO
                Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
                Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
                Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
                Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
                Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
                Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        nSomaValorTributo = !ValorTributo
                        nValorAtualizado = !ValorTributo
                        nSomaValorTributo = nSomaValorTributo + CDbl(CalculaJuros(!ValorTributo + CalculaCorrecao(!ValorTributo, !DataVencimento), !DataVencimento))
                        nSomaValorTributo = nSomaValorTributo + CDbl(CalculaMulta(!ValorTributo + CalculaCorrecao(!ValorTributo, !DataVencimento), !DataVencimento))
                        nSomaValorTributo = nSomaValorTributo + CDbl(CalculaCorrecao(!ValorTributo, !DataVencimento))
                    Else
                        nSomaValorTributo = 0
                    End If
                   .Close
                End With
                grdExtrato.AddRow
                grdExtrato.CellDetails grdExtrato.Rows, 1, !AnoExercicio, DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 2, Format(!CodLancamento, "000") & " - " & sDescReduz
                grdExtrato.CellDetails grdExtrato.Rows, 3, Format(!numsequencia, "00"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 4, Format(!NumParcela, "00"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 5, Format(!CODCOMPLEMENTO, "00"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 6, Format(!statuslanc, "00") & "-" & sDescSituacao
                grdExtrato.CellDetails grdExtrato.Rows, 7, Format(!DataVencimento, "dd/mm/yyyy"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 8, IIf(IsNull(!datainscricao), "N", "S"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 9, IIf(IsNull(!dataajuiza), "N", "S"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 10, FormatNumber(nValorAtualizado, 2), DT_RIGHT
                grdExtrato.CellDetails grdExtrato.Rows, 11, FormatNumber(nSomaValorTributo, 2), DT_RIGHT
                
                'TZEVA OT
                 With grdExtrato
                    If Left$(.CellText(.Rows, 6), 2) = "01" Or Left$(.CellText(.Rows, 6), 2) = "02" Or Left$(.CellText(.Rows, 6), 2) = "04" Then
                       .CellForeColor(.Rows, 6) = &H3F810C
                    ElseIf Left$(.CellText(.Rows, 6), 2) = "03" Then
                       .CellForeColor(.Rows, 6) = &HDC&
                    Else
                       .CellForeColor(.Rows, 6) = vbBlue
                    End If
                    For x = 1 To .Columns
                       .CellBackColor(.Rows, x) = &H9FFFC0
                    Next
                 End With
                'SCHUM
                If !statuslanc = 3 Then
                   If Val(Left$(grdExtrato.CellText(grdExtrato.Rows, 4), 2)) > 0 Then
                      If CDate(Now) < !DataVencimento Then
                         nSomaVencer = nSomaVencer + nSomaValorTributo
                      End If
                   Else
                      nSomaDebitoUnica = nSomaDebitoUnica + nSomaValorTributo
                      If CDate(Now) < !DataVencimento Then
                         nSomaVencerUnica = nSomaVencerUnica + nValorDebito
                      End If
                   End If
                End If
                
                .MoveNext
             Loop
        End With
proximo3:
        .MoveNext
     Loop
End With

'REPARCELAMENTO SMAR
Sql = "SELECT DISTINCT * FROM REPARCTMP WHERE  CODREDUZO=" & Val(txtCod.Text) & " AND CODREDUZD<>" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZD & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & !CODSEQD
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                If !CodLancamento = 20 And !statuslanc = 5 Then GoTo PROXIMO2
                Sql = "SELECT DESCREDUZ FROM LANCAMENTO WHERE CODLANCAMENTO=" & !CodLancamento
                Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                sDescReduz = RdoS!descreduz & " (SMAR: " & !CODREDUZIDO & ")"
                Sql = "SELECT DESCSITUACAO FROM SITUACAOLANCAMENTO WHERE CODSITUACAO=" & !statuslanc
                Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                sDescSituacao = RdoS!DescSituacao
                'BUSCA VALOR LANÇADO
                Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
                Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
                Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
                Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
                Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & RdoAux2!CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & RdoAux2!SeqLancamento
                Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
                Set RdoAuxS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAuxS
                    nSomaValorTributo = 0
                    If .RowCount > 0 Then
                        nValorAtualizado = !ValorTributo
                    Else
                        nValorAtualizado = 0
                    End If
                   .Close
                End With
                
                grdExtrato.AddRow
                grdExtrato.CellDetails grdExtrato.Rows, 1, !AnoExercicio, DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 2, Format(!CodLancamento, "000") & " - " & sDescReduz
                grdExtrato.CellDetails grdExtrato.Rows, 3, Format(!SeqLancamento, "00"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 4, Format(!NumParcela, "00"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 5, Format(!CODCOMPLEMENTO, "00"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 6, Format(!statuslanc, "00") & "-" & sDescSituacao
                grdExtrato.CellDetails grdExtrato.Rows, 7, Format(!DataVencimento, "dd/mm/yyyy"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 8, IIf(IsNull(!datainscricao), "N", "S"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 9, IIf(IsNull(!dataajuiza), "N", "S"), DT_CENTER
                grdExtrato.CellDetails grdExtrato.Rows, 10, FormatNumber(nValorAtualizado, 2), DT_RIGHT
                grdExtrato.CellDetails grdExtrato.Rows, 11, FormatNumber(nSomaValorTributo, 2), DT_RIGHT
                
                'TZEVA OT
                 With grdExtrato
                    If Left$(.CellText(.Rows, 6), 2) = "01" Or Left$(.CellText(.Rows, 6), 2) = "02" Or Left$(.CellText(.Rows, 6), 2) = "04" Then
                       .CellForeColor(.Rows, 6) = &H3F810C
                    ElseIf Left$(.CellText(.Rows, 6), 2) = "03" Then
                       .CellForeColor(.Rows, 6) = &HDC&
                    Else
                       .CellForeColor(.Rows, 6) = vbBlue
                    End If
                    For x = 1 To .Columns
                       .CellBackColor(.Rows, x) = &H9FFFC0
                    Next
                 End With
                'SCHUM
                If !statuslanc = 3 Then
                   If Val(Left$(grdExtrato.CellText(grdExtrato.Rows, 4), 2)) > 0 Then
                      If CDate(Now) < !DataVencimento Then
                         nSomaVencer = nSomaVencer + nSomaValorTributo
                      End If
                   Else
                      nSomaDebitoUnica = nSomaDebitoUnica + nSomaValorTributo
                      If CDate(Now) < !DataVencimento Then
                         nSomaVencerUnica = nSomaVencerUnica + nValorDebito
                      End If
                   End If
                End If
PROXIMO2:
               .MoveNext
            Loop
        End With
    End If
End With

fim:
'MEMALE ET A SCHUM
lblDebito.Caption = FormatNumber(nSomaDebito, 2)
lblVencer.Caption = FormatNumber(nSomaVencer, 2)
bCarregado = True
grdExtrato.Redraw = True

'cmdFilter.SetFocus
End Sub



Private Sub GridHeader(bIgnore As Boolean)

With grdExtrato
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    If Not bIgnore Then
        .AddColumn "kAno", "Ano", ecgHdrTextALignCentre, , 38
        .AddColumn "kLanc", "Lançamento", ecgHdrTextALignLeft, , 180
        .AddColumn "kSeq", "Sq", ecgHdrTextALignCentre, , 25
        .AddColumn "kParc", "Pc", ecgHdrTextALignCentre, , 25
        .AddColumn "kComp", "Cp", ecgHdrTextALignCentre, , 25
        .AddColumn "kSit", "Situação", ecgHdrTextALignLeft, , 105
        .AddColumn "kVenc", "Vencto", ecgHdrTextALignCentre, , 67
        .AddColumn "kDa", "D", ecgHdrTextALignLeft, , 17
        .AddColumn "kAj", "A", ecgHdrTextALignLeft, , 17
        .AddColumn "kVL", "Vl.Lanc", ecgHdrTextALignRight, , 70
        .AddColumn "kVA", "Vl.Atual", ecgHdrTextALignRight, , 70
        .AddColumn "kSL", "Sl", ecgHdrTextALignRight, , 25, False
        .AddColumn "kNT", "N", ecgHdrTextALignCentre, , 17
        .AddColumn "kEF", "Ex.Fiscal", ecgHdrTextALignLeft, , 150
        .AddColumn "kPNum", "Certidão", ecgHdrTextALignCentre, , 67
        .AddColumn "kPDtR", "Dt.Remes.", ecgHdrTextALignLeft, , 67
    End If
End With

End Sub

Private Sub NovaObs()
Dim itmX As ListItem, z As Long

frBotao.Enabled = False
pnlObs.Visible = True
pnlObs.ZOrder 0
frTop.Enabled = False
grdExtrato.Enabled = False

CarregaObs
If lvObserv.ListItems.Count > 0 Then
    lvObserv.ListItems(1).Selected = True
    lvObserv_Click
End If
EventosObs True
Liberado

End Sub

Private Sub NovaObsParcela()
Dim itmX As ListItem, z As Long

frBotao.Enabled = False
pnlObs.Visible = True
pnlObs.ZOrder 0
frTop.Enabled = False
grdExtrato.Enabled = False

CarregaObsParcela
If lvObserv.ListItems.Count > 0 Then
    lvObserv.ListItems(1).Selected = True
    lvObserv_Click
End If
EventosObs True
Liberado

End Sub

Private Sub CarregaObs()

z = SendMessage(lvObserv.hwnd, LVM_DELETEALLITEMS, 0, 0)
txtObservacao.Text = ""
Ocupado
Sql = "SELECT debitoobservacao.codreduzido, debitoobservacao.seq, debitoobservacao.dataobs, debitoobservacao.obs, debitoobservacao.userid, usuario.nomelogin "
Sql = Sql & "FROM debitoobservacao LEFT OUTER JOIN usuario ON debitoobservacao.userid = usuario.Id "
Sql = Sql & "Where CODREDUZIDO = " & Val(txtCod.Text) & " ORDER BY DATAOBS"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvObserv.ListItems.Add(, , Format(!Seq, "0000"))
        itmX.SubItems(1) = !NomeLogin
        itmX.SubItems(2) = Format(!DATAOBS, "dd/mm/yyyy")
        itmX.SubItems(3) = !obs
       .MoveNext
    Loop
   .Close
End With
Liberado

End Sub

Private Sub CarregaObsParcela()
Dim nAno As Integer, nLanc As Integer, nSeqLanc As Integer, nParc As Integer, nCompl As Integer

z = SendMessage(lvObserv.hwnd, LVM_DELETEALLITEMS, 0, 0)
txtObservacao.Text = ""

With grdExtrato
    If .Rows = 0 Then
        MsgBox "Não existem débitos.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If .SelectedRow = 0 Then Exit Sub
    nAno = .CellText(.SelectedRow, 1)
    nLanc = Val(Left$(.CellText(.SelectedRow, 2), 3))
    nSeqLanc = Val(.CellText(.SelectedRow, 3))
    nParc = IIf(.CellText(.SelectedRow, 4) = "Unica", 0, .CellText(.SelectedRow, 4))
    nComp = Val(.CellText(.SelectedRow, 5))
End With

Sql = "SELECT obsparcela.*, usuario.nomelogin FROM obsparcela INNER JOIN usuario ON obsparcela.userid = usuario.Id WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno
Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nParc
Sql = Sql & " AND CODCOMPLEMENTO=" & nComp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvObserv.ListItems.Add(, , !Seq)
        itmX.SubItems(1) = SubNull(!NomeLogin)
        itmX.SubItems(2) = Format(!Data, "dd/mm/yyyy")
        itmX.SubItems(3) = !obs
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub EventosObs(bInicio As Boolean)

cmdNovoObs.Visible = bInicio
cmdAlterarObs.Visible = bInicio
cmdExcluirObs.Visible = bInicio
cmdSairObs.Visible = bInicio
cmdGravarObs.Visible = Not bInicio
cmdCancelarObs.Visible = Not bInicio
txtObservacao.Locked = bInicio
lvObserv.Enabled = bInicio
If bInicio Then
    txtObservacao.BackColor = txtObservacao.Parent.BackColor
Else
    txtObservacao.BackColor = Branco
End If

End Sub

Private Sub EventosEF(bInicio As Boolean)

cmdNovoEF.Visible = bInicio
cmdAlterarEF.Visible = bInicio
cmdExcluirEF.Visible = bInicio
cmdRetornar.Visible = bInicio
cmdGravarEF.Visible = Not bInicio
cmdCancelarEF.Visible = Not bInicio
cmbEF.Enabled = bInicio
lvEFOrigem.Enabled = Not bInicio
cmdC1.Enabled = Not bInicio
cmdC2.Enabled = Not bInicio
cmdAllEfo.Enabled = Not bInicio
cmdDelEfo.Enabled = Not bInicio
txtDocEF.Locked = bInicio

If bInicio Then
    txtEF.Visible = False
    cmbEF.Visible = True
    cmbEF.BackColor = vbWhite
    lvEFOrigem.BackColor = frEFiscal.BackColor
Else
    cmbEF.BackColor = frEFiscal.BackColor
    lvEFOrigem.BackColor = vbWhite
End If
End Sub

Private Function ValidaMI() As Boolean
Dim aMI() As Multa, nTipo As Integer, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, x As Integer, Y As Integer
ReDim aMI(0)

MI = False
With grdExtrato
    For x = 1 To .Rows
       'APENAS AS LINHAS SELECIONADAS QUE SEJAM MULTA OU MI
        If (Val(Left(.CellText(x, 2), 3)) = 69 Or Right(.CellText(x, 2), 4) = "(MI)") And .CellText(x, 12) = "S" Then
            .SelectedRow = x
            Exit For
        End If
    Next
    If Val(Left(.CellText(.SelectedRow, 2), 3)) = 69 Or Right(.CellText(.SelectedRow, 2), 4) = "(MI)" Then
       'CARREGA DADOS DA MULTA
        nAno = Val(.CellText(.SelectedRow, 1))
        nLanc = Val(Left(.CellText(.SelectedRow, 2), 3))
        nSeq = Val(.CellText(.SelectedRow, 3))
        nParc = Val(.CellText(.SelectedRow, 4))
        nCompl = Val(.CellText(.SelectedRow, 5))
        If Val(Left(.CellText(.SelectedRow, 2), 3)) = 69 Then
            nTipo = 2 'VIEW MULTA
        Else
            nTipo = 3 'VIEW LANC
            Sql = "SELECT * FROM MULTAINFRACAO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                nAno = !Ano
                nLanc = !lancamento
                nSeq = !Sequencia
                nParc = !Parcela
                nCompl = !Complemento
               .Close
            End With
        End If
            
        'CARREGA BLOCO DAS MULTAS
        Sql = "SELECT multainfracao.CODIGO, multainfracao.ANO, multainfracao.LANCAMENTO, multainfracao.SEQUENCIA, multainfracao.PARCELA, multainfracao.COMPLEMENTO, "
        Sql = Sql & "multainfracao.CODREDUZIDO, multainfracao.ANOEXERCICIO, multainfracao.CODLANCAMENTO, multainfracao.SEQLANCAMENTO, multainfracao.NUMPARCELA,"
        Sql = Sql & "multainfracao.CODCOMPLEMENTO , debitoparcela.statuslanc FROM multainfracao INNER JOIN debitoparcela ON "
        Sql = Sql & "multainfracao.CODREDUZIDO = debitoparcela.codreduzido AND multainfracao.ANOEXERCICIO = debitoparcela.anoexercicio AND "
        Sql = Sql & "multainfracao.CODLANCAMENTO = debitoparcela.codlancamento AND multainfracao.SEQLANCAMENTO = debitoparcela.seqlancamento AND "
        Sql = Sql & "multainfracao.NumParcela = debitoparcela.NumParcela And multainfracao.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
        Sql = Sql & "WHERE multainfracao.CODREDUZIDO=" & Val(txtCod.Text) & " AND ANO=" & nAno & " AND LANCAMENTO=" & nLanc & " AND SEQUENCIA=" & nSeq & " AND PARCELA=" & nParc & " AND COMPLEMENTO=" & nCompl & " AND (STATUSLANC=3 OR STATUSLANC=20)"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then GoTo continua
            aMI(0).nAno = !Ano
            aMI(0).nLanc = !lancamento
            aMI(0).nSeq = !Sequencia
            aMI(0).nParc = !Parcela
            aMI(0).nCompl = !Complemento
            aMI(0).nStatus = !statuslanc
            aMI(0).bAchou = False
            Do Until .EOF
                ReDim Preserve aMI(UBound(aMI) + 1)
                aMI(UBound(aMI)).nAno = !AnoExercicio
                aMI(UBound(aMI)).nLanc = !CodLancamento
                aMI(UBound(aMI)).nSeq = !SeqLancamento
                aMI(UBound(aMI)).nParc = !NumParcela
                aMI(UBound(aMI)).nCompl = !CODCOMPLEMENTO
                aMI(UBound(aMI)).nStatus = !statuslanc
                aMI(UBound(aMI)).bAchou = False
               .MoveNext
            Loop
           .Close
        End With
        
continua:
       'CONFRONTA BLOCO COM OS REGISTROS SELECIONADOS
        With grdExtrato
            For x = 1 To .Rows
               'APENAS AS LINHAS SELECIONADAS QUE SEJAM MULTA OU MI
                If (Val(Left(.CellText(x, 2), 3)) = 69 Or Right(.CellText(x, 2), 4) = "(MI)") And .CellText(x, 12) = "S" Then
                    nAno = Val(.CellText(x, 1))
                    nLanc = Val(Left(.CellText(x, 2), 3))
                    nSeq = Val(.CellText(x, 3))
                    nParc = Val(.CellText(x, 4))
                    nCompl = Val(.CellText(x, 5))
                    For Y = 0 To UBound(aMI)
                        If aMI(Y).nAno = nAno And aMI(Y).nLanc = nLanc And aMI(Y).nSeq = nSeq And aMI(Y).nParc = nParc And aMI(Y).nCompl = nCompl Then
                            aMI(Y).bAchou = True 'marca a parcela
                        End If
                    Next
                End If
            Next
        End With
        
       'VERIFICA SE ALGUMA PARCELA DA MATRIZ É NEGATIVA
       If UBound(aMI) > 0 Then
            For Y = 0 To UBound(aMI)
                If aMI(Y).bAchou = False Then
                '    If aMI(Y).nStatus <> 6 Then
                        GoTo Erro
                 '   End If
                End If
            Next
        End If
        MI = True
    Else
        MI = False
        GoTo fim
    End If
End With


fim:
ValidaMI = True
Exit Function

Erro:
MsgBox "Todas as parcelas que constituem a multa de infração devem ser selecionadas.", vbCritical, "Atenção"
ValidaMI = False

End Function

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
   
   Set m_cMenuOpcoes = New cPopupMenu
   With m_cMenuOpcoes
      .hwndOwner = Me.hwnd
      .GradientHighlight = True
      
      i = .AddItem("Cancelamento de Débitos", "", 1, , , , , "mnuCancelDebito")
      .OwnerDraw(i) = True
      i = .AddItem("Anexação de Documento", "", 1, , , , , "mnuAnexaDoc")
      .OwnerDraw(i) = True
      i = .AddItem("Divida Ativa", "", 1, , , , , "mnuDA")
      .OwnerDraw(i) = True
      i = .AddItem("Ajuizamento", "", 1, , , , , "mnuAjuiza")
      .OwnerDraw(i) = True
      i = .AddItem("Observação da Parcela", "", 1, , , , , "mnuObs")
      .OwnerDraw(i) = True
      i = .AddItem("Edição da Parcela", "", 1, , , , , "mnuEditaParcela")
      .OwnerDraw(i) = True
      i = .AddItem("Correçao Parcelamento SMAR", "", 1, , , , , "mnuSmar")
      .OwnerDraw(i) = True
      i = .AddItem("Débito sem movimento", "", 1, , , , , "mnuSemMov")
      .OwnerDraw(i) = True
      i = .AddItem("Compensação de Débitos", "", 1, , , , , "mnuCompensa")
      .OwnerDraw(i) = True
      i = .AddItem("Suspenso/Em Tramite", "", 1, , , , , "mnuSuspenso")
      .OwnerDraw(i) = True
      i = .AddItem("DAM com Honorários", "", 1, , , , , "mnuDAMH")
      .OwnerDraw(i) = True
      i = .AddItem("Reativação de Parcelamento", "", 1, , , , , "mnuReativaParc")
      .OwnerDraw(i) = True
      i = .AddItem("Multa de Infração", "", 1, , , , , "mnuMultaInfracao")
      .OwnerDraw(i) = True
      i = .AddItem("Extrato para o Fórum", "", 1, , , , , "mnuExtratoForum")
      .OwnerDraw(i) = True
      i = .AddItem("Alterar Status do Lançamento", "", 1, , , , , "mnuChangeStatus")
      .OwnerDraw(i) = True
      i = .AddItem("Busca número de documento", "", 1, , , , , "mnuBuscaDoc")
      .OwnerDraw(i) = True
      i = .AddItem("Retido pelo Tomador", "", 1, , , , , "mnuRetidoTomador")
      .OwnerDraw(i) = True
      i = .AddItem("Retirar marcação de SERASA no GTI", "", 1, , , , , "mnuSerasa")
      .OwnerDraw(i) = True
   End With
   
   Set m_cMenuExtrato = New cPopupMenu
   With m_cMenuExtrato
      .hwndOwner = Me.hwnd
      .GradientHighlight = True
      
      i = .AddItem("Completo", "", 1, , , , , "mnuExtratoCompleto")
      .OwnerDraw(i) = True
      i = .AddItem("Filtrado", "", 1, , , , , "mnuExtratoFiltro")
      .OwnerDraw(i) = True
   End With
   
   Set m_cMenuInterno = New cPopupMenu
   With m_cMenuInterno
      .hwndOwner = Me.hwnd
      .GradientHighlight = True
      
      i = .AddItem("Cancelar Ajuizamento", "", 1, , , , , "mnuCancelAjuiza")
      .OwnerDraw(i) = True
      i = .AddItem("Cancelamento de Notificação ISS Eletrônico", "", 1, , , , , "mnuCancelNotifISS")
      .OwnerDraw(i) = True
      i = .AddItem("Editar Parcela", "", 1, , , , , "mnuEditParcela")
      .OwnerDraw(i) = True
      i = .AddItem("Reativar Débito Suspenso", "", 1, , , , , "mnuReativar")
      .OwnerDraw(i) = True
      i = .AddItem("Reativar Débito em Julgamento", "", 1, , , , , "mnuReativarJ")
      .OwnerDraw(i) = True
      i = .AddItem("Visualizar Dados Suspenso/Cancelado", "", 1, , , , , "mnuVerObs")
      .OwnerDraw(i) = True
      i = .AddItem("Exibir vinculo de multa de infração", "", 1, , , , , "mnuVinculoMI")
      .OwnerDraw(i) = True
      i = .AddItem("Permitir agrupamento", "", 1, , , , , "mnuAgrupa")
      .OwnerDraw(i) = True
      i = .AddItem("2ª via de auto de infração", "", 1, , , , , "mnu2viaAuto")
      .OwnerDraw(i) = True
   End With
   
   
End Sub

Private Sub CorrigeUnica()
Dim Sql As String, RdoAux As rdoResultset, aAno() As Integer, x As Integer
Exit Sub
ReDim aAno(0)
Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND NUMPARCELA=0 AND STATUSLANC=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aAno(UBound(aAno) + 1)
        aAno(UBound(aAno)) = !AnoExercicio
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(aAno)
    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & aAno(x) & " AND NUMPARCELA>0 AND STATUSLANC=3"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=1 WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
            cn.Execute Sql, rdExecDirect
           .MoveNext
        Loop
       .Close
    End With
Next

End Sub

Private Sub CarregaOrigemEF()
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDA As String, sAj As String, sEF As String
Dim nSit As Integer, itmX As ListItem, z As Long, sVencto As String, Sql As String, RdoAux As rdoResultset, x As Integer, Y As Integer
Dim sNum As String, nNum As Integer

'txtDocEF.Text = ""

z = SendMessage(lvEFOrigem.hwnd, LVM_DELETEALLITEMS, 0, 0)
CarregalvDoc

With grdExtrato
    
    For x = 1 To .Rows
        nAno = Val(.CellText(x, 1))
        nLanc = Val(Left$(.CellText(x, 2), 3))
        nSeq = Val(.CellText(x, 3))
        nParc = IIf(.CellText(x, 4) = "Unica", 0, Val(.CellText(x, 4)))
        nCompl = Val(.CellText(x, 5))
        nSit = Val(Left(.CellText(x, 6), 2))
        sVencto = .CellText(x, 7)
        sDA = .CellText(x, 8)
        sAj = .CellText(x, 9)
        sEF = .CellText(x, 14)
        
        'If nLanc <> 20 And nSit = 3 And CDate(sVencto) < CDate(Now) And (sEF = "" Or sEF = "00000/0") Then
        'If (nSit = 3 Or nSit = 4) And CDate(sVencto) < CDate(Now) And (sEF = "" Or sEF = "00000/0") Then
        If (nSit = 3 Or nSit = 4) And CDate(sVencto) < CDate(Now) And (Len(sEF) < 12) Then
        
            With lvEFDest
                If .ListItems.Count > 0 Then
                    For Y = 1 To .ListItems.Count
                        If Val(.ListItems(Y).Text) = nAno And Val(.ListItems(Y).SubItems(1)) = nLanc And Val(.ListItems(Y).SubItems(2)) = nSeq And Val(.ListItems(Y).SubItems(3)) = nParc And Val(.ListItems(Y).SubItems(4)) = nCompl Then
                            GoTo proximo
                        End If
                    Next
                End If
            End With
        
            Set itmX = lvEFOrigem.ListItems.Add(, "C" & Format(x, "0000"), nAno)
            itmX.SubItems(1) = Format(nLanc, "000")
            itmX.SubItems(2) = Format(nSeq, "000")
            itmX.SubItems(3) = Format(nParc, "00")
            itmX.SubItems(4) = Format(nCompl, "00")
            itmX.SubItems(5) = sVencto
        End If
        
proximo:
    Next

End With

For x = 1 To lvDoc.ListItems.Count
    lvDoc.ListItems(x).Checked = False
    lvDoc.ListItems(x).SubItems(1) = "0"
Next

sNum = cmbEF.Text
If sNum = "" Then Exit Sub
'nNum = Val(Left$(sNum, InStr(1, sNum, "/", vbBinaryCompare) - 1))
'nAno = Val(Right$(sNum, 4))

Sql = "SELECT * FROM EXECUCAOFISCALDOC WHERE ANOEXEC=" & nAno & " AND processocnj='" & sNum & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For x = 1 To lvDoc.ListItems.Count
            If Val(Right(lvDoc.ListItems(x).Key, 3)) = !NumDoc Then
                lvDoc.ListItems(x).Checked = True
                lvDoc.ListItems(x).SubItems(1) = !QTDE
                lvDoc.ListItems(x).SubItems(2) = SubNull(!Situacao)
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregalvDoc()
Dim z As Long, itmX As ListItem, Sql As String, RdoAux As rdoResultset

z = SendMessage(lvDoc.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT * FROM DOCUMENTOEF ORDER BY CODIGO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvDoc.ListItems.Add(, "C" & Format(!Codigo, "000"), !descricao)
        itmX.SubItems(1) = "0"
        .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub RetiraSerasa()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, sObs As String, nSeq As Integer

If imgSerasa.Visible = False Then
    MsgBox "O contribuinte não está no Serasa.", vbCritical, "Erro"
    Exit Sub
End If


If MsgBox("Deseja retirar a marcação de Serasa no GTI?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbNo Then Exit Sub

nCodReduz = Val(txtCod.Text)

Sql = "update serasa set dtsaida='" & Format(Now, "mm/dd/yyyy") & "' where codigo=" & nCodReduz
cn.Execute Sql, rdExecDirect

imgSerasa.Visible = False

sObs = "Remoção de marcação do Serasa no GTI pelo usuário " & RetornaUsuarioFullName & "."
Sql = "SELECT MAX(SEQ) AS MAXIMO FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USUARIO,DATAOBS,OBS) VALUES(" & nCodReduz & "," & nSeq & ",'GTI','" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
cn.Execute Sql, rdExecDirect

MsgBox "Removida a marcação de Serasa no GTI", vbInformation, "Atenção"

End Sub

