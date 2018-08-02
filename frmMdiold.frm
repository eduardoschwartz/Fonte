VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{B60B1875-E5CA-11D2-BC3D-78A407C10000}#1.0#0"; "ksdpanel.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Begin VB.MDIForm frmMdi 
   BackColor       =   &H00B8AC78&
   Caption         =   "Gestão de Tributação Municipal Integrada"
   ClientHeight    =   4110
   ClientLeft      =   -660
   ClientTop       =   2655
   ClientWidth     =   11325
   Icon            =   "frmMdi.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   WindowState     =   2  'Maximized
   Begin KSDPanel.Panel frTeste 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   3495
      Visible         =   0   'False
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   556
      Caption         =   "BASE DE TESTES"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   9
      ForeColor       =   16777215
      BackColor       =   255
   End
   Begin VB.Timer Timer1 
      Left            =   1050
      Top             =   1380
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   11325
      TabIndex        =   2
      Top             =   0
      Width           =   11325
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         Picture         =   "frmMdi.frx":08CA
         ScaleHeight     =   855
         ScaleWidth      =   9015
         TabIndex        =   16
         Top             =   60
         Width           =   9015
         Begin VB.PictureBox PicBack 
            Height          =   465
            Left            =   8295
            ScaleHeight     =   405
            ScaleWidth      =   540
            TabIndex        =   17
            Top             =   165
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin prjChameleon.chameleonButton cmdAjuda 
         Height          =   375
         Left            =   7050
         TabIndex        =   3
         ToolTipText     =   "Ajuda do Sistema"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "  Aj&uda   "
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":0DBC
         PICN            =   "frmMdi.frx":0DD8
         PICH            =   "frmMdi.frx":11E6
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdLogin 
         Height          =   375
         Left            =   5640
         TabIndex        =   4
         ToolTipText     =   "Efetuar Login"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "  &Login "
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   14737632
         MPTR            =   1
         MICON           =   "frmMdi.frx":15F0
         PICN            =   "frmMdi.frx":160C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdJanela 
         Height          =   375
         Left            =   4230
         TabIndex        =   5
         ToolTipText     =   "Menu Janela"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Janela"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":16A3
         PICN            =   "frmMdi.frx":16BF
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRel 
         Height          =   375
         Left            =   7050
         TabIndex        =   6
         ToolTipText     =   "Relatórios do Sistema"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Relatório"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1733
         PICN            =   "frmMdi.frx":174F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdConsulta 
         Height          =   375
         Left            =   5640
         TabIndex        =   7
         ToolTipText     =   "Consultas do Sistema"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "C&onsulta"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1823
         PICN            =   "frmMdi.frx":183F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdMobiliario 
         Height          =   375
         Left            =   4230
         TabIndex        =   8
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Mobiliário"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":193C
         PICN            =   "frmMdi.frx":1958
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdImobiliario 
         Height          =   375
         Left            =   2820
         TabIndex        =   9
         ToolTipText     =   "Dados Imobiliários"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Imobiliário"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   -2147483633
         FCOLO           =   -2147483633
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1B2D
         PICN            =   "frmMdi.frx":1B49
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCadastro 
         Height          =   375
         Left            =   1410
         TabIndex        =   10
         ToolTipText     =   "Cadastro Básico"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Cadastro"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   -2147483633
         FCOLO           =   -2147483633
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1C33
         PICN            =   "frmMdi.frx":1C4F
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdTabela 
         Height          =   375
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "Tabelas Básicas do Sistema"
         Top             =   120
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Tabelas"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   -2147483633
         FCOLO           =   -2147483633
         MCOL            =   16777215
         MPTR            =   99
         MICON           =   "frmMdi.frx":20A2
         PICN            =   "frmMdi.frx":23BC
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin AniGIFCtrl.AniGIF AniGIF1 
         Height          =   525
         Left            =   9630
         TabIndex        =   12
         Top             =   180
         Width           =   645
         BackColor       =   12632256
         PLaying         =   -1  'True
         Transparent     =   -1  'True
         Speed           =   1
         Stretch         =   2
         AutoSize        =   0   'False
         SequenceString  =   ""
         Sequence        =   0
         HTTPProxy       =   ""
         HTTPUserName    =   ""
         HTTPPassword    =   ""
         MousePointer    =   0
         GIF             =   "frmMdi.frx":2438
         ExtendWidth     =   1138
         ExtendHeight    =   926
         Loop            =   0
         AutoRewind      =   0   'False
         Synchronized    =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdOpcoes 
         Height          =   375
         Left            =   1410
         TabIndex        =   13
         ToolTipText     =   "Opções do Sistema"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Opções"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   -2147483633
         FCOLO           =   -2147483633
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1D080
         PICN            =   "frmMdi.frx":1D09C
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdTributario 
         Height          =   375
         Left            =   0
         TabIndex        =   14
         ToolTipText     =   "Dados Tributários"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Tri&butário"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   -2147483633
         FCOLO           =   -2147483633
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1D14A
         PICN            =   "frmMdi.frx":1D166
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAvançado 
         Height          =   375
         Left            =   2820
         TabIndex        =   15
         ToolTipText     =   "Recursos Avançados"
         Top             =   480
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Avançado"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   -2147483633
         FCOLO           =   -2147483633
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMdi.frx":1D54D
         PICN            =   "frmMdi.frx":1D569
         UMCOL           =   -1  'True
         SOFT            =   -1  'True
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "de Jaboticabal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   10170
         TabIndex        =   19
         Top             =   435
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Prefeitura Municipal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   10230
         TabIndex        =   18
         Top             =   210
         Width           =   1710
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   11325
      TabIndex        =   0
      Top             =   3810
      Width           =   11325
      Begin MSComctlLib.StatusBar Sbar 
         Height          =   270
         Left            =   45
         TabIndex        =   1
         Top             =   15
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   476
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   6
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Width           =   8185
               MinWidth        =   8185
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Object.Width           =   5292
               MinWidth        =   5292
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Enabled         =   0   'False
               Object.Width           =   531
               MinWidth        =   531
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   2
               Alignment       =   1
               Object.Width           =   1059
               MinWidth        =   1059
               TextSave        =   "NUM"
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Style           =   3
               Alignment       =   1
               Enabled         =   0   'False
               Object.Width           =   1059
               MinWidth        =   1059
               TextSave        =   "INS"
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   3528
               MinWidth        =   3528
               Text            =   "Data Base: 00/00/0000"
               TextSave        =   "Data Base: 00/00/0000"
            EndProperty
         EndProperty
      End
      Begin VB.Image imStatus 
         Height          =   240
         Index           =   2
         Left            =   5970
         Picture         =   "frmMdi.frx":1D907
         Top             =   390
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imStatus 
         Height          =   240
         Index           =   1
         Left            =   6240
         Picture         =   "frmMdi.frx":1DC91
         Top             =   330
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imOK 
         Height          =   240
         Left            =   11370
         Picture         =   "frmMdi.frx":1E01B
         Top             =   30
         Width           =   240
      End
      Begin VB.Image imWorking 
         Height          =   240
         Left            =   11640
         Picture         =   "frmMdi.frx":1E3A5
         Top             =   30
         Width           =   240
      End
      Begin VB.Image imStatus 
         Height          =   240
         Index           =   0
         Left            =   5100
         Picture         =   "frmMdi.frx":1E72F
         Top             =   420
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   3795
      Top             =   3345
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   63
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1EAB9
            Key             =   "NULL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1EC13
            Key             =   "TABELA"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1EF2F
            Key             =   "BAIRRO"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1F093
            Key             =   "CIDADE"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1F1EF
            Key             =   "LOTEAMENTO"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1F353
            Key             =   "TABSISTEMA"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":1FC2F
            Key             =   "CIDADAO"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2050B
            Key             =   "LOGRADOURO"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":20827
            Key             =   "CADIMOB"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":21103
            Key             =   "LOG"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":219DF
            Key             =   "SECURITY"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":21CFF
            Key             =   "BLOQ"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":225DB
            Key             =   "CONFIG"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":22EB7
            Key             =   "CNSIMOVEL"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":23017
            Key             =   "AJUDA"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":23177
            Key             =   "TOOL"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":23493
            Key             =   "UNIFICA"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":237AF
            Key             =   "DESMEMBRA"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":23ACB
            Key             =   "CONDOMINIO"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":23DEB
            Key             =   "FACEQ"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2410B
            Key             =   "IMUNE"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2442B
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":24587
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":246E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2483F
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":24953
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":24B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":24C43
            Key             =   "INTERNET"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":273F7
            Key             =   "BENF"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":27553
            Key             =   "CTCO"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":276AF
            Key             =   "CTPR"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2780B
            Key             =   "TIPC"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":27967
            Key             =   "USOC"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":27AC3
            Key             =   "USOT"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":27C1F
            Key             =   "MOEDA"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":27F3B
            Key             =   "FPED"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":28097
            Key             =   "FSIT"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":281F3
            Key             =   "FTOP"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2834F
            Key             =   "FDIS"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2866B
            Key             =   "FPRO"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":28987
            Key             =   "FGEN"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":28CA3
            Key             =   "FGLE"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":28FBF
            Key             =   "FBEN"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2911B
            Key             =   "FCAT"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":29277
            Key             =   "TEXP"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":293D7
            Key             =   "TIPC2"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":29533
            Key             =   "BANCO"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2984F
            Key             =   "CALCGERAL"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":29B6B
            Key             =   "FER1"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":29C6F
            Key             =   "FERI1"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":29D53
            Key             =   "CDBA1"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":29E17
            Key             =   "PAGA1"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2A1F7
            Key             =   "DEBA1"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2A5E7
            Key             =   "CNNU1"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2A6F3
            Key             =   "CNLA1"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2A807
            Key             =   "PAGA"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2AB2B
            Key             =   "FERI"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2B407
            Key             =   "CREP"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2BCE3
            Key             =   "CNLA"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2C007
            Key             =   "CDBA"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2C8E3
            Key             =   "REPA"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2D1BF
            Key             =   "DEBA"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMdi.frx":2DA9B
            Key             =   "CNNU"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuTabela As cPopupMenu
Attribute m_cMenuTabela.VB_VarHelpID = -1
Public WithEvents m_cMenuCadastro As cPopupMenu
Attribute m_cMenuCadastro.VB_VarHelpID = -1
Public WithEvents m_cMenuConsulta As cPopupMenu
Attribute m_cMenuConsulta.VB_VarHelpID = -1
Public WithEvents m_cMenuImobiliario As cPopupMenu
Attribute m_cMenuImobiliario.VB_VarHelpID = -1
Public WithEvents m_cMenuMobiliario As cPopupMenu
Attribute m_cMenuMobiliario.VB_VarHelpID = -1
Public WithEvents m_cMenuTributario As cPopupMenu
Attribute m_cMenuTributario.VB_VarHelpID = -1
Public WithEvents m_cMenuOpcoes As cPopupMenu
Attribute m_cMenuOpcoes.VB_VarHelpID = -1
Public WithEvents m_cMenuAjuda As cPopupMenu
Attribute m_cMenuAjuda.VB_VarHelpID = -1
Public WithEvents m_cMenuAvancado As cPopupMenu
Attribute m_cMenuAvancado.VB_VarHelpID = -1
Public WithEvents m_cMenuJanela As cPopupMenu
Attribute m_cMenuJanela.VB_VarHelpID = -1
Public WithEvents m_cMenuRelatorio As cPopupMenu
Attribute m_cMenuRelatorio.VB_VarHelpID = -1

Dim RunOnce As Boolean
Dim bPush As Boolean
Dim AllowPopup As Boolean 'This is for Pop-up windows

Private Sub cmdAjuda_Click()
lIndex = m_cMenuAjuda.ShowPopupMenu(cmdAjuda.Left, cmdAjuda.Top, cmdAjuda.Left, cmdAjuda.Top, Me.ScaleWidth - cmdAjuda.Left - cmdAjuda.Width, cmdAjuda.Top + cmdAjuda.Height, False)

If (lIndex > 0) Then
   Picture1.Refresh
   Select Case m_cMenuAjuda.ItemKey(lIndex)
        Case "mnuConteúdo"
            With hHelp
              .CHMFile = sPathHelp & "\Tribut.chm"
              .HHWindow = "Main"
              .HHDisplayContents
            End With
        Case "mnuIndice"
            With hHelp
              .CHMFile = sPathHelp & "\Tribut.chm"
              .HHWindow = "Main"
              .HHDisplayIndex
            End With
        Case "mnuLocalizar"
            With hHelp
              .CHMFile = sPathHelp & "\Tribut.chm"
              .HHWindow = "Main"
              .HHDisplaySearch
            End With
        Case "mnuSobre"
            frmAbout.show 1
   End Select
Else
   lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdavançado_Click()
  lIndex = m_cMenuAvancado.ShowPopupMenu(cmdAvançado.Left, cmdAvançado.Top, cmdAvançado.Left, cmdAvançado.Top, Me.ScaleWidth - cmdAvançado.Left - cmdAvançado.Width, cmdAvançado.Top + cmdAvançado.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuAvancado.ItemKey(lIndex)
        Case "mnuSql"
            Set frm = frmSql
        Case "mnuGeraDebito"
            Set frm = frmGeraDebito
   End Select
   frm.show
   frm.ZOrder 0
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdCadastro_Click()
  lIndex = m_cMenuCadastro.ShowPopupMenu(cmdCadastro.Left, cmdCadastro.Top, cmdCadastro.Left, cmdCadastro.Top, Me.ScaleWidth - cmdCadastro.Left - cmdCadastro.Width, cmdCadastro.Top + cmdCadastro.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuCadastro.ItemKey(lIndex)
        Case "mnuCidadao"
            Set frm = frmCidadao
        Case "mnuBanco"
            Set frm = frmBanco
        Case "mnuFeriado"
            Set frm = frmFeriados
        Case "mnuLivro"
            Set frm = frmLivro
   End Select
   frm.show
   frm.ZOrder 0
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdConsulta_Click()
  lIndex = m_cMenuConsulta.ShowPopupMenu(cmdConsulta.Left, cmdConsulta.Top + 400, cmdConsulta.Left, cmdConsulta.Top, Me.ScaleWidth - cmdConsulta.Left - cmdConsulta.Width, cmdConsulta.Top + cmdConsulta.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuConsulta.ItemKey(lIndex)
        Case "mnuCnsProcesso"
            Set frm = frmCnsProcesso
   End Select
   frm.show
   frm.ZOrder 0
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdImobiliario_Click()
lIndex = m_cMenuImobiliario.ShowPopupMenu(cmdImobiliario.Left, cmdImobiliario.Top, cmdImobiliario.Left, cmdImobiliario.Top, Me.ScaleWidth - cmdImobiliario.Left - cmdImobiliario.Width, cmdImobiliario.Top + cmdImobiliario.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuImobiliario.ItemKey(lIndex)
        Case "mnuCadImob"
            Set frm = frmCadImob
        Case "mnuCondominio"
            Set frm = frmCadCondominio
        Case "mnuLogr"
            Set frm = frmLogradouro
        Case "mnuFaceQuadra"
            Set frm = frmFaceQuadra
        Case "mnuCnsLog"
            Set frm = frmCnsLogradouro
        Case "mnuCnsImovel"
            Set frm = frmCnsImovel
        Case "mnuDetImovel"
            Set frm = frmDadosImovel
        Case "mnuDesmem"
            Set frm = frmDesmembramento
        Case "mnuUnifica"
            Set frm = frmUnifica
        Case "mnuImun"
            Set frm = frmIsencao
   End Select
   frm.show
   frm.ZOrder 0
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdJanela_Click()
Dim x As Integer

lIndex = m_cMenuJanela.ShowPopupMenu(cmdJanela.Left, cmdJanela.Top, cmdJanela.Left, cmdJanela.Top, Me.ScaleWidth - cmdJanela.Left - cmdJanela.Width, cmdJanela.Top + cmdJanela.Height, False)
If (lIndex > 0) Then
   Picture1.Refresh
   Select Case m_cMenuJanela.ItemKey(lIndex)
        Case "mnuCloseAll"
INICIO:
            For x = 0 To Forms.Count - 1
                If Forms(x).Name <> "frmMdi" Then
                   Unload Forms(x)
                   GoTo INICIO:
                End If
            Next
        Case Else
            For x = 0 To Forms.Count - 1
                If Forms(x).Name = m_cMenuJanela.ItemKey(lIndex) Then
                   Forms(x).ZOrder 0
                End If
            Next
   End Select
Else
   lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdLogin_Click()
    If Forms.Count > 1 Then
       If MsgBox("Deseja fechar todas as telas e bloquear o sistema ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
          mnuCloseAll_Click
          frmLogin.show vbModal
       End If
    Else
       frmLogin.show vbModal
    End If
End Sub

Private Sub cmdMobiliario_Click()
lIndex = m_cMenuMobiliario.ShowPopupMenu(cmdMobiliario.Left, cmdMobiliario.Top, cmdMobiliario.Left, cmdMobiliario.Top, Me.ScaleWidth - cmdMobiliario.Left - cmdMobiliario.Width, cmdMobiliario.Top + cmdMobiliario.Height, False)
If (lIndex > 0) Then
   Picture1.Refresh
   Ocupado
   Select Case m_cMenuMobiliario.ItemKey(lIndex)
        Case "mnuTabAtivTL"
            Set frm = frmAtiv
        Case "mnuTabAtivISS"
            Set frm = frmAtivISS
        Case "mnuEscContab"
             Set frm = frmEscContab
        Case "mnuVigSan"
             Set frm = frmVigSanitaria
        Case "mnuCadMobiliario"
             Set frm = frmCadMob
        Case "mnuSuspende"
             Set frm = frmSuspReativ
        Case "mnuCnsEmpresa"
             Set frm = frmCnsMob
        Case "mnuCnsNF"
             Set frm = frmNF
   End Select
   frm.show
   frm.ZOrder 0
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdOpcoes_Click()
lIndex = m_cMenuOpcoes.ShowPopupMenu(cmdOpcoes.Left, cmdOpcoes.Top, cmdOpcoes.Left, cmdOpcoes.Top, Me.ScaleWidth - cmdOpcoes.Left - cmdOpcoes.Width, cmdOpcoes.Top + cmdOpcoes.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuOpcoes.ItemKey(lIndex)
        Case "mnuLog"
            Set frm = frmLog
        Case "mnuConfig"
            Set frm = frmConfig
        Case "mnuLock"
            If Forms.Count > 1 Then
               If MsgBox("Deseja fechar todas as telas e bloquear o sistema ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                  mnuCloseAll_Click
                  Set frm = frmLogin
               Else
                  Liberado
                  Exit Sub
               End If
            Else
               Set frm = frmLogin
            End If
        Case "mnuUser"
            If InStr(1, UCase$(cn.Connect), "DEVELOPER", vbBinaryCompare) > 0 Then
                Liberado
                MsgBox "Não é possível acessar a segurança do sistema porque a ODBC esta configurada para acessar a base de Testes TributacaoDeveloper", vbCritical, "Atenção"
                Exit Sub
            End If
            Set frm = frmUser
        Case "mnuSegEvento"
            Set frm = frmEventSecurity
        Case "mnuAtribSeg"
            If InStr(1, UCase$(cn.Connect), "DEVELOPER", vbBinaryCompare) > 0 Then
                Liberado
                MsgBox "Não é possível acessar a segurança do sistema porque a ODBC esta configurada para acessar a base de Testes ", vbCritical, "Atenção"
                Exit Sub
            End If
            Set frm = frmSecurity
   End Select
   If frm.Name = "frmLogin" Then
       frm.show vbModal
   Else
        frm.show
        frm.ZOrder 0
   End If
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub cmdRel_Click()
lIndex = m_cMenuRelatorio.ShowPopupMenu(cmdRel.Left, cmdRel.Top + 380, cmdRel.Left, cmdRel.Top, Me.ScaleWidth - cmdRel.Left - cmdRel.Width, cmdRel.Top + cmdRel.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuRelatorio.ItemKey(lIndex)
        Case "mnuLCID"
            frmReport.ShowReport "CIDADE", frmMdi.hwnd, Me.hwnd
        Case "mnuLBAI"
            frmReport.ShowReport "BAIRRO", frmMdi.hwnd, Me.hwnd
        Case "mnuLLOT"
            frmReport.ShowReport "LOTEAMENTO", frmMdi.hwnd, Me.hwnd
        Case "mnuTMOE"
            frmReport.ShowReport "MOEDAS", frmMdi.hwnd, Me.hwnd
        Case "mnuTUFI"
            frmReport.ShowReport "UFIR", frmMdi.hwnd, Me.hwnd
        Case "mnuTEXP"
            frmReport.ShowReport "EXPEDIENTE", frmMdi.hwnd, Me.hwnd
        Case "mnuFBEN"
            frmReport.ShowReport "FATORBENFEITORIA", frmMdi.hwnd, Me.hwnd
        Case "mnuFCAT"
            frmReport.ShowReport "FATORCATEGORIA", frmMdi.hwnd, Me.hwnd
        Case "mnuFDIS"
            frmReport.ShowReport "FATORDISTRITO", frmMdi.hwnd, Me.hwnd
        Case "mnuFGLE"
            frmReport.ShowReport "FATORGLEBA", frmMdi.hwnd, Me.hwnd
        Case "mnuFPED"
            frmReport.ShowReport "FATORPEDOLOGIA", frmMdi.hwnd, Me.hwnd
        Case "mnuFPRO"
            frmReport.ShowReport "FATORPROFUN", frmMdi.hwnd, Me.hwnd
        Case "mnuFSIT"
            frmReport.ShowReport "FATORSITUACAO", frmMdi.hwnd, Me.hwnd
        Case "mnuFTOP"
            frmReport.ShowReport "FATORTOPOGRAFIA", frmMdi.hwnd, Me.hwnd
        Case "mnuBENF"
            frmReport.ShowReport "BENFEITORIAS", frmMdi.hwnd, Me.hwnd
        Case "mnuCATC"
            frmReport.ShowReport "CATEGCONSTR", frmMdi.hwnd, Me.hwnd
        Case "mnuCATP"
            frmReport.ShowReport "CATEGPROPR", frmMdi.hwnd, Me.hwnd
        Case "mnuPPAR"
            frmReport.ShowReport "PARAMPARCELAS", frmMdi.hwnd, Me.hwnd
        Case "mnuPEDO"
            frmReport.ShowReport "PEDOLOGIA", frmMdi.hwnd, Me.hwnd
        Case "mnuPGEN"
            frmReport.ShowReport "PLANTAGENERICA", frmMdi.hwnd, Me.hwnd
        Case "mnuTCON"
            frmReport.ShowReport "TIPOCONSTR", frmMdi.hwnd, Me.hwnd
        Case "mnuTPLO"
            frmReport.ShowReport "TIPOLOGRADOURO", frmMdi.hwnd, Me.hwnd
        Case "mnuTTLO"
            frmReport.ShowReport "TITLOGRADOURO", frmMdi.hwnd, Me.hwnd
        Case "mnuUSOC"
            frmReport.ShowReport "USOCONSTR", frmMdi.hwnd, Me.hwnd
        Case "mnuUSOT"
            frmReport.ShowReport "USOTERRENO", frmMdi.hwnd, Me.hwnd
        Case "mnuALVA"
            frmReport.ShowReport "ALVARA", frmMdi.hwnd, Me.hwnd
        Case "mnu2Via"
            frmEmissao2Via.show
        Case "mnuCertDeb"
            frmCertidao.show
            frmCertidao.lblTipo.Caption = "CERTIDÃO DE DÉBITO"
            frmCertidao.lblCodCert.Caption = 1
        Case "mnuCertEnd"
            frmCertidao.show
            frmCertidao.lblTipo.Caption = "CERTIDÃO DE ENDEREÇO ATUALIZADO"
            frmCertidao.lblCodCert.Caption = 2
        Case "mnuCertVaV"
            frmCertidao.show
            frmCertidao.lblTipo.Caption = "CERTIDÃO DE VALOR VENAL"
            frmCertidao.lblCodCert.Caption = 4
        Case "mnuCertDem"
            frmCertidao.show
            frmCertidao.lblTipo.Caption = "CERTIDÃO DE DEMOLIÇÃO"
            frmCertidao.lblCodCert.Caption = 5
        Case "mnuCertIse"
            frmCertidao.show
            frmCertidao.lblTipo.Caption = "CERTIDÃO DE ISENÇÃO"
            frmCertidao.lblCodCert.Caption = 6
        Case "mnuTermConf"
            frmConfissaoDivida.show
        Case "mnuRelAtivTL"
            frmReport.ShowReport "ATIVIDADETL", frmMdi.hwnd, Me.hwnd
        Case "mnuRelAtivISS"
            frmReport.ShowReport "ATIVIDADEISS", frmMdi.hwnd, Me.hwnd
        Case "mnuRelAtivISSFixo"
            frmReport.ShowReport "ATIVIDADEISSFIXO", frmMdi.hwnd, Me.hwnd
        Case "mnuRelDevedorVariavel"
            frmReport.ShowReport "ISSVARIAVELNAOPAGO", frmMdi.hwnd, Me.hwnd
        Case "mnuRelCartaCobrança"
            frmReport.ShowReport "COBRANCAAMIGAVEL", frmMdi.hwnd, Me.hwnd
        Case "mnuRelAjuiza"
            frmRelAjuiza.show
   End Select

   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdTabela_Click()
Dim lIndex As Long
Dim frm As Form

lIndex = m_cMenuTabela.ShowPopupMenu(cmdTabela.Left, cmdTabela.Top, cmdTabela.Left, cmdTabela.Top, Me.ScaleWidth - cmdTabela.Left - cmdTabela.Width, cmdTabela.Top + cmdTabela.Height, False)

If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuTabela.ItemKey(lIndex)
        Case "mnuBairro"
            Set frm = frmBairro
        Case "mnuCidade"
            Set frm = frmCidade
        Case "mnuLoteamento"
            Set frm = frmLoteam
        Case "mnuTitLog"
            Set frm = frmTitLogradouro
        Case "mnuTipoLog"
            Set frm = frmTipoLogradouro
        Case "mnuTabSistemaBenf"
            sParamForm = "BENF"
            Set frm = frmParam1
        Case "mnuTabSistemaCatC"
            sParamForm = "CATC"
            Set frm = frmParam1
        Case "mnuTabSistemaCatP"
            sParamForm = "CATP"
            Set frm = frmParam1
        Case "mnuTabSistemaTipC"
            sParamForm = "TIPC"
            Set frm = frmParam1
        Case "mnuTabSistemaUsoC"
            sParamForm = "USOC"
            Set frm = frmParam1
        Case "mnuTabSistemaUsoT"
            sParamForm = "USOT"
            Set frm = frmParam1
        Case "mnuTabSistemaMoeda"
            sParamForm = "MOED"
            Set frm = frmParam1
        Case "mnuTabSistemaUfir"
            sParamForm = "UFIR"
            Set frm = frmParam1
        Case "mnuTabSistemaFPED"
            sParamForm = "PEDO"
            Set frm = frmParam2
        Case "mnuTabSistemaFSIT"
            sParamForm = "SITU"
            Set frm = frmParam2
        Case "mnuTabSistemaFTOP"
            sParamForm = "TOPO"
            Set frm = frmParam2
        Case "mnuTabSistemaFDIS"
            sParamForm = "DIST"
            Set frm = frmParam2
        Case "mnuTabSistemaFPRO"
            sParamForm = "FPRO"
            Set frm = frmParam2
        Case "mnuTabSistemaFGEN"
            sParamForm = "PGEN"
            Set frm = frmParam2
        Case "mnuTabSistemaFGLE"
            sParamForm = "FGLE"
            Set frm = frmParam2
        Case "mnuTabSistemaFBEN"
            sParamForm = "FBEN"
            Set frm = frmParam2
        Case "mnuTabSistemaFCAT"
            sParamForm = "FCAT"
            Set frm = frmParam2
        Case "mnuTabSistemaTEXP"
            Set frm = frmTxExpe
        Case "mnuTabSistemaPPARC"
            Set frm = frmParamParcela
        Case "mnuTabSistemaINDC"
            Set frm = frmIndCor
        Case "mnuTabSistemaTLAN"
            Set frm = frmLancamento
        Case "mnuTabSistemaTTRI"
            Set frm = frmTributo
        Case "mnuTabSistemaTTLA"
            Set frm = frmTributoLanc
        Case "mnuTabTributoAliq"
            Set frm = frmTributoAliquota
        Case "mnuChangeUser"
            If Forms.Count > 1 Then
               If MsgBox("Deseja fechar todas as telas e bloquear o sistema ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                  mnuCloseAll_Click
                  Set frm = frmLogin
               End If
            Else
               Set frm = frmLogin
            End If
        Case "mnuClose"
            Unload Me
            Exit Sub
   End Select
   If frm.Name = "frmLogin" Then
       frm.show vbModal
   Else
        frm.show
        frm.ZOrder 0
   End If
   Liberado
Else
    lIndex = 0 ' cancelled the menu.
End If

End Sub

Private Sub cmdTributario_Click()
Dim z As Variant
lIndex = m_cMenuTributario.ShowPopupMenu(cmdTributario.Left, cmdTributario.Top, cmdTributario.Left, cmdTributario.Top, Me.ScaleWidth - cmdTributario.Left - cmdTributario.Width, cmdTributario.Top + cmdTributario.Height, False)
If (lIndex > 0) Then
    Picture1.Refresh
   Ocupado
   Select Case m_cMenuTributario.ItemKey(lIndex)
        Case "mnuCalcGeral"
            Set frm = frmCalcGeral
        Case "mnuCalcGeralISS"
            Set frm = frmCalcGeralISS
        Case "mnuIsento"
            Set frm = frmIsencao
        Case "mnuRecalc"
            Set frm = frmCalcGeral
        Case "mnuArqLaser"
            Set frm = frmArquivoLaser
        Case "mnuImportaArq"
            Set frm = frmArqBanco
        Case "mnuOptanteDA"
            Set frm = frmOptanteDA
        Case "mnuPagAuto"
            Set frm = frmPagAutomatico
        Case "mnuDebAutomatico"
            Set frm = frmDebAutomatico
        Case "mnuAnaliseReceita"
            Set frm = frmAnaliseReceita
        Case "mnuRelPagamento"
            Set frm = frmRelBanco
        Case "mnuBaixaManual"
            Set frm = frmBaixaManual
        Case "mnu2ViaLaser"
            Set frm = frm2ViaLaser
        Case "mnuMovimento"
            Set frm = frmMovEconomico
        Case "mnuReparcelDebito"
            Set frm = frmReparcelamento
        Case "mnuCancelReparc"
            Set frm = frmCancelReparc
        Case "mnuManAluguel"
            Set frm = frmManAluguel
        Case "mnuITBI"
            Set frm = frmITBI
        Case "mnuDividaAtiva"
            Set frm = frmDividaAtiva
        Case "mnuEmiteLivro"
            Set frm = frmGeraLivro
        Case "mnuImpAluguel"
            Set frm = frmEmissaoAluguel
        Case "mnuCnsNumDoc"
            Set frm = frmDoc
        Case "mnuCnsNumDV"
            z = InputBox("Digite o numero do documento.", "Retorna DV")
            If Val(z) > 0 Then
                MsgBox "DV -> " & RetornaDVNumDoc(CLng(z)), vbExclamation, "DV"
            Else
                MsgBox "Documento Inválido", vbExclamation, "Atenção"
            End If
            Liberado
            Exit Sub
        Case "CnsDebitoImob"
            Set frm = frmDebitoImob
   End Select
   If frm.Name = "frmDoc" Then
      frm.show vbModeless, Me
   Else
      frm.show
   End If
   frm.ZOrder 0
   Liberado
Else
   lIndex = 0 ' cancelled the menu.
End If
End Sub

Private Sub MDIForm_Resize()
Unload frmCnsParcela
End Sub

Private Sub MDIForm_Activate()
On Error Resume Next

Picture3.Visible = False
If Not RunOnce Then
     frmLogin.show 1
    RunOnce = True
Else
    Unload frmLogin
End If

End Sub

Private Sub MDIForm_Load()

Ocupado
RunOnce = False
'MsgBox "A PARTIR DE HOJE 31/05 TODOS OS REPARCELAMENTOS QUE POSSUIREM MAIS DE 1 CÓDIGO DENTRO DO MESMO PROCESSO DEVEM SER EFETUADOS TODOS JUNTOS DENTRO DO MESMO REPARCELAMENTO.", vbCritical, "ATENÇÃO"
'Log Logon, Me.Name, Nenhum, "Logon no Sistema"
Set gtiObj = New gtiProc.Tmuna
Dim l As Long
Dim lIndex As Long
Dim lC As Long
Dim lMajor As Long, lMinor As Long, lBuild As Long

lMajor = App.Major
lMinor = App.Minor
lBuild = App.Revision
Me.Caption = Me.Caption & " - Versão: " & lMajor & "." & lMinor & "." & lBuild
Me.Visible = False
frmDummyChild.TileMDI 0
Unload frmDummyChild
Me.Visible = True
 '   Log Logon, Me.Name, Nenhum, "Logon no Sistema"
MontaMenu


Liberado

End Sub

Private Sub MontaMenu()

MontaMenuTabelas
MontaMenuCadastro
MontaMenuConsulta
MontaMenuImobiliario
MontaMenuMobiliario
MontaMenuTributario
MontaMenuOpcoes
MontaMenuAjuda
MontaMenuAvancado
MontaMenuJanela
MontaMenuRelatorio

End Sub

Private Sub MontaMenuTabelas()

   Set m_cMenuTabela = New cPopupMenu
   With m_cMenuTabela
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("-Cadastro Básico")
      .OwnerDraw(i) = True
      i = .AddItem("&Cadastro de Bairros", "Cadastra os Bairros", 1, , 3, , , "mnuBairro")
      .OwnerDraw(i) = True
      i = .AddItem("C&adastro de Cidades", "Cadastra as Cidades", 1, , 4, , , "mnuCidade")
      .OwnerDraw(i) = True
      i = .AddItem("T&ítulo de Logradouro", "Cadastra os Títulos dos Logradouros", 1, , , , , "mnuTitLog")
      .OwnerDraw(i) = True
      i = .AddItem("Tip&o de Logradouro", "Cadastra os Tipos de Logradouro", 1, , , , , "mnuTipoLog")
      .OwnerDraw(i) = True
      i = .AddItem("-Parametrização do Sistema")
      .OwnerDraw(i) = True
      i = .AddItem("T&abelas Básicas", , 1, , , , , "mnuTabSistemaTabelasBásicas")
      .OwnerDraw(i) = True
      .AddItem "Benfeitoria", "Cadastro de Benfeitorias", 2, i, 28, , , "mnuTabSistemaBenf"
      .AddItem "Categoria da Construção", "Categoria da Construção", 2, i, 29, , , "mnuTabSistemaCatC"
      .AddItem "Categoria da Propriedade", "Categoria da Propriedade", 2, i, 30, , , "mnuTabSistemaCatP"
      .AddItem "Tipo de Construção", "Tipo de Construção", 2, i, 31, , , "mnuTabSistemaTipC"
      .AddItem "Uso da Construção", "Uso da Construção", 2, i, 32, , , "mnuTabSistemaUsoC"
      .AddItem "Uso do Terreno", "Uso do Terreno", 2, i, 33, , , "mnuTabSistemaUsoT"
      .AddItem "Tabela de Moedas", "Tabela de Moedas", 2, i, 34, , , "mnuTabSistemaMoeda"
      .AddItem "Tabela de UFIR", "Tabela de UFIR", 2, i, , , , "mnuTabSistemaUfir"
      i = .AddItem("&Fatores", , 1, , , , , "mnuFatores")
      .OwnerDraw(i) = True
      .AddItem "Fator Pedologia", "Fator Pedologia", 3, i, 35, , , "mnuTabSistemaFPED"
      .AddItem "Fator Situação", "Fator Situação", 3, i, 36, , , "mnuTabSistemaFSIT"
      .AddItem "Fator Topografia", "Fator Topografia", 3, i, 37, , , "mnuTabSistemaFTOP"
      .AddItem "Fator Distrito", "Fator Distrito", 3, i, 38, , , "mnuTabSistemaFDIS"
      .AddItem "Fator Profundidade", "Fator Profundidade", 3, i, 39, , , "mnuTabSistemaFPRO"
      .AddItem "Planta Genérica", "Planta Genérica", 3, i, 40, , , "mnuTabSistemaFGEN"
      .AddItem "Fator Gleba", "Fator Gleba", 3, i, 41, , , "mnuTabSistemaFGLE"
      .AddItem "Fator Benfeitoria", "Fator Benfeitoria", 3, i, 42, , , "mnuTabSistemaFBEN"
      .AddItem "Fator Categoria", "Fator Categoria", 3, i, 43, , , "mnuTabSistemaFCAT"
       i = .AddItem("Cálculo", , 4, , , , , "mnuCalculo")
      .OwnerDraw(i) = True
      .AddItem "Taxa de Expediente", "Taxa de Expediente", 4, i, 44, , , "mnuTabSistemaTEXP"
      .AddItem "Parâmetro de Parcelas", "Parâmetro de Parcelas", 4, i, , , , "mnuTabSistemaPPARC"
      .AddItem "Índice de Correção", "Índice de Correção", 4, i, , , , "mnuTabSistemaINDC"
      .AddItem "Tabela de Lançamentos", "Tabela de Lançamentos", 4, i, , , , "mnuTabSistemaTLAN"
      .AddItem "Tabela de Tributos", "Tabela de Tributos", 4, i, , , , "mnuTabSistemaTTRI"
      .AddItem "Tabela de Tributos/Lançamentos", "Tabela de Tributos/Lançamentos", 4, i, , , , "mnuTabSistemaTTLA"
      .AddItem "Tabela de Preços Públicos", "Tabela de Preços Públicos", 4, i, , , , "mnuTabTributoAliq"
      i = .AddItem("-Saída")
      .OwnerDraw(i) = True
      i = .AddItem("T&rocar de Usuário", , , , , , , "mnuChangeUser")
      .OwnerDraw(i) = True
      i = .AddItem("&Sair do Sistema", , , , , , , "mnuClose")
      .OwnerDraw(i) = True
      
   End With

End Sub

Private Sub MontaMenuCadastro()

   Set m_cMenuCadastro = New cPopupMenu
   With m_cMenuCadastro
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("-Cadastro Básico")
      .OwnerDraw(i) = True
      i = .AddItem("&Cadastro de Cidadão", "Cadastra os Cidadões", 1, , 6, , , "mnuCidadao")
      .OwnerDraw(i) = True
      i = .AddItem("C&adastro de Bancos", "Cadastra os Bancos", 1, , 49, , , "mnuBanco")
      .OwnerDraw(i) = True
      i = .AddItem("Cadastro de &Feriados", "Cadastra os Feriados", 1, , 46, , , "mnuFeriado")
      .OwnerDraw(i) = True
      i = .AddItem("Livro Divida Ativa", "Livro Divida Ativa", 4, , , , , "mnuLivro")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MontaMenuConsulta()

   Set m_cMenuConsulta = New cPopupMenu
   With m_cMenuConsulta
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      i = .AddItem("-Consultas Básicas")
      .OwnerDraw(i) = True
      i = .AddItem("Consulta rápida a números de processo.", "", 1, , , , , "mnuCnsProcesso")
      .OwnerDraw(i) = True
      
   End With
   
End Sub


Private Sub MontaMenuJanela()

   Set m_cMenuJanela = New cPopupMenu
   With m_cMenuJanela
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("&Fechar Todas", "Fechar Todas", 1, , 6, , , "mnuCloseAll")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MontaMenuImobiliario()

   Set m_cMenuImobiliario = New cPopupMenu
   With m_cMenuImobiliario
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("-Cadastro Básico")
      .OwnerDraw(i) = True
      i = .AddItem("&Cadastro de Imóvel", "Cadastro de Imóvel", 1, , 8, , , "mnuCadImob")
      .OwnerDraw(i) = True
      i = .AddItem("Cadastro de Cond&omínios", "Cadastro de Cond&omínios", 1, , 18, , , "mnuCondominio")
      .OwnerDraw(i) = True
      i = .AddItem("Cadastro de &Logradouros", "Cadastro de &Logradouros", 1, , 7, , , "mnuLogr")
      .OwnerDraw(i) = True
      i = .AddItem("Cadastro de &Face de Quadra", "Cadastro de Face de Quadra", 1, , 19, , , "mnuFaceQuadra")
      .OwnerDraw(i) = True
      i = .AddItem("-Consulta")
      .OwnerDraw(i) = True
      i = .AddItem("Con&sulta de Logradouros", "Consulta de Logradouros", 1, , , , , "mnuCnsLog")
      .OwnerDraw(i) = True
      i = .AddItem("Consulta de &Imóveis", "Consulta de Imóveis", 1, , 13, , , "mnuCnsImovel")
      .OwnerDraw(i) = True
      i = .AddItem("Detalhes do &Imóvel", "Detalhes do &Imóvel", 1, , , , , "mnuDetImovel")
      .OwnerDraw(i) = True
      i = .AddItem("-Atividades")
      .OwnerDraw(i) = True
      i = .AddItem("&Desmembramento", "Desmembramento", 1, , 17, , , "mnuDesmem")
      .OwnerDraw(i) = True
      i = .AddItem("&Unificação", "Unificação", 1, , 16, , , "mnuUnifica")
      .OwnerDraw(i) = True
      i = .AddItem("&Imunidade/Isenção", "Imunidade/Isenção", 1, , , , , "mnuImun")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MontaMenuMobiliario()

   Set m_cMenuMobiliario = New cPopupMenu
   With m_cMenuMobiliario
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("-Cadastro Básico")
      .OwnerDraw(i) = True
      i = .AddItem("&Atividades para Taxa de Licença", "Atividades para Taxa de Licença", 1, , 8, , , "mnuTabAtivTL")
      .OwnerDraw(i) = True
      i = .AddItem("At&ividades para Cobrança de ISS", "Atividades para Cobrança de ISS", 1, , 8, , , "mnuTabAtivISS")
      .OwnerDraw(i) = True
      i = .AddItem("Atividades para Vigilância &Sanitária", "Atividades para Vigilância &Sanitária", 1, , , , , "mnuVigSan")
      .OwnerDraw(i) = True
      i = .AddItem("&Escritórios Contábeis", "Escritórios Contábeis", 1, , , , , "mnuEscContab")
      .OwnerDraw(i) = True
      i = .AddItem("Cadastro &Mobiliário", "Cadastro Mobiliário", 1, , 18, , , "mnuCadMobiliario")
      .OwnerDraw(i) = True
      i = .AddItem("Suspenção/Reativação de Empresas", "Suspenção/Reativação de Empresas", 1, , , , , "mnuSuspende")
      .OwnerDraw(i) = True
      i = .AddItem("-Consulta")
      .OwnerDraw(i) = True
      i = .AddItem("Consulta de &Empresas", "Consulta de Empresas", 1, , 13, , , "mnuCnsEmpresa")
      .OwnerDraw(i) = True
      i = .AddItem("Consulta Notas Emitidas", "Consulta Notas Emitidas", 1, , , , , "mnuCnsNF")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MontaMenuTributario()

   Set m_cMenuTributario = New cPopupMenu
   With m_cMenuTributario
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("-Cálculo")
      .OwnerDraw(i) = True
      i = .AddItem("Cálculo Geral", "Cálculo Geral", 1, , , , , "mnuCalcTit")
      .OwnerDraw(i) = True
      .AddItem "&Cálculo Geral de IPTU", "Cálculo Geral de IPTU", 1, i, , , , "mnuCalcGeral"
      .AddItem "&Cálculo Geral de ISS", "Cálculo Geral de ISS", 1, i, , , , "mnuCalcGeralISS"
      .AddItem "&Geração de Arquivo Laser", "Geração de Arquivo Laser", 1, i, , , , "mnuArqLaser"
      i = .AddItem("&Isenção/Imunidade", "Isenção/Imunidade", 1, , , , , "mnuIsento")
      .OwnerDraw(i) = True
      i = .AddItem("Recálculo Ind&ividual", "Recálculo Ind&ividual", 1, , , , , "mnuRecalc")
      .OwnerDraw(i) = True
      i = .AddItem("-Débitos")
      .OwnerDraw(i) = True
      i = .AddItem("Optantes por Dé&bito Automático", "Optantes por Dé&bito Automático", 1, , 59, , , "mnuOptanteDA")
      .OwnerDraw(i) = True
      i = .AddItem("Importação de Ar&quivos Bancários", "Importação de Ar&quivos Bancários", 1, , , , , "mnuImportaArq")
      .OwnerDraw(i) = True
      i = .AddItem("Pa&gamento de Débito", "Pa&gamento de Débito", 1, , , , , "mnuPagDebito")
      .OwnerDraw(i) = True
      .AddItem "Pagam&ento Automático", "Pagam&ento Automático", 1, i, 55, , , "mnuPagAuto"
      .AddItem "Débito &Automático", "Débito &Automático", 1, i, 61, , , "mnuDebAutomatico"
      .AddItem "&Baixa Manual", "Baixa Manual", 1, i, 54, , , "mnuBaixaManual"
      .AddItem "&Analise da Receita", "Analise da Receita", 1, i, , , , "mnuAnaliseReceita"
      .AddItem "&Relatório de Pagamentos", "Relatorio de Pagamentos", 1, i, , , , "mnuRelPagamento"
      i = .AddItem("-Lançamentos")
      .OwnerDraw(i) = True
      i = .AddItem("Emissão de Guia Laser", "emissao guia laser", 1, , , , , "mnu2ViaLaser")
      .OwnerDraw(i) = True
      i = .AddItem("Movimento Econômico", "Movimento Econômico", 1, , , , , "mnuMovimento")
      .OwnerDraw(i) = True
      i = .AddItem("Parcelamento de Divida Fiscal", "Re&parcelamento de Débito", 1, , 60, , , "mnuReparcelDebito")
      .OwnerDraw(i) = True
      i = .AddItem("&Cancelamento de Parcelamento", "Cancelamento de Reparcelamento", 1, , 57, , , "mnuCancelReparc")
      .OwnerDraw(i) = True
      i = .AddItem("Cobrança de Aluguel", "Cobrança de Aluguel", 1, , , , , "mnuCobAluguel")
      .OwnerDraw(i) = True
      .AddItem "Manutenção dos Aluguéis", "Manutenção dos Aluguéis", 1, i, , , , "mnuManAluguel"
      .AddItem "Emissão dos Boletos de Cobrança", "Impressão dos Boletos", 1, i, , , , "mnuImpAluguel"
      i = .AddItem("Emissão de ITBI", "Emissão de ITBI", 1, , , , , "mnuITBI")
      .OwnerDraw(i) = True
      i = .AddItem("Divida Ativa", "Divida Ativa", 1, , , , , "mnuDivAtiv")
      .OwnerDraw(i) = True
      .AddItem "Encerramento do Livro", "Encerramento do Livro", 1, i, , , , "mnuDividaAtiva"
      .AddItem "Emissão dos Livros", "Emissão dos Livros", 1, i, , , , "mnuEmiteLivro"
      i = .AddItem("-Consulta")
      .OwnerDraw(i) = True
      i = .AddItem("&Consulta/Reativação de Documento", "Consulta Número de Documento", 1, , 62, , , "mnuCnsNumDoc")
      .OwnerDraw(i) = True
      i = .AddItem("&Consulta Digito Verif. de Documento", "Consulta Digito Verificador de Documento", 1, , , , , "mnuCnsNumDV")
      .OwnerDraw(i) = True
      i = .AddItem("Co&nsulta de Lançamentos", "Consulta Lançamento - Imobiliário", 1, , 58, , , "CnsDebitoImob")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MontaMenuOpcoes()

   Set m_cMenuOpcoes = New cPopupMenu
   With m_cMenuOpcoes
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("&Log do Sistema", "Log do Sistema", 1, , 9, , , "mnuLog")
      .OwnerDraw(i) = True
      i = .AddItem("P&referências do Sistema", "P&referências do Sistema", 1, , 15, , , "mnuConfig")
      .OwnerDraw(i) = True
      i = .AddItem("&Bloquear Sistema", "Bloquear Sistema", 1, , 11, , , "mnuLock")
      .OwnerDraw(i) = True
      i = .AddItem("Se&gurança", "Se&gurança", 1, , 10, , , "mnuSegurança")
      .OwnerDraw(i) = True
      .AddItem "Cadastro de &Usuários", "Cadastro de &Usuários", 1, i, 12, , , "mnuUser"
      .AddItem "Atribuição de Segurança &por Evento", "Atribuição de Segurança &por Evento", 1, i, 1, , , "mnuSegEvento"
      .AddItem "Atribuição de Segurança &por Usuário", "Atribuição de Segurança &por Usuário", 1, i, 6, , , "mnuAtribSeg"
            
   End With
   
End Sub

Private Sub MontaMenuRelatorio()

   Set m_cMenuRelatorio = New cPopupMenu
   With m_cMenuRelatorio
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = Trues
      i = .AddItem("-Tabelas")
      .OwnerDraw(i) = True
      i = .AddItem("&Tabelas Básicas", , 1, , 5, , , "mnuTBAS")
      .OwnerDraw(i) = True
      .AddItem "Lista de &Cidades", "Lista de Cidades", 2, i, 4, , , "mnuLCID"
      .AddItem "Lista de &Bairros", "Lista de Bairros", 2, i, 3, , , "mnuLBAI"
'      .AddItem "Lista de &Loteamentos", "Lista de Loteamentos", 1, i, 5, , , "mnuLLOT"
      .AddItem "Tabela de &Moedas", "Tabela de Moedas", 2, i, 34, , , "mnuTMOE"
      .AddItem "Tabela de &UFIR", "Tabela de UFIR", 2, i, 1, , , "mnuTUFI"
      .AddItem "Tabela de Taxas de &Expediente", "Tabela de Taxas de Expediente", 2, i, 44, , , "mnuTEXP"
      i = .AddItem("&Tabelas de Fatores Mobiliários", , 1, , 38, , , "mnuTFAT")
      .OwnerDraw(i) = True
      .AddItem "Fator &Benfeitoria", "Fator Benfeitoria", 2, i, 42, , , "mnuFBEN"
      .AddItem "Fator &Categoria", "Fator Categoria", 2, i, 43, , , "mnuFCAT"
      .AddItem "Fator &Distrito", "Fator Distrito", 2, i, 38, , , "mnuFDIS"
      .AddItem "Fator &Gleba", "Fator Gleba", 2, i, 41, , , "mnuFGLE"
      .AddItem "Fator &Pedologia", "Fator Pedologia", 2, i, 35, , , "mnuFPED"
      .AddItem "Fator P&rofundidade", "Fator Profundidade", 2, i, 39, , , "mnuFPRO"
      .AddItem "Fator &Situação", "Fator Situação", 2, i, 36, , , "mnuFSIT"
      .AddItem "Fator &Topografia", "Fator Topografia", 2, i, 37, , , "mnuFTOP"
      i = .AddItem("&Tabelas de Parâmetros Mobiliários", , 1, , 30, , , "mnuTPMO")
      .OwnerDraw(i) = True
      .AddItem "&Benfeitorias", "Benfeitorias", 2, i, 28, , , "mnuBENF"
      .AddItem "Categoria da C&onstrução", "Categoria da Construção", 2, i, 29, , , "mnuCATC"
      .AddItem "Categoria da &Propriedade", "Categoria da Propriedade", 2, i, 30, , , "mnuCATP"
      .AddItem "P&arâmetros das Parcelas", "Parametros das Parcelas", 2, i, , , , "mnuPPAR"
      .AddItem "P&edologia", "Pedologia", 2, i, 35, , , "mnuPEDO"
      .AddItem "Planta &Genérica", "Planta Genérica", 2, i, , , , "mnuPGEN"
      .AddItem "T&ipo de Construção", "Tipo de Construção", 2, i, 31, , , "mnuTCON"
      .AddItem "Tipos de Logradouro", "Tipos de Logradouro", 2, i, , , , "mnuTPLO"
      .AddItem "Títulos de Logradouro", "Titulos de Logradouro", 2, i, , , , "mnuTTLO"
      .AddItem "Uso da Construção", "Uso da Construção", 2, i, 32, , , "mnuUSOC"
      .AddItem "Uso do Terreno", "Uso do Terreno", 2, i, 33, , , "mnuUSOT"
      i = .AddItem("-Documentos")
      .OwnerDraw(i) = True
      i = .AddItem("Certidões", , 1, , 30, , , "mnuCertidoes")
      .OwnerDraw(i) = True
      .AddItem "Certidão de Débito", "Certidao de Débito", 2, i, , , , "mnuCertDeb"
      .AddItem "Certidão de Isenção", "Certidao de Isenção", 2, i, , , , "mnuCertIse"
      .AddItem "Certidão de Valor Venal", "Certidao de Valor Venal", 2, i, , , , "mnuCertVaV"
      .AddItem "Certidão de Endereço", "Certidao de Endereço Atualizado", 2, i, , , , "mnuCertEnd"
      .AddItem "Certidão de Demolição", "Certidao de Demolição", 2, i, , , , "mnuCertDem"
      i = .AddItem("&Álvara de Funcionamento", , 1, , , , , "mnuALVA")
      .OwnerDraw(i) = True
      i = .AddItem("&Termo de Confissão de Divida", , 1, , , , , "mnuTermConf")
      .OwnerDraw(i) = True
      i = .AddItem("&Relatório de Ajuizamento", , 1, , , , , "mnuRelAjuiza")
      .OwnerDraw(i) = True
      i = .AddItem("&Termo de Notificação", , 1, , , , , "mnuNotif")
      .OwnerDraw(i) = True
      i = .AddItem("&Cobrança Amigavel", , 1, , , , , "mnuCobAmi")
      .OwnerDraw(i) = True
      i = .AddItem("-Outros")
      .OwnerDraw(i) = True
      i = .AddItem("&Relatórios Diversos", , 1, , , , , "mnuRDIV")
      .OwnerDraw(i) = True
      .AddItem "Atividades Taxa de Licença", "Atividades Taxa de Licença", 2, i, , , , "mnuRelAtivTL"
      .AddItem "Atividades de ISS Variável/Estimado", "Atividades de ISS", 2, i, , , , "mnuRelAtivISS"
      .AddItem "Atividades de ISS Fixo", "Atividades de ISS", 2, i, , , , "mnuRelAtivISSFixo"
      .AddItem "Relatório de Devedores de ISS Variável", "Devedor Variavel", 2, i, , , , "mnuRelDevedorVariavel"
      .AddItem "Carta de Cobrança Amigável (IPTU)", "Cobrança Amigavel", 2, i, , , , "mnuRelCartaCobrança"
      i = .AddItem("&Emissão de 2ª Via", , 1, , , , , "mnu2Via")
      .OwnerDraw(i) = True
            
   End With
   
End Sub

Private Sub MontaMenuAjuda()

   Set m_cMenuAjuda = New cPopupMenu
   With m_cMenuAjuda
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("&Conteúdo", "Conteúdo", , , 14, , , "mnuConteúdo")
      .OwnerDraw(i) = True
      i = .AddItem("Í&ndice", "Índice", , , , , , "mnuIndice")
      .OwnerDraw(i) = True
      i = .AddItem("&Localizar", "Localizar", , , , , , "mnuLocalizar")
      .OwnerDraw(i) = True
      i = .AddItem("-Créditos")
      .OwnerDraw(i) = True
      i = .AddItem("Sobre o Sistema &GTI", "Sobre o Sistema GTI", , , , , , "mnuSobre")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MontaMenuAvancado()

   Set m_cMenuAvancado = New cPopupMenu
   With m_cMenuAvancado
      ' Set up for cPopupMenu:
      Set .BackgroundPicture = Picture3.Picture
      .hwndOwner = Me.hwnd
      .ImageList = ilsIcons
      .HeaderStyle = ecnmHeaderCaptionBar
      .GradientHighlight = True
      
      i = .AddItem("&Sql Builder", "Sql Builder", 1, , 27, , , "mnuSql")
      .OwnerDraw(i) = True
      i = .AddItem("&Geração Manual de Débitos", "Geração Manual de Débitos", 1, , , , , "mnuGeraDebito")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Sql As String, RdoAux As rdoResultset, nTest As Integer
On Error Resume Next

If MsgBox("Deseja  Sair do Sistema ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   Cancel = 1
   frmLogin.show 1
   Exit Sub
End If


Set hHelp = Nothing
Set gtiObj = Nothing
Set m_cMenuCadastro = Nothing
Set m_cMenuTabela = Nothing
Set m_cMenuImobiliario = Nothing
Set m_cMenuMobiliario = Nothing
Set m_cMenuTributario = Nothing
Set m_cMenuOpcoes = Nothing
Set m_cMenuAjuda = Nothing
Set m_cMenuAvancado = Nothing
Set m_cMenuRelatorio = Nothing
Set DC = Nothing
Unload frmCnsParcela
'Unload frmLogo

End Sub

Private Sub mnuCloseAll_Click()
Dim x As Integer
INICIO:
For x = 0 To Forms.Count - 1
    If Forms(x).Name <> "frmMdi" Then
       Unload Forms(x)
       GoTo INICIO:
    End If
Next

End Sub


Public Sub AddWindow(sNOME As String, sCaption As String)
    
   With m_cMenuJanela
      ' Set up for cPopupMenu:
      
      i = .AddItem(sCaption, sCaption, , , , , , sNOME)
      .OwnerDraw(i) = True
      
   End With
  
    
End Sub

Public Sub RemoveWindow(sNOME As String)
    
   With m_cMenuJanela
      ' Set up for cPopupMenu:
     
       .RemoveItem (sNOME)
      
   End With
  
   
End Sub

Private Sub Sbar_PanelClick(ByVal Panel As MSComctlLib.Panel)
Dim K As Long, sData As String, sRet As String

If Panel.Index = 6 Then
    sRet = RetEventUserForm("frmDataBase")
    If InStr(1, sRet, "001", vbBinaryCompare) > 0 Then
        frmDataBase.show vbModeless
        frmDataBase.Mv.Day = Val(Mid$(frmMdi.Sbar.Panels(6).text, 12, 2))
        frmDataBase.Mv.Month = Val(Mid$(frmMdi.Sbar.Panels(6).text, 15, 2))
        frmDataBase.Mv.Year = Val(Right$(frmMdi.Sbar.Panels(6).text, 4))
        frmDataBase.lblDB.Caption = "Data Base: " & frmDataBase.Mv.Day & "/" & frmDataBase.Mv.Month & "/" & frmDataBase.Mv.Year
    Else
        MsgBox "Você não tem permissão para alterar a Data Base", vbCritical, "Atenção"
    End If
End If

End Sub

Private Sub myTips_CustomSelection(UserID As String, Category As String, Value As Variant)
' processing tips from both this form and all of it's MDI child forms
StatusBar1.Panels(1).text = UserID
End Sub

Private Sub myTips_DisplayTip(TipText As String)
' Display the menu tips on the statusbar

StatusBar1.Panels(1).text = TipText

End Sub

Private Sub Timer1_Timer()
Dim Sql As String, RdoAux As rdoResultset, sDataBase As String, sOldData As String
Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='DATABASE'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
       Sql = "INSERT PARAMETROS(NOMEPARAM,VALPARAM) VALUES('DATABASE'" & ",'" & CStr(Format(Now, "dd/mm/yyyy")) & "')"
       cn.Execute Sql, rdExecDirect
       sDataBase = CStr(Format(Now, "dd/mm/yyyy"))
    Else
       sDataBase = !VALPARAM
    End If
   .Close
End With
sOldData = Right$(frmMdi.Sbar.Panels(6).text, 10)
If sDataBase <> sOldData Then
   MsgBox "A Data Base foi atualizada para " & sDataBase, vbInformation, "ATENÇÃO !!!"
   frmMdi.Sbar.Panels(6).text = "Data Base: " & sDataBase
End If

End Sub
