VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCadImob 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Imobiliário"
   ClientHeight    =   5910
   ClientLeft      =   5715
   ClientTop       =   3630
   ClientWidth     =   10170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   120
      TabIndex        =   33
      ToolTipText     =   "Consulta Imóveis Cadastrados"
      Top             =   2610
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmCadImob.frx":0000
      PICN            =   "frmCadImob.frx":001C
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
      Left            =   120
      TabIndex        =   35
      ToolTipText     =   "Gravar o Registro"
      Top             =   1530
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadImob.frx":0176
      PICN            =   "frmCadImob.frx":0192
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
      Left            =   120
      TabIndex        =   36
      ToolTipText     =   "Cancelar Edição"
      Top             =   1890
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadImob.frx":0537
      PICN            =   "frmCadImob.frx":0553
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtInativo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "frmCadImob.frx":06AD
      Top             =   4620
      Visible         =   0   'False
      Width           =   1335
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Novo Registro"
      Top             =   1530
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
      MICON           =   "frmCadImob.frx":06BF
      PICN            =   "frmCadImob.frx":06DB
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
      Left            =   120
      TabIndex        =   30
      ToolTipText     =   "Editar Registro"
      Top             =   1890
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
      MICON           =   "frmCadImob.frx":0835
      PICN            =   "frmCadImob.frx":0851
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
      Left            =   120
      TabIndex        =   31
      ToolTipText     =   "Desativar este imóvel"
      Top             =   2250
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
      MICON           =   "frmCadImob.frx":09AB
      PICN            =   "frmCadImob.frx":09C7
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
      Left            =   90
      TabIndex        =   32
      ToolTipText     =   "Sair da Tela"
      Top             =   4140
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
      MICON           =   "frmCadImob.frx":0A69
      PICN            =   "frmCadImob.frx":0A85
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdFoto 
      Height          =   315
      Left            =   120
      TabIndex        =   34
      ToolTipText     =   "Foto do Imóvel"
      Top             =   2970
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Foto     "
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
      MICON           =   "frmCadImob.frx":0AF3
      PICN            =   "frmCadImob.frx":0B0F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frTit 
      BackColor       =   &H00C0C000&
      BorderStyle     =   0  'None
      Height          =   1185
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   10125
      Begin VB.CheckBox chkImune 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
         Caption         =   "Possui imunidade:"
         Height          =   195
         Left            =   8325
         TabIndex        =   121
         Top             =   810
         Width           =   1590
      End
      Begin VB.TextBox txtMat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8775
         MaxLength       =   8
         TabIndex        =   24
         Top             =   90
         Width           =   1170
      End
      Begin VB.OptionButton optM 
         BackColor       =   &H00C0C000&
         Caption         =   "T"
         Height          =   195
         Index           =   1
         Left            =   8325
         TabIndex        =   23
         Top             =   135
         Width           =   420
      End
      Begin VB.OptionButton optM 
         BackColor       =   &H00C0C000&
         Caption         =   "M"
         Height          =   195
         Index           =   0
         Left            =   7830
         TabIndex        =   22
         Top             =   135
         Value           =   -1  'True
         Width           =   420
      End
      Begin VB.TextBox txtLote 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6450
         MaxLength       =   5
         TabIndex        =   27
         Top             =   750
         Width           =   705
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5310
         MaxLength       =   4
         TabIndex        =   26
         Top             =   750
         Width           =   585
      End
      Begin VB.TextBox txtSeq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7710
         MaxLength       =   2
         TabIndex        =   21
         Top             =   750
         Width           =   435
      End
      Begin MSComctlLib.ImageList ImlTv 
         Left            =   2340
         Top             =   675
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
               Picture         =   "frmCadImob.frx":0EBD
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":1019
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":1175
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":12D5
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":1435
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":1591
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":16ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":1849
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCadImob.frx":1B65
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSetor 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   4020
         TabIndex        =   25
         Top             =   810
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da SubUnidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   6810
         TabIndex        =   19
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lblSubUnid 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   195
         Left            =   8520
         TabIndex        =   18
         Top             =   450
         Width           =   795
      End
      Begin VB.Label lblUnid 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Height          =   195
         Left            =   5820
         TabIndex        =   17
         Top             =   450
         Width           =   795
      End
      Begin VB.Label lblIC 
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
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   1470
         TabIndex        =   16
         Top             =   450
         Width           =   2805
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Condomínio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   2130
         TabIndex        =   15
         Top             =   90
         Width           =   1065
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº da Unidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   7
         Left            =   4470
         TabIndex        =   14
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label lblCodReduz 
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
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   855
         TabIndex        =   12
         Top             =   105
         Width           =   1080
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7695
         TabIndex        =   11
         Top             =   90
         Width           =   135
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "("
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3150
         TabIndex        =   10
         Top             =   90
         Width           =   135
      End
      Begin VB.Label lblCond 
         BackStyle       =   0  'Transparent
         Caption         =   "Não Selecionado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3285
         TabIndex        =   9
         Top             =   90
         Width           =   4410
      End
      Begin VB.Label lblDist 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   840
         TabIndex        =   8
         Top             =   810
         Width           =   2445
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Insc.Cadastral:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   7
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   6
         Top             =   90
         Width           =   675
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   810
         Width           =   675
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   3420
         TabIndex        =   4
         Top             =   810
         Width           =   525
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   4560
         TabIndex        =   3
         Top             =   810
         Width           =   705
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   5970
         TabIndex        =   2
         Top             =   810
         Width           =   435
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Seq:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   7230
         TabIndex        =   1
         Top             =   810
         Width           =   375
      End
   End
   Begin prjChameleon.chameleonButton cmdAtivar 
      Height          =   315
      Left            =   90
      TabIndex        =   120
      ToolTipText     =   "Ativar este imóvel"
      Top             =   3330
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Ativar"
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
      MICON           =   "frmCadImob.frx":1C51
      PICN            =   "frmCadImob.frx":1C6D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frTab 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   4390
      Index           =   5
      Left            =   1440
      TabIndex        =   114
      Top             =   1260
      Width           =   8715
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H00000080&
         Height          =   4245
         Left            =   90
         TabIndex        =   115
         Top             =   0
         Width           =   7215
         Begin VB.TextBox txtHist 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Height          =   795
            Left            =   60
            MultiLine       =   -1  'True
            TabIndex        =   116
            Top             =   3015
            Width           =   7095
         End
         Begin MSFlexGridLib.MSFlexGrid grdHist 
            Height          =   2835
            Left            =   60
            TabIndex        =   117
            Top             =   150
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   5001
            _Version        =   393216
            Rows            =   1
            Cols            =   4
            FixedCols       =   0
            BackColorSel    =   192
            ForeColorSel    =   16777215
            BackColorBkg    =   15658734
            FocusRect       =   0
            SelectionMode   =   1
            Appearance      =   0
            FormatString    =   "^Data                |^Seq     |<Histórico                                                            |<Usuário                   "
         End
         Begin prjChameleon.chameleonButton cmdEditHist 
            Height          =   315
            Left            =   90
            TabIndex        =   118
            ToolTipText     =   "Editar Histórico"
            Top             =   3855
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Editar Histórico"
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
            MICON           =   "frmCadImob.frx":1DC7
            PICN            =   "frmCadImob.frx":1DE3
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
   End
   Begin VB.Frame frTab 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   1
      Left            =   1440
      TabIndex        =   44
      Top             =   1230
      Width           =   8685
      Begin VB.Frame frLI 
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H00000080&
         Height          =   3300
         Left            =   780
         TabIndex        =   45
         Top             =   570
         Width           =   6615
         Begin VB.TextBox txtNomeLogLI 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   675
            Width           =   4830
         End
         Begin VB.TextBox txtCodLogrLI 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   51
            Top             =   315
            Width           =   855
         End
         Begin VB.TextBox txtNum 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            TabIndex        =   50
            Text            =   "0"
            Top             =   990
            Width           =   855
         End
         Begin VB.TextBox txtCompl 
            Appearance      =   0  'Flat
            Height          =   690
            Left            =   1440
            MaxLength       =   200
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   1350
            Width           =   4845
         End
         Begin VB.ComboBox cmbBairroImovel 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1440
            TabIndex        =   48
            Top             =   2130
            Width           =   2835
         End
         Begin VB.TextBox txtQuadras 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   47
            Top             =   2520
            Width           =   2835
         End
         Begin VB.TextBox txtLotes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1440
            MaxLength       =   25
            TabIndex        =   46
            Top             =   2880
            Width           =   2835
         End
         Begin VB.Label lblBairroImovel 
            Height          =   225
            Left            =   1470
            TabIndex        =   134
            Top             =   2220
            Width           =   4875
         End
         Begin VB.Label lblCEP 
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000080&
            Height          =   165
            Left            =   4050
            TabIndex        =   61
            Top             =   1050
            Width           =   1245
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CEP.........................:"
            Height          =   225
            Index           =   7
            Left            =   2430
            TabIndex        =   60
            Top             =   1050
            Width           =   1545
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Lograd.....:"
            Height          =   225
            Index           =   6
            Left            =   180
            TabIndex        =   59
            Top             =   690
            Width           =   1275
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Logradouro.:"
            Height          =   225
            Index           =   5
            Left            =   180
            TabIndex        =   58
            Top             =   330
            Width           =   1275
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Lotes.................:"
            Height          =   225
            Left            =   180
            TabIndex        =   57
            Top             =   2940
            Width           =   1305
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Quadra..............:"
            Height          =   225
            Left            =   180
            TabIndex        =   56
            Top             =   2580
            Width           =   1305
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro.................:"
            Height          =   225
            Left            =   180
            TabIndex        =   55
            Top             =   2205
            Width           =   1305
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento.....:"
            Height          =   225
            Left            =   180
            TabIndex        =   54
            Top             =   1395
            Width           =   1305
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Número..............:"
            Height          =   225
            Left            =   180
            TabIndex        =   53
            Top             =   1035
            Width           =   1305
         End
      End
   End
   Begin VB.Frame frTab 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   4390
      Index           =   3
      Left            =   1530
      TabIndex        =   86
      Top             =   1260
      Width           =   8685
      Begin VB.Frame frDT 
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H00000080&
         Height          =   3975
         Left            =   810
         TabIndex        =   87
         Top             =   270
         Width           =   6705
         Begin VB.CheckBox chkCIP 
            Caption         =   "Isento da CIP"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   4020
            TabIndex        =   135
            Top             =   3330
            Width           =   1395
         End
         Begin VB.CheckBox chkConjugado 
            Alignment       =   1  'Right Justify
            Caption         =   "Conjugado............:"
            Height          =   195
            Left            =   120
            TabIndex        =   95
            Top             =   2970
            Width           =   1725
         End
         Begin VB.Frame frTestada 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Testadas"
            ForeColor       =   &H00000080&
            Height          =   2655
            Left            =   4050
            TabIndex        =   97
            Top             =   330
            Width           =   2475
            Begin VB.TextBox txtFace 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   660
               TabIndex        =   99
               Text            =   "0"
               Top             =   1950
               Width           =   435
            End
            Begin VB.TextBox txtTestada 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   660
               TabIndex        =   98
               Text            =   "0,00"
               Top             =   2280
               Width           =   885
            End
            Begin MSFlexGridLib.MSFlexGrid grdTestada 
               Height          =   1455
               Left            =   90
               TabIndex        =   100
               Top             =   330
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   2566
               _Version        =   393216
               Rows            =   1
               FixedCols       =   0
               BackColor       =   12632256
               BackColorFixed  =   15658734
               BackColorBkg    =   15658734
               FocusRect       =   0
               SelectionMode   =   1
               BorderStyle     =   0
               Appearance      =   0
               FormatString    =   "^Face        |^Metros              "
            End
            Begin prjChameleon.chameleonButton cmdAddTestada 
               Height          =   285
               Left            =   1620
               TabIndex        =   101
               ToolTipText     =   "Adicionar Testada"
               Top             =   2100
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
               MICON           =   "frmCadImob.frx":1F3D
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
               Left            =   1950
               TabIndex        =   102
               ToolTipText     =   "Remover Testada"
               Top             =   2100
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   503
               BTYPE           =   14
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
               MCOL            =   13026246
               MPTR            =   1
               MICON           =   "frmCadImob.frx":1F59
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
               TabIndex        =   104
               Top             =   2280
               Width           =   495
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Face:"
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   103
               Top             =   1980
               Width           =   495
            End
         End
         Begin VB.TextBox txtFracaoIdeal 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000080&
            Height          =   285
            Left            =   1620
            TabIndex        =   96
            Text            =   "0,00"
            Top             =   3285
            Width           =   900
         End
         Begin VB.ComboBox cmbPedol 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   94
            Top             =   2550
            Width           =   1905
         End
         Begin VB.ComboBox cmbTopog 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   91
            Top             =   1470
            Width           =   1905
         End
         Begin VB.ComboBox cmbSit 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   2190
            Width           =   1905
         End
         Begin VB.ComboBox cmbCatProp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   1830
            Width           =   1905
         End
         Begin VB.TextBox txtAreaTerreno 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1620
            TabIndex        =   88
            Text            =   "0,00"
            Top             =   330
            Width           =   1185
         End
         Begin VB.ComboBox cmbUso 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   750
            Width           =   1905
         End
         Begin VB.ComboBox cmbBenf 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1620
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   1110
            Width           =   1905
         End
         Begin VB.Label Label29 
            BackStyle       =   0  'Transparent
            Caption         =   "Fração Ideal..........:"
            Height          =   225
            Left            =   120
            TabIndex        =   112
            Top             =   3330
            Width           =   1500
         End
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   "Pedologia..............:"
            Height          =   225
            Left            =   120
            TabIndex        =   111
            Top             =   2640
            Width           =   1425
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Topografia.............:"
            Height          =   225
            Left            =   120
            TabIndex        =   110
            Top             =   1542
            Width           =   1425
         End
         Begin VB.Label Label19 
            BackStyle       =   0  'Transparent
            Caption         =   "Situação................:"
            Height          =   225
            Left            =   120
            TabIndex        =   109
            Top             =   2274
            Width           =   1425
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "Área do Terreno....:"
            Height          =   225
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   1425
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Uso do Terreno.....:"
            Height          =   225
            Left            =   120
            TabIndex        =   107
            Top             =   810
            Width           =   1425
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Benfeitoria.............:"
            Height          =   225
            Left            =   120
            TabIndex        =   106
            Top             =   1176
            Width           =   1425
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Categ. Propriedade:"
            Height          =   225
            Left            =   120
            TabIndex        =   105
            Top             =   1908
            Width           =   1425
         End
      End
   End
   Begin VB.Frame frTab 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   0
      Left            =   1470
      TabIndex        =   37
      Top             =   1215
      Width           =   8685
      Begin VB.Frame frPP 
         BackColor       =   &H00EEEEEE&
         Height          =   3795
         Left            =   930
         TabIndex        =   38
         Top             =   330
         Width           =   6525
         Begin VB.CheckBox chkReside 
            Caption         =   "Proprietário reside no imóvel"
            Height          =   240
            Left            =   135
            TabIndex        =   123
            Top             =   3465
            Width           =   3390
         End
         Begin MSComctlLib.TreeView tvProp 
            Height          =   3105
            Left            =   120
            TabIndex        =   39
            Top             =   180
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   5477
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
            Left            =   5205
            TabIndex        =   40
            ToolTipText     =   "Adicionar Proprietário/Compromissário"
            Top             =   405
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
            MICON           =   "frmCadImob.frx":1F75
            PICN            =   "frmCadImob.frx":1F91
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
            Left            =   5205
            TabIndex        =   41
            ToolTipText     =   "Remover Proprietário/Compromissário"
            Top             =   765
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
            MICON           =   "frmCadImob.frx":20EB
            PICN            =   "frmCadImob.frx":2107
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdCadCid 
            Height          =   315
            Left            =   5205
            TabIndex        =   42
            ToolTipText     =   "Cadastrar um novo Cidadão"
            Top             =   1155
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Cadastrar"
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
            MICON           =   "frmCadImob.frx":2261
            PICN            =   "frmCadImob.frx":227D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdCnsCid 
            Height          =   315
            Left            =   5205
            TabIndex        =   43
            ToolTipText     =   "Consultar dados do Cidadão"
            Top             =   1545
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Consultar"
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
            MICON           =   "frmCadImob.frx":23D7
            PICN            =   "frmCadImob.frx":23F3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdHistProp 
            Height          =   495
            Left            =   5205
            TabIndex        =   119
            ToolTipText     =   "Enviar para histórico os dados do proprietário"
            Top             =   1935
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   873
            BTYPE           =   3
            TX              =   "Enviar p/Histórico"
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
            MICON           =   "frmCadImob.frx":254D
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
   End
   Begin vbalDTab6.vbalDTabControl TabMob 
      Height          =   4665
      Left            =   1470
      TabIndex        =   20
      Top             =   1230
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   8229
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15658734
   End
   Begin VB.Frame frTab 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   4395
      Index           =   2
      Left            =   1440
      TabIndex        =   62
      Top             =   1260
      Width           =   8715
      Begin VB.Frame frTipoEE 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Tipo de Endereço"
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   750
         TabIndex        =   82
         Top             =   300
         Width           =   6735
         Begin VB.OptionButton optTEnd 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "Endereço do Imóvel"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   330
            TabIndex        =   85
            Top             =   330
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton optTEnd 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "Endereço do Proprietário"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2280
            TabIndex        =   84
            Top             =   330
            Width           =   2145
         End
         Begin VB.OptionButton optTEnd 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "Endereço de Entrega"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   4560
            TabIndex        =   83
            Top             =   330
            Width           =   1905
         End
      End
      Begin VB.Frame frEE 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Endereço de Entrega"
         ForeColor       =   &H00000080&
         Height          =   3165
         Left            =   765
         TabIndex        =   63
         Top             =   1110
         Width           =   6735
         Begin VB.ListBox lstNomeLog 
            BackColor       =   &H00C0FFFF&
            Height          =   1620
            ItemData        =   "frmCadImob.frx":2569
            Left            =   1740
            List            =   "frmCadImob.frx":256B
            TabIndex        =   65
            Top             =   585
            Visible         =   0   'False
            Width           =   4860
         End
         Begin VB.TextBox txtCodLogr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            MaxLength       =   6
            TabIndex        =   72
            Top             =   270
            Width           =   975
         End
         Begin VB.ComboBox cmbBairro 
            Height          =   315
            Left            =   1740
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   2220
            Width           =   2865
         End
         Begin VB.ComboBox cmbCidade 
            Height          =   315
            Left            =   1740
            Sorted          =   -1  'True
            TabIndex        =   70
            Text            =   "cmbCidade"
            Top             =   1890
            Width           =   2865
         End
         Begin VB.TextBox txtComplImovel 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            TabIndex        =   69
            Top             =   1260
            Width           =   4845
         End
         Begin VB.TextBox txtNumImovel 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            TabIndex        =   68
            Top             =   930
            Width           =   975
         End
         Begin VB.TextBox txtNomeLogr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            TabIndex        =   67
            Top             =   600
            Width           =   4845
         End
         Begin VB.ComboBox cmbUF 
            Height          =   315
            ItemData        =   "frmCadImob.frx":256D
            Left            =   1740
            List            =   "frmCadImob.frx":256F
            Sorted          =   -1  'True
            TabIndex        =   66
            Top             =   1560
            Width           =   2865
         End
         Begin esMaskEdit.esMaskedEdit mskCEP 
            Height          =   285
            Left            =   1740
            TabIndex        =   64
            Top             =   2580
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            MouseIcon       =   "frmCadImob.frx":2571
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   0
            MaxLength       =   9
            Mask            =   "99999-999"
            SelText         =   ""
            Text            =   "_____-___"
            HideSelection   =   -1  'True
         End
         Begin prjChameleon.chameleonButton cmdAddBairro 
            Height          =   270
            Left            =   4635
            TabIndex        =   122
            ToolTipText     =   "Cadastrar um novo bairro"
            Top             =   2250
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   476
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
            MPTR            =   99
            MICON           =   "frmCadImob.frx":258D
            PICN            =   "frmCadImob.frx":28A7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   -1  'True
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "CEP.........................:"
            Height          =   225
            Index           =   4
            Left            =   180
            TabIndex        =   81
            Top             =   2640
            Width           =   1545
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "UF...........................:"
            Height          =   225
            Index           =   2
            Left            =   180
            TabIndex        =   80
            Top             =   1620
            Width           =   1455
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade.....................:"
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   79
            Top             =   1950
            Width           =   1515
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Complemento...........:"
            Height          =   225
            Index           =   3
            Left            =   180
            TabIndex        =   78
            Top             =   1320
            Width           =   1545
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Número....................:"
            Height          =   225
            Index           =   2
            Left            =   180
            TabIndex        =   77
            Top             =   990
            Width           =   1545
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Logradouro....:"
            Height          =   225
            Index           =   1
            Left            =   180
            TabIndex        =   76
            Top             =   660
            Width           =   1545
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro......................:"
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   75
            Top             =   2280
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Logradouro.......:"
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   74
            Top             =   330
            Width           =   1545
         End
         Begin VB.Label lblCid 
            Caption         =   "to aqui"
            Height          =   195
            Left            =   4860
            TabIndex        =   73
            Top             =   1755
            Visible         =   0   'False
            Width           =   825
         End
      End
   End
   Begin VB.Frame frTab 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   4390
      Index           =   4
      Left            =   1410
      TabIndex        =   113
      Top             =   1230
      Width           =   8565
      Begin VB.ComboBox cmbAnoIPTU 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7140
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   2790
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvArea 
         Height          =   2145
         Left            =   180
         TabIndex        =   124
         Top             =   540
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3784
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
            Text            =   "Id"
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
      Begin prjChameleon.chameleonButton cmdDelArea 
         Height          =   315
         Left            =   1470
         TabIndex        =   125
         ToolTipText     =   "Remover uma Área"
         Top             =   2790
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
         MICON           =   "frmCadImob.frx":2A01
         PICN            =   "frmCadImob.frx":2A1D
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
         Left            =   240
         TabIndex        =   126
         ToolTipText     =   "Adicionar uma Área"
         Top             =   2790
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
         MICON           =   "frmCadImob.frx":2B77
         PICN            =   "frmCadImob.frx":2B93
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
         Left            =   2700
         TabIndex        =   127
         ToolTipText     =   "Editar uma Área"
         Top             =   2790
         Width           =   1155
         _ExtentX        =   2037
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
         MICON           =   "frmCadImob.frx":2CED
         PICN            =   "frmCadImob.frx":2D09
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdHist 
         Height          =   315
         Left            =   3930
         TabIndex        =   128
         ToolTipText     =   "Enviar para histórico os dados da área"
         Top             =   2790
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Enviar p/Histórico"
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
         MICON           =   "frmCadImob.frx":2E63
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblValorIPTU 
         Caption         =   "R$ 0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   7140
         TabIndex        =   133
         Top             =   3150
         Width           =   1275
      End
      Begin VB.Label Label9 
         Caption         =   "Valor do IPTU em: "
         Height          =   255
         Left            =   5760
         TabIndex        =   131
         Top             =   2850
         Width           =   1305
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade de Edificações:"
         ForeColor       =   &H00000080&
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   130
         Top             =   3210
         Width           =   2055
      End
      Begin VB.Label lblQtdeEdif 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2340
         TabIndex        =   129
         Top             =   3210
         Width           =   345
      End
   End
   Begin VB.Label lblSkin 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Height          =   285
      Left            =   180
      TabIndex        =   13
      Top             =   30
      Width           =   1335
   End
   Begin VB.Image ImgSkin 
      Height          =   345
      Index           =   0
      Left            =   -330
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
End
Attribute VB_Name = "frmCadImob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tALTERACAOIMOVEL
    Quadra   As Integer
    Lote As Integer
    Seq As Integer
    AreaTerreno As Double
    FracaoIdeal As Double
    UsoTerreno As String
    Benfeitoria As String
    Topografia As String
    CategProp As String
    Situacao As String
    Pedologia As String
    Logradouro As String
    Numero As Integer
    Quadras As String
    Lotes As String
End Type

Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim Sql As String, bExec As Boolean, bResize As Boolean
Dim Evento As String, NodX As Object
Private m_bEditFromCode As Boolean
Dim i As Integer, xImovel As clsImovel
Dim sRet As String
Dim bEsp As Boolean, evEsp As Integer
Dim bDel As Boolean, evDel As Integer
Dim bNew As Boolean, evNew As Integer
Dim bEdit As Boolean, evEdit As Integer
Dim bHist As Boolean, evHist As Integer
Dim HistImovel As tALTERACAOIMOVEL
Dim aProprietario() As Long

Private Sub cmbAnoIPTU_Click()
ValorIPTU
End Sub

Private Sub ValorIPTU()
Dim Sql As String, RdoAux As rdoResultset
lblValorIPTU.Caption = "R$ 0,00"
If tvProp.Nodes.Count = 2 Then Exit Sub

Sql = "select * from laseriptu where ano=" & cmbAnoIPTU.Text & " and codreduzido=" & Val(lblCodReduz.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblValorIPTU.Caption = "R$ " & FormatNumber(!valortotalparc * !qtdeparc, 2)
    End If
   .Close
End With

End Sub

Private Sub cmbBairro_GotFocus()
Dim i  As Long
    If cmbCidade.ListIndex = -1 Then Exit Sub
    If cmbCidade.ListIndex > -1 Then
       i = SendMessage(cmbCidade.hwnd, CB_FINDSTRING, -1, ByVal cmbCidade.Text)
       If i <> CB_ERR Then
          PostMessage cmbCidade.hwnd, CB_SETCURSEL, i, 0
       End If
    End If
    cmbBairro.Clear
    cmbBairro.AddItem ""
    lblCid.Caption = cmbCidade.ListIndex
    nCid = cmbCidade.ItemData(cmbCidade.ListIndex)
    Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Do While Not .EOF
           If !DescBairro <> "" Then
                cmbBairro.AddItem !DescBairro
                cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
           End If
          .MoveNext
        Loop
       .Close
    End With
End Sub

Private Sub cmbBairro_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            m_bEditFromCode = True
        Case vbKeyBack
            m_bEditFromCode = True
    End Select

End Sub

Private Sub cmbBairro_Validate(Cancel As Boolean)
Dim strPartial As String, i As Long
With cmbBairro
    If .Text <> "" Then
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i = CB_ERR Then
           MsgBox "Bairro Inválido.", vbExclamation, "Atenção"
           Cancel = True
        Else
           Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE DESCBAIRRO='" & strPartial & "'"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           If RdoAux.RowCount = 0 Then
              MsgBox "Bairro Inválido.", vbExclamation, "Atenção"
              Cancel = True
           Else
              .Text = UCase$(.Text)
           End If
           RdoAux.Close
        End If
    End If
End With

End Sub

Private Sub cmbBairroImovel_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String
    If optTEnd(0).value = True Then
        optTEnd_Click (0)
    End If
    'Prevent processing as a result of changes from code
    If m_bEditFromCode Then
        m_bEditFromCode = False
        Exit Sub
    End If
    With cmbBairroImovel
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i <> CB_ERR Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                m_bEditFromCode = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            Else
                PostMessage cmbBairroImovel.hwnd, CB_SETCURSEL, i, 0
            End If
        End If
    End With
End Sub

Private Sub cmbBairroImovel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            m_bEditFromCode = True
        Case vbKeyBack
            m_bEditFromCode = True
    End Select

End Sub

Private Sub cmbBairroImovel_Validate(Cancel As Boolean)
Dim strPartial As String, i As Long
With cmbBairroImovel
    If .Text <> "" Then
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i = CB_ERR Then
           MsgBox "Bairro Inválido.", vbExclamation, "Atenção"
           Cancel = True
        Else
           Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE DESCBAIRRO='" & strPartial & "'"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           If RdoAux.RowCount = 0 Then
              MsgBox "Bairro Inválido.", vbExclamation, "Atenção"
              Cancel = True
           Else
              .Text = UCase$(.Text)
           End If
           RdoAux.Close
        End If
    End If
End With
End Sub

Private Sub cmbCidade_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String

    'Prevent processing as a result of changes from code
    If m_bEditFromCode Then
        m_bEditFromCode = False
        Exit Sub
    End If
    With cmbCidade
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i <> CB_ERR Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                m_bEditFromCode = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            Else
                PostMessage cmbCidade.hwnd, CB_SETCURSEL, i, 0
            End If
        End If
    End With

End Sub

Public Sub cmbCidade_Click()

'Carrega Bairro
If Not bExec Then Exit Sub
If cmbCidade.ListIndex = -1 Then Exit Sub
cmbBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
'    cmbBairro.AddItem ""
    Do While Not .EOF
        If !DescBairro <> "" Then
            cmbBairro.AddItem !DescBairro
            cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
        End If
        .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbCidade_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            m_bEditFromCode = True
        Case vbKeyBack
            m_bEditFromCode = True
    End Select

End Sub

Private Sub cmbCidade_Validate(Cancel As Boolean)
Dim strPartial As String, i As Long
With cmbCidade
    If .Text <> "" Then
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i = CB_ERR Then
           MsgBox "Cidade Inválida.", vbExclamation, "Atenção"
           Cancel = True
        Else
           Sql = "SELECT DESCCIDADE FROM CIDADE WHERE DESCCIDADE='" & Mask(strPartial) & "'"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           If RdoAux.RowCount = 0 Then
              MsgBox "Cidade Inválida.", vbExclamation, "Atenção"
              Cancel = True
           Else
              .Text = UCase$(.Text)
           End If
           RdoAux.Close
        End If
    End If
End With


End Sub

Private Sub cmbUF_Change()
    Dim i As Long, j As Long
    Dim strPartial As String, strTotal As String

    'Prevent processing as a result of changes from code
    If m_bEditFromCode Then
        m_bEditFromCode = False
        Exit Sub
    End If
    With cmbUF
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i <> CB_ERR Then
            strTotal = .List(i)
            j = Len(strTotal) - Len(strPartial)
            If j <> 0 Then
                m_bEditFromCode = True
                .SelText = Right$(strTotal, j)
                .SelStart = Len(strPartial)
                .SelLength = j
            Else
                PostMessage cmbUF.hwnd, CB_SETCURSEL, i, 0
            End If
        End If
    End With

End Sub

Private Sub cmbUF_Click()
On Error Resume Next

'Carrega Cidade
If Not bExec Then Exit Sub
cmbCidade.Clear
cmbBairro.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do While Not .EOF
       cmbCidade.AddItem !descCidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbUF_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            m_bEditFromCode = True
        Case vbKeyBack
            m_bEditFromCode = True
    End Select

End Sub

Private Sub cmbUF_Validate(Cancel As Boolean)
Dim strPartial As String, i As Long
With cmbUF
    If .Text <> "" Then
        strPartial = .Text
        i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
        If i = CB_ERR Then
           MsgBox "UF Inválida.", vbExclamation, "Atenção"
           Cancel = True
        Else
           Sql = "SELECT SIGLAUF,DESCUF FROM UF WHERE  SIGLAUF='" & Left$(strPartial, 2) & "' AND DESCUF='" & Right$(strPartial, Len(strPartial) - 3) & "'"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           If RdoAux.RowCount = 0 Then
              MsgBox "UF Inválida.", vbExclamation, "Atenção"
              Cancel = True
           Else
              .Text = UCase$(.Text)
              bExec = True
              cmbUF_Click
           End If
           RdoAux.Close
        End If
    End If
End With

End Sub

Private Sub cmdAddArea_Click()
On Error GoTo Erro

Set frm = frmAreas
frm.sForm = Me.Name

frm.sEvento = "Novo"
ReDim aClassObra(0)
frm.show 1

lblQtdeEdif.Caption = lvArea.ListItems.Count

Exit Sub
Erro:
MsgBox "Selecione a área que deseja alterar.", vbExclamation, "Atenção"
End Sub

Private Sub cmdAddBairro_Click()
If cmbUF.ListIndex = -1 Or cmbCidade.ListIndex = -1 Then
    MsgBox "Selecione UF e Cidade do bairro a ser cadastardo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Left(cmbUF.Text, 2) = "SP" And cmbCidade.ItemData(cmbCidade.ListIndex) = 413 Then
    MsgBox "Apenas bairros de fora podem ser cadastrados.", vbExclamation, "Atenção"
    Exit Sub
End If

Set frm = frmBairro
frm.FormCall = Me.Name
frm.SiglaUF = cmbUF.Text
frm.CodCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
frmBairro.show

End Sub

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
   MsgBox "Selecione na árvore Proprietário ou Proprietário Solidário.", vbExclamation, "Atenção"
   
End Sub

Private Sub cmdAddTestada_Click()

Dim Achou As Boolean

If Val(txtFace.Text) = 0 Then
   MsgBox "Digite a Face da Testada.", vbExclamation, "Atenção"
   txtFace.SetFocus
   Exit Sub
End If

If CDbl(txtTestada.Text) = 0 Then
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

grdTestada.AddItem Format(txtFace.Text, "00") & Chr(9) & Format(txtTestada, "0#.00")
txtFace.Text = Val(grdTestada.TextMatrix(grdTestada.Rows - 1, 0)) + 1
txtTestada.Text = "0,00"

End Sub

Private Sub cmdAlterar_Click()

If txtInativo.Visible = True Then
   MsgBox "Este imóvel esta inativo.", vbExclamation, "Atenção"
   Exit Sub
End If

If tvProp.Nodes.Count = 2 Then
   MsgBox "Selecione o Imóvel a ser alterado.", vbExclamation, "Atenção"
   Exit Sub
End If

Evento = "Alterar"
Eventos "INCLUIR"

LiberaEE

If bEsp Then
    LiberaProp
    LiberaLI
    LiberaAT
    LiberaDT
    LiberaTT
    LiberaDC
    LiberaHI
End If

End Sub

Private Sub cmdAtivar_Click()
If txtInativo.Visible = True Then
    If MsgBox("Deseja tornar este imóvel ATIVO ????", vbQuestion + vbYesNo, "CONFIRMAÇÃO") = vbYes Then
        Sql = "UPDATE CADIMOB SET INATIVO=0 WHERE CODREDUZIDO=" & Val(lblCodReduz.Caption)
        cn.Execute Sql, rdExecDirect
        txtInativo.Visible = False
    End If
End If

End Sub

Private Sub cmdCadCid_Click()

On Error GoTo Erro:
   Set frm2 = frmCidadao
   frm2.sForm = Me.Name
   frm2.sTipoCidadao = Left$(tvProp.SelectedItem.Key, 1)
   frmCidadao.show
   Exit Sub
   
Erro:
   MsgBox "Clique na Árvore para selecionar Proprietário ou Proprietário Solidário.", vbExclamation, "Atenção"

End Sub

Private Sub cmdCancel_Click()

If Evento <> "Alterar" Then
   Limpa
End If
TabMob.Tabs.Item(1).Selected = True
Evento = ""
Eventos "INICIAR"
End Sub

Private Sub cmdCnsCid_Click()

Dim n As Integer
On Error GoTo Erro:
   n = tvProp.SelectedItem.Parent.Index
   CodCidadao = Val(Mid(tvProp.SelectedItem.Key, 5, Len(tvProp.SelectedItem.Key) - 4))
   frmCidadao.show
   Exit Sub
Erro:
   MsgBox "Selecione o Proprietário/Proprietário Solidário que deseja Consultar.", vbExclamation, "Atenção"

End Sub

Private Sub cmdConsultar_Click()
CodImovel = ""
sForm = "CI"
frmCnsImovel.show
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


Private Sub cmdDelCid_Click()
Dim n As Integer, nc As Integer
On Error GoTo Erro:
   n = tvProp.SelectedItem.Parent.Index
   nc = tvProp.SelectedItem.Index
'   Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & Val(Right(tvProp.Nodes(nc).Key, 6)) & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'   Sql = Sql & "O Cidadão foi removido de proprietário/proprietário solidário do imóvel de inscrição:" & lblIC.Caption & "." & lblUnid.Caption & "." & lblSubUnid.Caption & "')"
   Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & Val(Right(tvProp.Nodes(nc).Key, 6)) & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
   Sql = Sql & "O Cidadão foi removido de proprietário/proprietário solidário do imóvel de inscrição:" & lblIC.Caption & "." & lblUnid.Caption & "." & lblSubUnid.Caption & "')"
   cn.Execute Sql, rdExecDirect
   tvProp.Nodes.Remove (nc)
   If frmCadImob.tvProp.Nodes("PROP").Children > 0 And Right$(frmCadImob.tvProp.Nodes("PROP").Child.Text, 9) <> "Principal" Then
        frmCadImob.tvProp.Nodes("PROP").Child.Text = frmCadImob.tvProp.Nodes("PROP").Child.Text & " - Principal"
   End If
   Exit Sub
   
Erro:
   MsgBox "Selecione o Proprietário/Proprietário Solidário que deseja Remover.", vbExclamation, "Atenção"

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
If lvArea.SelectedItem.SubItems(2) <> "" Then
    frm.dDataConstrucao = lvArea.SelectedItem.SubItems(2)
End If
frm.nAreaConstrucao = CDbl(Left(lvArea.SelectedItem.SubItems(1), Len(lvArea.SelectedItem.SubItems(1)) - 3))

frm.show 1

End Sub

Private Sub cmdEditHist_Click()

Set frm = frmEditHist
frm.sForm = Me.Name
frmEditHist.show 1
End Sub

Private Sub cmdExcluir_Click()

If txtInativo.Visible = False Then
    If MsgBox("Deseja tornar este imóvel INATIVO ????", vbQuestion + vbYesNo, "CONFIRMAÇÃO") = vbYes Then
        Sql = "UPDATE CADIMOB SET INATIVO=1 WHERE CODREDUZIDO=" & Val(lblCodReduz.Caption)
        cn.Execute Sql, rdExecDirect
        txtInativo.Visible = True
    End If
End If

End Sub

Private Sub cmdFoto_Click()

If Val(lblSetor.Caption) > 0 Then
    
    sFormFoto = "I"
    frmImageImovel.show
    frmImageImovel.ZOrder 0
Else
    MsgBox "selecione um imóvel.", vbInformation, "Atenção"
End If

End Sub

Private Sub cmdGravar_Click()
If bLocal Then
    Exit Sub
End If

If Valida() Then
   If MsgBox("Gravar os dados deste Imóvel ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
      If Evento = "Novo" Then
           lblCodReduz.Caption = NovoCodReduzido
      End If
      Grava
   Else
      Exit Sub
   End If
Else
   Exit Sub
End If

Evento = ""
Eventos "INICIAR"

End Sub

Private Function Valida() As Boolean
Dim nSomaArea As Double
Valida = True

If optTEnd(0).value = True Then
   optTEnd_Click (0)
End If

If Evento = "Novo" Then
    Sql = "SELECT * FROM CADIMOB WHERE "
    Sql = Sql & "DISTRITO=" & Val(Left$(lblDist.Caption, 2)) & " AND "
    Sql = Sql & "SETOR=" & Val(lblSetor.Caption) & " AND "
    Sql = Sql & "QUADRA=" & Val(txtQuadra.Text) & " AND "
    Sql = Sql & "LOTE=" & Val(txtLote.Text) & " AND "
    Sql = Sql & "UNIDADE=" & Val(lblUnid.Caption) & " AND "
    Sql = Sql & "SUBUNIDADE=" & Val(lblSubUnid.Caption) & " AND "
    Sql = Sql & "CODREDUZIDO<>" & Val(Left$(lblCodReduz.Caption, 7)) & " AND INATIVO=0"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Esta Inscrição já pertence ao imóvel " & Format(!CODREDUZIDO, "000000") & ", por favor verifique.", vbExclamation, "Inscrição Duplicada"
           .Close
            GoTo Falso
        End If
       .Close
    End With
End If

If Val(txtQuadra.Text) = 0 Then
    MsgBox "Número de Quadra invalida na inscriçao."
    GoTo Falso
End If
If Val(txtLote.Text) = 0 Then
    MsgBox "Número de Lote invalido na inscriçao."
    GoTo Falso
End If

If Val(txtSeq.Text) = 0 Then
    MsgBox "Face de Quadra invalida na inscriçao."
    GoTo Falso
End If

Sql = "SELECT CODAGRUPA FROM FACEQUADRA WHERE "
Sql = Sql & "CODDISTRITO=" & Val(Left$(lblDist.Caption, 2)) & " AND "
Sql = Sql & "CODSETOR=" & Val(lblSetor.Caption) & " AND "
Sql = Sql & "CODQUADRA=" & Val(txtQuadra.Text) & " AND "
Sql = Sql & "CODFACE=" & Val(txtSeq.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount = 0 Then
   MsgBox "No de Face nao existente para esta quadra.", vbExclamation, "Atenção"
   txtSeq.SetFocus
   RdoAux.Close
   GoTo Falso
End If

'Proprietário
If tvProp.Nodes("PROP").Children = 0 Then
   MsgBox "Selecione um Proprietário.", vbCritical, "Erro de Validação."
   GoTo Falso
End If
'Local do Imovel
'If Val(lblBairroImovel.Tag) = 0 Then
'   MsgBox "Selecione o Bairro do Imóvel.", vbCritical, "Erro de Validação."
'   GoTo Falso
'End If

If mskCEP.ClipText = "" Then
   MsgBox "Digite o CEP do Imóvel.", vbCritical, "Erro de Validação."
   GoTo Falso
End If

'Dados do Terreno
If Val(txtAreaTerreno.Text) = 0 Then
   MsgBox "Digite a Área do Terreno.", vbCritical, "Erro de Validação."
   GoTo Falso
End If
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
If cmbSit.ItemData(cmbSit.ListIndex) = 2 And grdTestada.Rows <= 2 Then
   MsgBox "O Lote de esquina deve ter mais de 1 testada .", vbCritical, "Erro de Validação."
   GoTo Falso
End If
If cmbSit.ItemData(cmbSit.ListIndex) = 6 And grdTestada.Rows <= 2 Then
   MsgBox "O Lote de Quadra Inteira deve ter mais de 1 testada .", vbCritical, "Erro de Validação."
   GoTo Falso
End If
'If (cmbSit.ItemData(cmbSit.ListIndex) = 3 Or cmbSit.ItemData(cmbSit.ListIndex) = 4) And grdTestada.Rows > 1 Then
'   MsgBox "O Lote Interno ou Encravado não pode ter testadas .", vbCritical, "Erro de Validação."
'   GoTo Falso
'End If

If optTEnd(2).value = True And cmbCidade.Text = "JABOTICABAL" And Val(txtCodLogr.Text) = 0 Then
   MsgBox "Selecione o logradouro do endereço de entrega", vbCritical, "Erro de Validação."
   GoTo Falso
End If



If grdTestada.Rows > 1 Then
    If Val(txtSeq.Text) <> Val(grdTestada.TextMatrix(1, 0)) Then
       MsgBox "A testada principal (primeira testada) deve ser igual ao nº da face do imóvel.", vbCritical, "Erro de Validação."
       GoTo Falso
    End If
End If

nSomaArea = 0
For x = 1 To lvArea.ListItems.Count
    nSomaArea = nSomaArea + CDbl(Left(lvArea.ListItems(x).SubItems(1), Len(lvArea.ListItems(x).SubItems(1)) - 3))
Next

If optTEnd(2).value = True And cmbBairro.ListIndex = -1 Then
    MsgBox "Selecione o bairro do endereço de entrega", vbExclamation, "Atenção"
    GoTo Falso
End If


Valida = True
Exit Function

Falso:
   Valida = False

End Function


Private Sub cmdHist_Click()

If lvArea.ListItems.Count = 0 Then Exit Sub

If MsgBox("Enviar os dados da área selecionada para o histórico ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

s = "Enviado área para o histórico:" & vbNewLine
For n = 1 To 9
    Select Case n
        Case 1, 2, 4, 6, 8, 9
            s = s & lvArea.ColumnHeaders(n + 1) & "=> " & lvArea.SelectedItem.SubItems(n) & ", "
    End Select
Next

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(lblCodReduz.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With

grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & Format(nSeq, "00") & Chr(9) & s & Chr(9) & NomeDeLogin

End Sub

Private Sub cmdHistProp_Click()
Dim n As Integer
Dim nSeq As Integer, s As String
On Error GoTo Erro

n = tvProp.SelectedItem.Parent.Index

If MsgBox("Enviar os dados do proprietário selecionado para o histórico ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

s = "Antigo proprietário: " & Val(Mid(tvProp.SelectedItem.Key, 5, Len(tvProp.SelectedItem.Key) - 4)) & " - " & tvProp.SelectedItem.Text

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(lblCodReduz.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With

Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST) VALUES("
Sql = Sql & Val(lblCodReduz.Caption) & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & Mask(s) & "')"
cn.Execute Sql, rdExecDirect

grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & Format(nSeq, "00") & Chr(9) & s

Exit Sub
Erro:
MsgBox "Selecione o proprietário.", vbExclamation, "Atenção"

End Sub

Private Sub cmdNovo_Click()

lblCodReduz.Caption = "0"
lblIC.Caption = ""
lblDist.Caption = "0"
lblSetor.Caption = "0"
txtQuadra.Text = "0"
txtLote.Text = "0"
txtSeq.Text = "0"
lblUnid.Caption = "0"
lblSubUnid.Caption = "0"
txtInativo.Text = ""
chkReside.value = vbChecked
'Proprietario e Área
Inicio:
For i = 1 To tvProp.Nodes.Count
    tvProp.Nodes.Remove (i)
    GoTo Inicio
Next
Buildtree

'Local do Imovel
txtCodLogrLI.Text = "0"
txtNomeLogLI.Text = ""
txtNum.Text = "0"
txtCompl.Text = ""
cmbBairroImovel.ListIndex = -1
txtQuadras.Text = ""
txtLotes.Text = ""
lblCep.Caption = ""
'Endereço de Entrega
optTEnd(0).value = True
txtCodLogr.Text = "0"
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
txtComplImovel.Text = ""
cmbUF.ListIndex = -1
cmbCidade.Clear
cmbBairro.Clear
LimpaMascara mskCEP
'Dados do Terreno
txtAreaTerreno.Text = "0,00"
cmbUso.ListIndex = -1
cmbBenf.ListIndex = -1
cmbTopog.ListIndex = -1
cmbCatProp.ListIndex = -1
cmbSit.ListIndex = -1
cmbPedol.ListIndex = -1
'LimpaMascara mskAgua
grdTestada.Rows = 1
txtTestada.Text = "0,00"
txtFace.Text = "0"
txtFracaoIdeal.Text = "0"
'Historico
txtHist.Text = ""
grdHist.Rows = 1
frmSelCond.show 1

If lblIC.Caption = "" Then Exit Sub

bExec = True
Evento = "Novo"
Eventos "INCLUIR"
CarregaCondominio
Sql = "SELECT CODLOGR From FACEQUADRA WHERE "
Sql = Sql & "CODDISTRITO = " & Val(Left$(lblDist.Caption, 2)) & " AND CODSETOR = " & Val(lblSetor.Caption) & "  AND  CODQUADRA = " & Val(txtQuadra.Text) & " And CODFACE = " & Val(txtSeq.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
If RdoAux.RowCount > 0 Then
     txtCodLogrLI.Text = Format(RdoAux!CodLogr, "0000")
End If
RdoAux.Close
txtCodLogrLI_LostFocus
LiberaCampos
'Proprietario e Área
InicioA:
Inicio2A:

If CodCond = 999 Then
   cmdAddCid_Click
End If

End Sub

Private Sub CarregaCondominio()
Dim nAptos As Integer, itmX As ListItem
If CodCond = 999 Then Exit Sub

With xImovel
    .CarregaCondominio CLng(CodCond)
    txtCodLogrLI.Text = Format$(.CodLogr, "0000")
    txtCodLogrLI_LostFocus
    txtNum.Text = .Li_Num
    lblCep.Caption = .Li_Cep
    txtCompl.Text = .Li_Compl
    If .Li_CodBairro <> 999 Then
       For x = 0 To cmbBairroImovel.ListCount - 1
           cmbBairroImovel.ListIndex = x
           If cmbBairroImovel.ItemData(cmbBairroImovel.ListIndex) = .Li_CodBairro Then
              Exit For
           End If
       Next
    Else
       cmbBairroImovel.ListIndex = -1
    End If
     txtQuadras.Text = .Li_Quadras
     txtLotes.Text = .Li_Lotes
     txtAreaTerreno.Text = FormatNumber(.Dt_AreaTerreno, 2)
     If Not IsNull(.Dt_CodUsoTerreno) Then
        For x = 0 To cmbUso.ListCount - 1
            cmbUso.ListIndex = x
            If cmbUso.ItemData(cmbUso.ListIndex) = .Dt_CodUsoTerreno Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodBenf) Then
        For x = 0 To cmbBenf.ListCount - 1
            cmbBenf.ListIndex = x
            If cmbBenf.ItemData(cmbBenf.ListIndex) = .Dt_CodBenf Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodTopog) Then
        For x = 0 To cmbTopog.ListCount - 1
            cmbTopog.ListIndex = x
            If cmbTopog.ItemData(cmbTopog.ListIndex) = .Dt_CodTopog Then
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodCategProp) Then
        For x = 0 To cmbCatProp.ListCount - 1
            cmbCatProp.ListIndex = x
            If cmbCatProp.ItemData(cmbCatProp.ListIndex) = .Dt_CodCategProp Then
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
     If Not IsNull(.DescPedologia) Then
        For x = 0 To cmbPedol.ListCount - 1
            cmbPedol.ListIndex = x
            If cmbPedol.ItemData(cmbPedol.ListIndex) = .Dt_CodPedol Then
               Exit For
            End If
        Next
     End If
     optTEnd_Click (0)
     txtFracaoIdeal.Text = FormatNumber(.FracaoIdeal, 2)
        
     Sql = "SELECT CD_SUBUNIDADES FROM CONDOMINIOUNIDADE WHERE CD_CODIGO=" & CodCond
     Sql = Sql & " AND CD_UNIDADE=" & Val(lblUnid.Caption)
     Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
     With RdoAux
        If .RowCount > 0 Then
            nAptos = !CD_SUBUNIDADES
        Else
            nAptos = 1
        End If
       .Close
     End With
     
'TESTADAS

    .CarregaTestadaCond CLng(CodCond)
    For x = 1 To .QtdeTestadaCond
        If Val(.TestadaCond(x, 1)) = Val(txtSeq.Text) Then
            grdTestada.AddItem Format(.TestadaCond(x, 1), "00") & Chr(9) & FormatNumber(.TestadaCond(x, 2) / nAptos, 2)
            Exit For
        End If
    Next
    For x = 1 To .QtdeTestadaCond
        If Val(.TestadaCond(x, 1)) <> Val(txtSeq.Text) Then
            grdTestada.AddItem Format(.TestadaCond(x, 1), "00") & Chr(9) & FormatNumber(.TestadaCond(x, 2) / nAptos, 2)
        End If
    Next

    'Proprietario
    .CarregaProprietarioCondominio CodCond
     For x = 1 To .QtdeProp
           If .prop(x, 3) = "P" Then
               If .prop(x, 4) = 0 Then
                 Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(.prop(x, 1), "000000"), .prop(x, 2), 1)
              Else
                 Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(.prop(x, 1), "000000"), .prop(x, 2) & " - Principal", 1)
              End If
              tvProp.Nodes("PROP" & Format(.prop(x, 1), "000000")).ForeColor = vbBlue
           Else
              Set NodX = tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(.prop(x, 1), "000000"), .prop(x, 2), 2)
              tvProp.Nodes("COMP" & Format(.prop(x, 1), "000000")).ForeColor = vbBlue
           End If
     Next
    For x = 1 To frmCadImob.tvProp.Nodes.Count
       frmCadImob.tvProp.Nodes(x).EnsureVisible
    Next
    tvProp.Refresh

End With

'Areas
Dim z As Long
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)
Sql = "SELECT CONDOMINIOAREA.SEQAREA,CONDOMINIOAREA.QTDEPAV,CONDOMINIOAREA.TIPOAREA,CONDOMINIOAREA.DATAAPROVA,CONDOMINIOAREA.AREACONSTR,CONDOMINIOAREA.NUMPROCESSO,CONDOMINIOAREA.DATAPROCESSO,"
Sql = Sql & "CONDOMINIOAREA.USOCONSTR,USOCONSTR.DESCUSOCONSTR,CONDOMINIOAREA.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
Sql = Sql & "CONDOMINIOAREA.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM CONDOMINIOAREA INNER JOIN USOCONSTR ON "
Sql = Sql & "CONDOMINIOAREA.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
Sql = Sql & "CONDOMINIOAREA.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
Sql = Sql & "CONDOMINIOAREA.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
Sql = Sql & "WHERE CODCONDOMINIO=" & CodCond
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblQtdeEdif.Caption = .RowCount
    Do Until .EOF
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

Private Function NovoCodReduzido() As String

Dim s As String
Dim nCod As Long            'Ultimo codigo da Tabela

Sql = "SELECT MAX(CODREDUZIDO) AS LASTCOD FROM CADIMOB WHERE CODREDUZIDO<40000"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux2.RowCount > 0 Then
   If IsNull(h) Then
      nCod = 1
   Else
      nCod = RdoAux2!LASTCOD + 1
   End If
Else
   nCod = 1
End If
RdoAux2.Close

NovoCodReduzido = Format(nCod, "0000000") & "-" & RetornaDVCodReduzido(nCod)

End Function

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()

If Val(CodImovel) > 0 Then
   Ocupado
   lblCodReduz.Caption = CodImovel
   Le
End If
bExec = True
frPP.Refresh
Liberado
bResize = True
End Sub

Private Sub Form_Load()
Dim c As cTab, x As Integer
With TabMob
    .ShowCloseButton = False
    Set c = .Tabs.Add("Tab1", , "Proprietário")
    c.Panel = frTab(0)
    Set c = .Tabs.Add("Tab2", , "Local do Imóvel")
    c.Panel = frTab(1)
    Set c = .Tabs.Add("Tab3", , "Endereço de Entrega")
    c.Panel = frTab(2)
    Set c = .Tabs.Add("Tab4", , "Dados do Terreno")
    c.Panel = frTab(3)
    Set c = .Tabs.Add("Tab5", , "Áreas")
    c.Panel = frTab(4)
    Set c = .Tabs.Add("Tab6", , "Histórico")
    c.Panel = frTab(5)
End With

For x = 2006 To Year(Now)
    cmbAnoIPTU.AddItem (x)
Next
cmbAnoIPTU.ListIndex = cmbAnoIPTU.ListCount - 1
sRet = RetEventUserForm(Me.Name)
FormHagana
ReDim aProprietario(0)
Centraliza Me
Set xImovel = New clsImovel
CarregaCombo
Buildtree
Eventos "INICIAR"

End Sub

Private Sub CarregaCombo()

bExec = False
Sql = "SELECT SIGLAUF,DESCUF FROM UF ORDER BY DESCUF; " & _
      "SELECT CODSITUACAO,DESCSITUACAO FROM SITUACAO WHERE CODSITUACAO<>999 ORDER BY DESCSITUACAO; " & _
      "SELECT CODBENFEITORIA,DESCBENFEITORIA FROM BENFEITORIA WHERE CODBENFEITORIA<>999 ORDER BY DESCBENFEITORIA; " & _
      "SELECT CODPEDOLOGIA,DESCPEDOLOGIA FROM PEDOLOGIA WHERE CODPEDOLOGIA<>999 ORDER BY DESCPEDOLOGIA; " & _
      "SELECT CODTOPOGRAFIA,DESCTOPOGRAFIA FROM TOPOGRAFIA WHERE CODTOPOGRAFIA<>999 ORDER BY DESCTOPOGRAFIA; " & _
      "SELECT CODUSOTERRENO,DESCUSOTERRENO FROM USOTERRENO WHERE CODUSOTERRENO<>999 ORDER BY DESCUSOTERRENO; " & _
      "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO INNER JOIN CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE INNER JOIN UF ON CIDADE.SIGLAUF = UF.SIGLAUF WHERE (UF.SIGLAUF = 'SP') AND (DESCCIDADE = 'JABOTICABAL') AND (CODBAIRRO <> 999) ORDER BY DESCBAIRRO; " & _
      "SELECT CODCATEGPROP,DESCCATEGPROP FROM CATEGPROP WHERE CODCATEGPROP<>999 ORDER BY DESCCATEGPROP"

Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbUF.AddItem !SiglaUF & "-" & !DESCUF
      .MoveNext
    Loop
   .MoreResults
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
   cmbBairroImovel.AddItem ""
    Do Until .EOF
       cmbBairroImovel.AddItem !DescBairro
       cmbBairroImovel.ItemData(cmbBairroImovel.NewIndex) = !CodBairro
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

Private Sub Buildtree()

With tvProp
   .ImageList = ImlTv
    Set NodX = .Nodes.Add(, , "PROP", "Proprietários", 1)
    Set NodX = .Nodes.Add(, , "COMP", "Proprietário Solidário", 1)
End With


'Geral
With tvProp
    For x = 1 To .Nodes.Count
       .Nodes(x).EnsureVisible
    Next
   .Nodes("PROP").Bold = True
   .Nodes("COMP").Bold = True
End With


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
CodImovel = ""

End Sub

Private Sub grdHist_RowColChange()
If grdHist.Rows = 1 Then Exit Sub
If grdHist.Row > 0 Then
    txtHist.Text = grdHist.TextMatrix(grdHist.Row, 2)
End If

End Sub

Private Sub lstNomeLog_DblClick()

If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   If txtNumImovel.Enabled = True Then txtNumImovel.SetFocus
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
       txtNumImovel.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogLI.SetFocus
End If

End Sub

Private Sub mskCEP_GotFocus()
mskCEP.SelStart = 0
mskCEP.SelLength = Len(mskCEP.Text)
End Sub

Private Sub optTEnd_Click(Index As Integer)
If Index = 2 Then
   If Evento <> "" Then
      TravaEndereco
   End If
Else
   LiberaEndereco
End If

bExec = True
If Index = 0 Then
    CarregaEndImovel
ElseIf Index = 1 Then
    CarregaEndCidadao
ElseIf Index = 2 Then
    CarregaEndEntrega
End If

End Sub

Private Sub LiberaEndereco()
   txtCodLogr.Enabled = False
   txtCodLogr.BackColor = Kde
   txtNomeLogr.Enabled = False
   txtNomeLogr.BackColor = Kde
   txtComplImovel.Enabled = False
   txtComplImovel.BackColor = Kde
   txtNumImovel.Enabled = False
   txtNumImovel.BackColor = Kde
   cmbUF.Enabled = False
   cmbUF.BackColor = Kde
   cmbBairro.Enabled = False
   cmbBairro.BackColor = Kde
   cmbCidade.Enabled = False
   cmbCidade.BackColor = Kde
   mskCEP.BackColor = Kde
   mskCEP.Enabled = True
   cmdAddBairro.Enabled = False
End Sub

Private Sub TravaEndereco()
   txtCodLogr.Enabled = True
   txtCodLogr.BackColor = Branco
   txtNomeLogr.Enabled = True
   txtNomeLogr.BackColor = Branco
   txtComplImovel.Enabled = True
   txtComplImovel.BackColor = Branco
   txtNumImovel.Enabled = True
   txtNumImovel.BackColor = Branco
   cmbUF.Enabled = True
   cmbUF.BackColor = Branco
   cmbBairro.Enabled = True
   cmbBairro.BackColor = Branco
   cmbCidade.Enabled = True
   cmbCidade.BackColor = Branco
   mskCEP.BackColor = Branco
   mskCEP.Enabled = True
   cmdAddBairro.Enabled = True
End Sub

Private Sub TabMob_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal Y As Single)
bExec = True
If theTab.Index = 3 Then
    If optTEnd(0).value = True Then
       optTEnd_Click (0)
    End If
End If
End Sub

Private Sub tvProp_Collapse(ByVal Node As MSComctlLib.Node)
For x = 1 To frmCadImob.tvProp.Nodes.Count
   frmCadImob.tvProp.Nodes(x).EnsureVisible
Next
End Sub

Private Sub tvProp_DblClick()
If Val(Right$(tvProp.SelectedItem.Key, 6)) > 0 Then
    CodCidadao = Val(Right$(tvProp.SelectedItem.Key, 6))
    frmCidadao.show
    frmCidadao.ZOrder 0
End If
End Sub

Private Sub tvProp_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
tvProp.ToolTipText = Right$(tvProp.SelectedItem.Key, 6)
End Sub

Private Sub txtAreaTerreno_GotFocus()
txtAreaTerreno.SelStart = 0
txtAreaTerreno.SelLength = Len(txtAreaTerreno)
End Sub

Private Sub txtAreaTerreno_KeyPress(KeyAscii As Integer)

Tweak txtAreaTerreno, KeyAscii, DecimalPositive

End Sub

Private Sub txtCodLogr_Change()

If Val(txtCodLogr.Text) = 0 And txtComplImovel.BackColor = Branco Then
   txtNomeLogr.Enabled = True
   txtNomeLogr.BackColor = Branco
   txtNomeLogr.Text = ""
Else
   txtNomeLogr.Text = ""
   txtNomeLogr.Enabled = False
   txtNomeLogr.BackColor = Kde
End If

End Sub

Private Sub txtCodLogr_GotFocus()
txtCodLogr.SelStart = 0
txtCodLogr.SelLength = Len(txtCodLogr)
End Sub

Private Sub txtCodLogr_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
   KeyAscii = 0
End If
End Sub

Private Sub txtCodLogr_LostFocus()

txtCodLogr.Text = Val(txtCodLogr.Text)
On Error Resume Next
If Val(txtCodLogr.Text) > 0 Then
   Sql = "SELECT ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   If RdoAux.RowCount = 0 Then
      MsgBox "Logradouro não Cadastrado.", vbCritical, "Atenção"
      txtNomeLogr.Text = ""
      txtCodLogr.SetFocus
   Else
      txtNomeLogr.Text = Trim$(SubNull(RdoAux!AbrevTipoLog)) & " " & Trim$(SubNull(RdoAux!AbrevTitLog)) & " " & RdoAux!NomeLogradouro
   End If
   RdoAux.Close
End If

End Sub

Private Sub txtCodLogrLI_Change()
If Val(txtCodLogrLI.Text) = 0 Then
   txtNomeLogLI.Text = ""
End If
End Sub

Private Sub txtCodLogrLI_GotFocus()
txtCodLogrLI.SelStart = 0
txtCodLogrLI.SelLength = Len(txtCodLogrLI)
End Sub

Private Sub txtCodLogrLI_KeyPress(KeyAscii As Integer)
Tweak txtCodLogrLI, KeyAscii, IntegerPositive
End Sub

Private Sub txtCodLogrLI_LostFocus()

If Val(txtCodLogrLI.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogrLI.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogLI.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          lblNomeLog.Caption = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtCodLogrLI.SetFocus
       End If
   End With
End If

End Sub

Private Sub txtCompl_GotFocus()
txtCompl.SelStart = 0
txtCompl.SelLength = Len(txtCompl)
End Sub

Private Sub txtComplImovel_GotFocus()
txtComplImovel.SelStart = 0
txtComplImovel.SelLength = Len(txtComplImovel)
End Sub

Private Sub txtFace_GotFocus()
txtFace.SelStart = 0
txtFace.SelLength = Len(txtFace)
End Sub

Private Sub txtFace_KeyPress(KeyAscii As Integer)
Tweak txtFace, KeyAscii, IntegerPositive
End Sub

Private Sub txtLote_KeyPress(KeyAscii As Integer)
Tweak txtLote, KeyAscii, IntegerPositive
End Sub

Private Sub txtMat_KeyPress(KeyAscii As Integer)
Tweak txtMat, KeyAscii, IntegerPositive
End Sub

Private Sub txtNomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLogr.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLogr) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Digite o nome do logradouro a ser pesquisado, sem especificar o tipo e o título.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLogr.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.Text = 0
End If

End Sub

Private Sub txtNum_GotFocus()
txtNum.SelStart = 0
txtNum.SelLength = Len(txtNum)
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
Tweak txtNum, KeyAscii, IntegerPositive
End Sub

Private Sub txtNum_LostFocus()
lblCep.Caption = ""
If Val(txtNum.Text) > 10000 Then
    MsgBox "Nº inválido.", vbExclamation, "Atenção"
    txtNum.SetFocus
    Exit Sub
End If
If Val(txtCodLogrLI.Text) > 0 Then
     lblCep.Caption = RetornaCEP(Val(txtCodLogrLI.Text), Val(txtNum.Text))
     lblBairroImovel.Caption = RetornaBairro(RetornaNumero(lblCep.Caption)).Nome
     lblBairroImovel.Tag = RetornaBairro(RetornaNumero(lblCep.Caption)).Codigo
Else
    lblCep.Caption = ""
     lblBairroImovel.Caption = ""
     lblBairroImovel.Tag = ""
End If
End Sub

Private Sub txtNumImovel_GotFocus()
txtNumImovel.SelStart = 0
txtNumImovel.SelLength = Len(txtNumImovel)

End Sub

Private Sub txtNumImovel_KeyPress(KeyAscii As Integer)
Tweak txtNumImovel, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumImovel_LostFocus()

LimpaMascara mskCEP
If Val(txtNumImovel.Text) > 10000 Then
    MsgBox "Nº inválido.", vbExclamation, "Atenção"
    txtNumImovel.SetFocus
    Exit Sub
End If
If Val(txtCodLogr.Text) > 0 Then
     mskCEP.Text = RetornaCEP(Val(txtCodLogr.Text), Val(txtNumImovel.Text))
Else
    LimpaMascara mskCEP
End If


End Sub


Private Sub txtQuadra_KeyPress(KeyAscii As Integer)
Tweak txtQuadra, KeyAscii, IntegerPositive
End Sub

Private Sub txtSeq_KeyPress(KeyAscii As Integer)
Tweak txtSeq, KeyAscii, IntegerPositive
End Sub

Private Sub txtTestada_GotFocus()
txtTestada.SelStart = 0
txtTestada.SelLength = Len(txtTestada)
End Sub

Private Sub txtTestada_KeyPress(KeyAscii As Integer)
Dim Achou As Boolean

If KeyAscii = vbKeyReturn Then
        If Val(txtFace.Text) = 0 Then
           MsgBox "Digite a Face da Testada.", vbExclamation, "Atenção"
           txtFace.SetFocus
           Exit Sub
        End If
        
        If CDbl(txtTestada.Text) = 0 Then
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
End If

Tweak txtTestada, KeyAscii, DecimalPositive

End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   'frmMdi.m_cMenuImobiliario.Enabled(8) = True
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdConsultar.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmdSair.Visible = True
   cmdFoto.Enabled = True
   cmdEditHist.Visible = False
   mskCEP.Enabled = False
   mskCEP.BackColor = Kde
   txtMat.Locked = True
   txtMat.BackColor = frTit.BackColor
   chkReside.Enabled = False
   chkCIP.Enabled = False
   TravaCampos
   LiberaEndereco
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdConsultar.Visible = False
   cmdFoto.Enabled = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   cmdSair.Visible = False
   cmdEditHist.Visible = True
   txtMat.Locked = False
   txtMat.BackColor = Branco
   chkReside.Enabled = True
   chkCIP.Enabled = True
End If

End Sub

Private Sub TravaCampos()
   'frTit.Enabled = True
   cmdAddCid.Enabled = False
   cmdDelCid.Enabled = False
   cmdCadCid.Enabled = False
   cmdCnsCid.Enabled = False
   txtCodLogr.Enabled = False
   txtCodLogr.BackColor = Kde
   txtNomeLogr.Enabled = False
   txtNomeLogr.BackColor = Kde
   txtComplImovel.Enabled = False
   txtComplImovel.BackColor = Kde
   txtNumImovel.Enabled = False
   txtNumImovel.BackColor = Kde
   txtQuadra.Locked = True
   txtQuadra.BackColor = frTit.BackColor
   txtLote.Locked = True
   txtLote.BackColor = frTit.BackColor
   txtSeq.Locked = True
   txtSeq.BackColor = frTit.BackColor
   cmbUF.Enabled = False
   cmbUF.BackColor = Kde
   cmdAddBairro.Enabled = False
   cmbBairro.Enabled = False
   cmbBairro.BackColor = Kde
   cmbCidade.Enabled = False
   cmbCidade.BackColor = Kde
   txtNum.Enabled = False
   txtNum.BackColor = Kde
   txtCompl.Enabled = False
   txtCompl.BackColor = Kde
   cmbBairroImovel.Enabled = False
   cmbBairroImovel.BackColor = Kde
   mskCEP.BackColor = Kde
   mskCEP.Enabled = True
   txtQuadras.Enabled = False
   txtQuadras.BackColor = Kde
   txtLotes.Enabled = False
   txtLotes.BackColor = Kde
   txtAreaTerreno.Enabled = False
   txtAreaTerreno.BackColor = Kde
   cmbUso.Enabled = False
   cmbUso.BackColor = Kde
   cmbBenf.Enabled = False
   cmbBenf.BackColor = Kde
   cmbTopog.Enabled = False
   cmbTopog.BackColor = Kde
   cmbCatProp.Enabled = False
   cmbCatProp.BackColor = Kde
   cmbSit.Enabled = False
   cmbSit.BackColor = Kde
   cmbPedol.Enabled = False
   cmbPedol.BackColor = Kde
   txtFracaoIdeal.Enabled = False
   txtFracaoIdeal.BackColor = Kde
   txtTestada.Enabled = False
   txtTestada.BackColor = Kde
   txtFace.Enabled = False
   txtFace.BackColor = Kde
   cmdHist.Enabled = False
   cmdHistProp.Enabled = False
   cmdAddTestada.Enabled = False
   cmdDelTestada.Enabled = False
   cmdEditArea.Enabled = False
   cmdDelArea.Enabled = False
   cmdAddArea.Enabled = False
   optTEnd(0).Enabled = False
   optTEnd(1).Enabled = False
   optTEnd(2).Enabled = False
   chkImune.Enabled = False
   chkCIP.Enabled = False
   chkConjugado.Enabled = False
   chkReside.Enabled = False
   chkCIP.Enabled = False

End Sub

Private Sub LiberaCampos()
If Evento <> "Novo" Then Exit Sub

txtNum.Enabled = True
txtNum.BackColor = Branco
txtCompl.Enabled = True
txtCompl.BackColor = Branco
cmbBairroImovel.Enabled = True
cmbBairroImovel.BackColor = Branco
'txtMat.Enabled = True
'txtMat.BackColor = Branco
txtQuadras.Enabled = True
txtQuadras.BackColor = Branco
txtLotes.Enabled = True
txtLotes.BackColor = Branco
txtAreaTerreno.Enabled = True
txtAreaTerreno.BackColor = Branco
cmbUso.Enabled = True
cmbUso.BackColor = Branco
cmbBenf.Enabled = True
cmbBenf.BackColor = Branco
cmbTopog.Enabled = True
cmbTopog.BackColor = Branco
cmbCatProp.Enabled = True
cmbCatProp.BackColor = Branco
cmbSit.Enabled = True
cmbSit.BackColor = Branco
cmbPedol.Enabled = True
cmbPedol.BackColor = Branco
If bHist Then
    cmdHist.Enabled = True
End If
'frTit.Enabled = False
cmdAddCid.Enabled = True
cmdDelCid.Enabled = True
cmdCadCid.Enabled = True
cmdCnsCid.Enabled = True
txtTestada.Enabled = True
txtTestada.BackColor = Branco
txtFracaoIdeal.Enabled = True
txtFracaoIdeal.BackColor = Branco
txtFace.Enabled = True
txtFace.BackColor = Branco
cmdAddTestada.Enabled = True
cmdDelTestada.Enabled = True
lvArea.Enabled = True
cmdAddArea.Enabled = True
cmdEditArea.Enabled = True
cmdDelArea.Enabled = True
optTEnd(0).Enabled = True
optTEnd(1).Enabled = True
optTEnd(2).Enabled = True
txtHist.Enabled = True
txtHist.BackColor = vbWhite
chkImune.Enabled = True
chkCIP.Enabled = True
chkConjugado.Enabled = True
chkReside.Enabled = True
chkCIP.Enabled = True
End Sub

Private Sub Grava()
Dim nCodReduz As Long, nSeq As Integer, x As Integer, nCodCidadao As Long, bFind As Boolean
Dim qd As New rdoQuery, sSeq As String, sData As String
Dim sTemp As String, Sql2 As String, aHist() As String, Y As Integer
Dim nBairro As Integer, nCidade As Integer, nEnd As Integer, sNomeLogr As String
Set qd.ActiveConnection = cn

nCodReduz = Val(Left$(lblCodReduz.Caption, 7))
If Evento = "Alterar" Then GoTo Alteração

GoSub GravaImovel
GoSub GravaEndEntrega
GoSub GravaProprietario
GoSub GravaTestada
GoSub GravaArea
GoSub GravaHistorico

Exit Sub

'*******GRAVA IMOVEL**********************************************
GravaImovel:

If optTEnd(0).value = True Then
    nEnd = 0
ElseIf optTEnd(1).value = True Then
    nEnd = 1
ElseIf optTEnd(2).value = True Then
    nEnd = 2
End If

Sql = "INSERT CADIMOB(CODREDUZIDO,DV,CODCONDOMINIO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,LI_NUM,LI_COMPL,"
Sql = Sql & "LI_UF,LI_CODCIDADE,LI_CODBAIRRO,LI_QUADRAS,LI_LOTES,DT_AREATERRENO,DT_CODUSOTERRENO,DT_CODBENF,DT_CODTOPOG,"
Sql = Sql & "DT_CODCATEGPROP,DT_CODSITUACAO,DT_CODPEDOL,DT_NUMAGUA,DT_FRACAOIDEAL,DC_QTDEEDIF,DC_QTDEPAV,EE_TIPOEND,TIPOMAT,"
Sql = Sql & "NUMMAT,DATAINCLUSAO,IMUNE,CONJUGADO,RESIDEIMOVEL,CIP) values("
Sql = Sql & nCodReduz & "," & Val(Right$(lblCodReduz.Caption, 1)) & "," & IIf(Left$(lblCond.Caption, 1) = "N", 999, Val(Left$(lblCond.Caption, 4))) & ","
Sql = Sql & Val(lblDist.Caption) & "," & Val(lblSetor.Caption) & "," & Val(txtQuadra.Text) & "," & Val(txtLote.Text) & ","
Sql = Sql & Val(txtSeq.Text) & "," & Val(lblUnid.Caption) & "," & Val(lblSubUnid.Caption) & "," & Val(txtNum.Text) & ",'"
'Sql = Sql & Mask(txtCompl.Text) & "','" & "SP" & "'," & 413 & "," & IIf(Val(lblBairroImovel.Tag) > 0, Val(lblBairroImovel.Tag), "Null") & ",'"
Sql = Sql & Mask(txtCompl.Text) & "','" & "SP" & "'," & 413 & "," & IIf(cmbBairroImovel.ListIndex > -1, cmbBairroImovel.ItemData(cmbBairroImovel.ListIndex), "Null") & ",'"
Sql = Sql & Mask(txtQuadras.Text) & "','" & Mask(txtLotes.Text) & "'," & Virg2Ponto(RemovePonto(txtAreaTerreno.Text)) & "," & IIf(cmbUso.ListIndex > -1, cmbUso.ItemData(cmbUso.ListIndex), "Null") & ","
Sql = Sql & IIf(cmbBenf.ListIndex > -1, cmbBenf.ItemData(cmbBenf.ListIndex), "Null") & "," & IIf(cmbTopog.ListIndex > -1, cmbTopog.ItemData(cmbTopog.ListIndex), "Null") & ","
Sql = Sql & IIf(cmbCatProp.ListIndex > -1, cmbCatProp.ItemData(cmbCatProp.ListIndex), "Null") & "," & IIf(cmbSit.ListIndex > -1, cmbSit.ItemData(cmbSit.ListIndex), "Null") & ","
Sql = Sql & IIf(cmbPedol.ListIndex > -1, cmbPedol.ItemData(cmbPedol.ListIndex), "Null") & "," & "Null" & "," & Virg2Ponto(txtFracaoIdeal.Text) & "," & Val(lblQtdeEdif.Caption) & ",0," & nEnd & ",'"
Sql = Sql & IIf(optM(0).value = True, "M", "T") & "'," & Val(txtMat.Text) & ",'" & Format(Now, "mm/dd/yyyy") & "'," & IIf(chkImune.value = vbChecked, 1, 0) & "," & IIf(chkConjugado.value = vbChecked, 1, 0) & "," & IIf(chkReside.value = vbChecked, 1, 0) & "," & IIf(chkCIP.value = vbChecked, 1, 0) & ")"
cn.Execute Sql, rdExecDirect

Return
'*********************************************************************
'*******GRAVA ENDENTREGA *******************************************
GravaEndEntrega:
If cmbBairro.ListIndex > -1 Then
    nBairro = cmbBairro.ItemData(cmbBairro.ListIndex)
Else
    nBairro = 0
End If
If cmbCidade.ListIndex > -1 Then
    nCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
Else
    nCidade = 0
End If
If optTEnd(2).value = True Then
    If Val(txtCodLogr.Text) > 0 Then
        Sql = "SELECT NOMELOGRADOURO FROM LOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
        sNomeLogr = RdoAux!NomeLogradouro
        RdoAux.Close
    Else
        sNomeLogr = Mask(txtNomeLogr.Text)
    End If
    Sql = "INSERT ENDENTREGA(CODREDUZIDO,EE_CODLOG,EE_NOMELOG,EE_NUMIMOVEL,EE_COMPLEMENTO,EE_UF,EE_CIDADE,"
    Sql = Sql & "EE_BAIRRO,Ee_Cep) VALUES(" & nCodReduz & "," & Val(txtCodLogr.Text) & ",'" & Mask(sNomeLogr) & "',"
    Sql = Sql & Val(txtNumImovel.Text) & ",'" & Mask(txtComplImovel.Text) & "','" & Left$(cmbUF.Text, 2) & "',"
    Sql = Sql & IIf(nCidade > 0, nCidade, "Null") & "," & IIf(nBairro > 0, nBairro, "Null") & ",'" & mskCEP.Text & "')"
    cn.Execute Sql, rdExecDirect
End If
Return
'*********************************************************************
'*******GRAVA PROPRIETARIO *******************************************
GravaProprietario:
For x = 1 To tvProp.Nodes.Count
    If Len(tvProp.Nodes(x).Key) > 4 Then
        Sql = "INSERT PROPRIETARIO (CODREDUZIDO,CODCIDADAO,TIPOPROP,PRINCIPAL) VALUES("
        Sql = Sql & nCodReduz & "," & Val(Right$(tvProp.Nodes(x).Key, 6)) & ",'"
        Sql = Sql & Left$(tvProp.Nodes(x).Key, 1) & "'," & IIf(tvProp.Nodes("PROP").Child.Text = tvProp.Nodes(x).Text, 1, 0) & ")"
        cn.Execute Sql, rdExecDirect
        If Left$(tvProp.Nodes(x).Key, 1) = "P" And (tvProp.Nodes("PROP").Child.Text = tvProp.Nodes(x).Text) Then
            AtualizaPropDuplicado nCodReduz, Val(Right$(tvProp.Nodes(x).Key, 6))
        End If
    End If
Next
Return

'*********************************************************************
'*******GRAVA TESTADA *******************************************
GravaTestada:
For x = 1 To grdTestada.Rows - 1
    Sql = "INSERT TESTADA(CODREDUZIDO,NUMFACE,AREATESTADA) VALUES("
    Sql = Sql & nCodReduz & "," & Val(grdTestada.TextMatrix(x, 0)) & "," & Virg2Ponto(grdTestada.TextMatrix(x, 1)) & ")"
    cn.Execute Sql, rdExecDirect
Next
Return
'*********************************************************************
'*******GRAVA AREA *******************************************
GravaArea:

For x = 1 To lvArea.ListItems.Count
    Sql = "INSERT AREAS (CODREDUZIDO,SEQAREA,TIPOAREA,DATAAPROVA,AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,"
    Sql = Sql & "QTDEPAV) VALUES(" & nCodReduz & "," & x & ",'" & "'," & IIf(IsDate(lvArea.ListItems(x).SubItems(2)), "'" & Format(lvArea.ListItems(x).SubItems(2), "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & Virg2Ponto(Left(RemovePonto(lvArea.ListItems(x).SubItems(1)), Len(lvArea.ListItems(x).SubItems(1)) - 3)) & "," & lvArea.ListItems(x).SubItems(3) & "," & lvArea.ListItems(x).SubItems(5) & ","
    Sql = Sql & lvArea.ListItems(x).SubItems(7) & "," & lvArea.ListItems(x).SubItems(9) & ")"
    cn.Execute Sql, rdExecDirect
Next

Return

'*******HISTORICO DO IMÓVEL *******************************************
GravaHistorico:
Sql = "DELETE FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
cn.Execute Sql, rdExecDirect
With grdHist
    For x = 1 To .Rows - 1
'        Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'        Sql = Sql & nCodReduz & "," & x & ",'" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "','" & Mask(.TextMatrix(x, 2)) & "','" & Mask(.TextMatrix(x, 3)) & "','" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "')"
        Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
        Sql = Sql & nCodReduz & "," & x & ",'" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "','" & Mask(.TextMatrix(x, 2)) & "'," & RetornaUsuarioID(.TextMatrix(x, 3)) & ",'" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With

Return

'*********************************************************************
'*******ALTERAÇÃO DE IMÓVEL *******************************************
Alteração:

'caso o imovel for criado no mesmo dia da alteração não gravar historico de alteração
Sql = "SELECT CODREDUZIDO,DATAINCLUSAO FROM CADIMOB WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Not IsNull(RdoAux!DATAINCLUSAO) Then
    If Format(RdoAux!DATAINCLUSAO, "dd/mm/yyyy") = Format(Now, "dd/mm/yyyy") Then
        RdoAux.Close
        GoTo AfterHist
    End If
End If

If Evento = "Alterar" Then
    '*** HISTORICO ALTERAÇÃO ***
    ReDim aHist(0)
    If txtNomeLogLI.Text <> HistImovel.Logradouro Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado Logradouro de " & HistImovel.Logradouro & " para " & txtNomeLogLI.Text
    End If
    If txtNum.Text <> HistImovel.Numero Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado número do imóvel de " & HistImovel.Numero & " para " & txtNum.Text
    End If
    If txtLote.Text <> HistImovel.Lote Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado nº do lote de " & HistImovel.Lote & " para " & txtLote.Text
    End If
    If txtQuadra.Text <> HistImovel.Quadra Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado nº de quadra de " & HistImovel.Quadra & " para " & txtQuadra.Text
    End If
    If txtSeq.Text <> HistImovel.Seq Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado nº da seq de lote de " & HistImovel.Seq & " para " & txtSeq.Text
    End If
    If txtLotes.Text <> HistImovel.Lotes Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado nº de lotes da planta de " & HistImovel.Lotes & " para " & txtLotes.Text
    End If
    If txtQuadras.Text <> HistImovel.Quadras Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado nº de quadras da planta de " & HistImovel.Quadras & " para " & txtQuadras.Text
    End If
    If txtAreaTerreno.Text <> HistImovel.AreaTerreno Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado área do terreno de " & FormatNumber(HistImovel.AreaTerreno, 2) & " para " & FormatNumber(RemovePonto(txtAreaTerreno.Text), 2)
    End If
    If txtFracaoIdeal.Text <> HistImovel.FracaoIdeal Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado fração ideal de " & HistImovel.FracaoIdeal & " para " & txtFracaoIdeal.Text
    End If
    If cmbUso.Text <> HistImovel.UsoTerreno Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado uso do terreno de " & HistImovel.UsoTerreno & " para " & cmbUso.Text
    End If
    If cmbBenf.Text <> HistImovel.Benfeitoria Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado benfeitoria do terreno de " & HistImovel.Benfeitoria & " para " & cmbBenf.Text
    End If
    If cmbCatProp.Text <> HistImovel.CategProp Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado categoria da propriedade de " & HistImovel.CategProp & " para " & cmbCatProp.Text
    End If
    If cmbPedol.Text <> HistImovel.Pedologia Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado pedologia de " & HistImovel.Pedologia & " para " & cmbPedol.Text
    End If
    If cmbTopog.Text <> HistImovel.Topografia Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado topografia de " & HistImovel.Topografia & " para " & cmbTopog.Text
    End If
    If cmbSit.Text <> HistImovel.Situacao Then
        ReDim Preserve aHist(UBound(aHist) + 1)
        aHist(UBound(aHist)) = "Alterado situação de " & HistImovel.Situacao & " para " & cmbSit.Text
    End If
    
    
    'proprietario
    For x = 1 To tvProp.Nodes.Count
        nCodCidadao = Val(Right$(tvProp.Nodes(x).Key, 6))
        bFind = False
        If nCodCidadao > 0 Then
            For Y = 1 To UBound(aProprietario)
                If nCodCidadao = aProprietario(Y) Then
                    bFind = True
                End If
            Next
            If Not bFind Then
                ReDim Preserve aHist(UBound(aHist) + 1)
                aHist(UBound(aHist)) = "Incluido o proprietário código: " & nCodCidadao
            End If
        End If
    Next
    
    For x = 1 To UBound(aProprietario)
        nCodCidadao = aProprietario(x)
        bFind = False
        If nCodCidadao > 0 Then
            For Y = 1 To tvProp.Nodes.Count
                If nCodCidadao = Val(Right$(tvProp.Nodes(Y).Key, 6)) Then
                    bFind = True
                End If
            Next
            If Not bFind Then
                ReDim Preserve aHist(UBound(aHist) + 1)
                aHist(UBound(aHist)) = "Removido o proprietário código: " & nCodCidadao
            End If
        End If
    Next
    
    
    
    If UBound(aHist) > 0 Then
        For x = 1 To UBound(aHist)
            grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & aHist(x) & " pelo usuário: " & NomeDeLogin & Chr(9) & "GTI" & Chr(9) & Format(Now, "dd/mm/yyyy")
        Next
    End If
    
    LoadHistImovel
End If
'*******************************

AfterHist:
'Select Case sItemEdit
'    Case "PC"
         Sql = "DELETE FROM PROPRIETARIO WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
         GoSub GravaProprietario
'    Case "LI"
         Sql = "UPDATE CADIMOB SET "
         Sql = Sql & "QUADRA=" & Val(txtQuadra.Text) & ","
         Sql = Sql & "LOTE=" & Val(txtLote.Text) & ","
         Sql = Sql & "SEQ=" & Val(txtSeq.Text) & ","
         Sql = Sql & "TIPOMAT='" & IIf(optM(0).value = True, "M", "T") & "',"
         Sql = Sql & "NUMMAT=" & Val(txtMat.Text) & ","
         Sql = Sql & "LI_NUM=" & Val(txtNum.Text) & ","
         Sql = Sql & "LI_COMPL='" & Mask(txtCompl.Text) & "',"
         Sql = Sql & "LI_UF='" & "SP" & "',"
         Sql = Sql & "LI_CODCIDADE=" & 413 & ","
         If cmbBairroImovel.ListIndex > -1 Then
            Sql = Sql & "LI_CODBAIRRO=" & cmbBairroImovel.ItemData(cmbBairroImovel.ListIndex) & ","
         Else
            Sql = Sql & "LI_CODBAIRRO=" & 999 & ","
         End If
         Sql = Sql & "LI_QUADRAS='" & Mask(txtQuadras.Text) & "',"
         Sql = Sql & "LI_LOTES='" & Mask(txtLotes.Text) & "',"
         Sql = Sql & "IMUNE=" & IIf(chkImune.value = vbChecked, 1, 0) & ","
         Sql = Sql & "CONJUGADO=" & IIf(chkConjugado.value = vbChecked, 1, 0) & ", "
         Sql = Sql & "CIP=" & IIf(chkCIP.value = vbChecked, 1, 0) & ", "
         Sql = Sql & "RESIDEIMOVEL=" & IIf(chkReside.value = vbChecked, 1, 0) & " "
         Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
 '   Case "EE"
         Sql = "DELETE FROM ENDENTREGA WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
         If optTEnd(2).value = True Then
            GoSub GravaEndEntrega
            Sql = "UPDATE CADIMOB SET EE_TIPOEND=2 "
            Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz
            cn.Execute Sql, rdExecDirect
         Else
            Sql = "UPDATE CADIMOB SET "
            If optTEnd(0).value = True Then
               Sql = Sql & "EE_TIPOEND=0"
            ElseIf optTEnd(1).value = True Then
               Sql = Sql & "EE_TIPOEND=1"
            End If
            Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz
            cn.Execute Sql, rdExecDirect
         End If
  '  Case "TT"
         Sql = "DELETE FROM TESTADA WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
         GoSub GravaTestada
  '  Case "AT"
         Sql = "UPDATE CADIMOB SET "
         Sql = Sql & "DT_AREATERRENO=" & Virg2Ponto(RemovePonto(txtAreaTerreno.Text))
         Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
  '  Case "DT"
         Sql = "UPDATE CADIMOB SET "
         If cmbUso.ListIndex > -1 Then
            Sql = Sql & "DT_CODUSOTERRENO=" & cmbUso.ItemData(cmbUso.ListIndex) & ","
         End If
         If cmbBenf.ListIndex > -1 Then
            Sql = Sql & "DT_CODBENF=" & cmbBenf.ItemData(cmbBenf.ListIndex) & ","
         End If
         If cmbTopog.ListIndex > -1 Then
            Sql = Sql & "DT_CODTOPOG=" & cmbTopog.ItemData(cmbTopog.ListIndex) & ","
         End If
         If cmbCatProp.ListIndex > -1 Then
            Sql = Sql & "DT_CODCATEGPROP=" & cmbCatProp.ItemData(cmbCatProp.ListIndex) & ","
         End If
         If cmbSit.ListIndex > -1 Then
            Sql = Sql & "DT_CODSITUACAO=" & cmbSit.ItemData(cmbSit.ListIndex) & ","
         End If
         If cmbPedol.ListIndex > -1 Then
            Sql = Sql & "DT_CODPEDOL=" & cmbPedol.ItemData(cmbPedol.ListIndex) & ","
         End If
         Sql = Sql & "DT_NUMAGUA='" & "" & "',"
         If txtFracaoIdeal.Text = "" Then txtFracaoIdeal.Text = "0"
         Sql = Sql & "DT_FRACAOIDEAL=" & Virg2Ponto(RemovePonto(txtFracaoIdeal.Text))
         Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
 '   Case "DC"
         Sql = "DELETE FROM AREAS WHERE CODREDUZIDO=" & nCodReduz
         cn.Execute Sql, rdExecDirect
         GoSub GravaArea
 '   Case "HI"
         GoSub GravaHistorico
'End Select
'*********************************************************************

End Sub

Private Sub Le()
Dim x As Integer, nCodReduz As Long, nSeq As Integer
Dim itmX As ListItem, z As Long
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)
Limpa
lblCodReduz.Caption = CodImovel
nCodReduz = Val(Left$(CodImovel, 7))

With xImovel
    If nCodReduz = 0 Then Exit Sub
   .CarregaImovel nCodReduz
    
    txtInativo.Visible = .Inativo
    chkImune.value = IIf(.Imune, vbChecked, vbUnchecked)
    chkConjugado.value = IIf(.Conjugado, vbChecked, vbUnchecked)
    chkReside.value = IIf(.ResideImovel, vbChecked, vbUnchecked)
    chkCIP.value = IIf(.IsentoCIP, vbChecked, vbUnchecked)
    lblIC.Caption = Left$(.Inscricao, 15)
    lblUnid.Caption = Format(.Unidade, "00")
    lblSubUnid.Caption = Format(.SubUnidade, "000")
    lblDist.Caption = Format(.Distrito, "00") & " - " & .DescDistrito
    lblSetor.Caption = Format(.Setor, "00")
    txtQuadra.Text = Format(.Quadra, "0000")
    txtLote.Text = Format(.Lote, "00000")
    txtSeq.Text = Format(.Seq, "0")
    txtMat.Text = .NumMat
    If .TipoMat = "M" Or .TipoMat = "" Then
        optM(0).value = True
    ElseIf .TipoMat = "T" Then
        optM(1).value = True
    End If
    txtCodLogrLI.Text = .CodLogr
    txtCodLogrLI_LostFocus
    txtNum.Text = .Li_Num
    lblCep.Caption = RetornaCEP(.CodLogr, .Li_Num)
  '  lblBairroImovel.Caption = RetornaBairro(RetornaNumero(lblCep.Caption)).Nome
  '  lblBairroImovel.Tag = RetornaBairro(RetornaNumero(lblCep.Caption)).Codigo
    
    txtCompl.Text = .Li_Compl
    If .Li_CodBairro <> 999 Then
        For x = 0 To cmbBairroImovel.ListCount - 1
            If cmbBairroImovel.ItemData(x) = .Li_CodBairro Then
               cmbBairroImovel.ListIndex = x
               Exit For
            End If
        Next
     Else
        cmbBairroImovel.ListIndex = -1
     End If
     txtQuadras.Text = .Li_Quadras
     txtLotes.Text = .Li_Lotes
     txtAreaTerreno.Text = FormatNumber(.Dt_AreaTerreno, 2)
     If Not IsNull(.Dt_CodUsoTerreno) Then
        For x = 0 To cmbUso.ListCount - 1
            If cmbUso.ItemData(x) = .Dt_CodUsoTerreno Then
               cmbUso.ListIndex = x
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodBenf) Then
        For x = 0 To cmbBenf.ListCount - 1
            If cmbBenf.ItemData(x) = .Dt_CodBenf Then
               cmbBenf.ListIndex = x
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodTopog) Then
        For x = 0 To cmbTopog.ListCount - 1
            If cmbTopog.ItemData(x) = .Dt_CodTopog Then
               cmbTopog.ListIndex = x
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodCategProp) Then
        For x = 0 To cmbCatProp.ListCount - 1
            If cmbCatProp.ItemData(x) = .Dt_CodCategProp Then
               cmbCatProp.ListIndex = x
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodSituacao) Then
        For x = 0 To cmbSit.ListCount - 1
            If cmbSit.ItemData(x) = .Dt_CodSituacao Then
               cmbSit.ListIndex = x
               Exit For
            End If
        Next
     End If
     If Not IsNull(.Dt_CodPedol) Then
        For x = 0 To cmbPedol.ListCount - 1
            If cmbPedol.ItemData(x) = .Dt_CodPedol Then
               cmbPedol.ListIndex = x
               Exit For
            End If
        Next
     End If
     
     
    '*******************************

    'Proprietario
    .CarregaProprietario
     ReDim aProprietario(0)
     For x = 1 To .QtdeProp
           ReDim Preserve aProprietario(UBound(aProprietario) + 1)
           aProprietario(UBound(aProprietario)) = .prop(x, 1)
     
           If .prop(x, 3) = "P" Then
               If .prop(x, 4) = 0 Then
                 Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(.prop(x, 1), "000000"), IIf(.prop(x, 5) = 1, .prop(x, 2) & " " & .prop(x, 6), .prop(x, 2)), 1)
              Else
                 Set NodX = tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(.prop(x, 1), "000000"), IIf(.prop(x, 5) = 1, .prop(x, 2) & " " & .prop(x, 6), .prop(x, 2)) & " - Principal", 1)
              End If
              tvProp.Nodes("PROP" & Format(.prop(x, 1), "000000")).ForeColor = vbBlue
           Else
              Set NodX = tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(.prop(x, 1), "000000"), IIf(.prop(x, 5) = 1, .prop(x, 2) & .prop(x, 6), .prop(x, 2)), 1)
              tvProp.Nodes("COMP" & Format(.prop(x, 1), "000000")).ForeColor = vbBlue
           End If
     Next
    For x = 1 To frmCadImob.tvProp.Nodes.Count
       frmCadImob.tvProp.Nodes(x).EnsureVisible
    Next
    tvProp.Refresh
    ValorIPTU
     txtFracaoIdeal.Text = FormatNumber(.Dt_FracaoIdeal, 2)
     bExec = False
     optTEnd(.Ee_TipoEnd).value = True
     bExec = True
     lblQtdeEdif.Caption = .Dc_QtdeEdif
     If optTEnd(0).value = True Then 'endereco imovel
        CarregaEndImovel
     ElseIf optTEnd(1).value = True Then 'endereco imovel
        CarregaEndCidadao
     ElseIf optTEnd(2).value = True Then 'endereco entrega
        CarregaEndEntrega
     End If

    'Condominio
    .CarregaNomeCondominio .Distrito, .Setor, .Quadra, .Lote, .Seq
    If .CodCondominio = 0 Then
        lblCond.Caption = .NomeCondominio
    Else
        lblCond.Caption = Format(.CodCondominio, "0000") & " - " & .NomeCondominio
    End If
    
   'testadas
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
    
    If grdTestada.Rows > 1 Then
       txtFace.Text = Val(grdTestada.TextMatrix(1, 0)) + 1
    End If
    
    'Histórico
    'Sql = "SELECT SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2 FROM HISTORICO WHERE "
    Sql = "SELECT historico.*,Usuario.NomeLogin FROM historico INNER JOIN usuario ON historico.userid = usuario.Id WHERE "
    Sql = Sql & "CODREDUZIDO=" & nCodReduz & " ORDER BY SEQ"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            grdHist.AddItem Format(!DATAHIST2, "dd/mm/yyyy") & Chr(9) & Format(!Seq, "00") & Chr(9) & !DESCHIST & Chr(9) & SubNull(!NomeLogin)
          .MoveNext
        Loop
      .Close
    End With
    If grdHist.Rows > 1 Then
        grdHist.Row = 1
        grdHist.ColSel = 3
        grdHist_RowColChange
    End If
    grdHist.Refresh
    
    'Areas
    Sql = "SELECT AREAS.SEQAREA,AREAS.QTDEPAV,AREAS.TIPOAREA,AREAS.DATAAPROVA,AREAS.AREACONSTR,AREAS.NUMPROCESSO,AREAS.DATAPROCESSO,"
    Sql = Sql & "AREAS.USOCONSTR,USOCONSTR.DESCUSOCONSTR,AREAS.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
    Sql = Sql & "AREAS.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM AREAS INNER JOIN USOCONSTR ON "
    Sql = Sql & "AREAS.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
    Sql = Sql & "AREAS.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
    Sql = Sql & "AREAS.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
    Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        lblQtdeEdif.Caption = .RowCount
        nSeq = 1
        Do Until .EOF
        
           '****ListView
            Set itmX = lvArea.ListItems.Add(, "A" & Format(nSeq, "00"), Format(nSeq, "00"))
            itmX.SubItems(1) = FormatNumber(!AREACONSTR, 2) & " m²"
            itmX.SubItems(2) = IIf(IsNull(!DATAAPROVA), "", Format(!DATAAPROVA, "dd/mm/yyyy"))
            itmX.SubItems(3) = !USOCONSTR
            itmX.SubItems(4) = !descusoconstr
            itmX.SubItems(5) = !TIPOCONSTR
            itmX.SubItems(6) = !DESCTIPOCONSTR
            itmX.SubItems(7) = !CATCONSTR
            itmX.SubItems(8) = !desccategconstr
            itmX.SubItems(9) = Val(SubNull(!QTDEPAV))
        
           nSeq = nSeq + 1
          .MoveNext
        Loop
       .Close
    End With
    
End With

LoadHistImovel
If lvArea.ListItems.Count > 0 Then
    lvArea.ListItems(1).Selected = True
End If

For x = 1 To tvProp.Nodes.Count
    tvProp.Nodes(x).EnsureVisible
Next

CodImovel = ""
End Sub

Private Sub CarregaEndImovel()

If Not bExec Then Exit Sub
txtCodLogr.Text = ""
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
txtComplImovel.Text = ""
cmbUF.ListIndex = -1
cmbCidade.ListIndex = -1
cmbBairro.ListIndex = -1
LimpaMascara mskCEP

txtCodLogr.Text = txtCodLogrLI.Text
txtNomeLogr.Text = txtNomeLogLI.Text
txtNumImovel.Text = txtNum.Text
txtComplImovel.Text = txtCompl.Text

bExec = False
For x = 0 To cmbUF.ListCount - 1
    cmbUF.ListIndex = x
    If Left$(cmbUF.Text, 2) = "SP" Then
       Exit For
    End If
Next
bExec = True
cmbUF_Click
bExec = False
For x = 0 To cmbCidade.ListCount - 1
    cmbCidade.ListIndex = x
    If cmbCidade.Text = "JABOTICABAL" Then
       Exit For
    End If
Next

If cmbBairro.ListCount = 0 Then
   bExec = True
   cmbCidade_Click
   bExec = False
End If
If cmbBairroImovel.ListIndex > -1 Then
    For x = 0 To cmbBairro.ListCount - 1
        cmbBairro.ListIndex = x
        If cmbBairro.ItemData(x) = cmbBairroImovel.ItemData(cmbBairroImovel.ListIndex) Then
           Exit For
        End If
    Next
End If
bExec = True
mskCEP.Text = lblCep.Caption
frEE.Refresh
End Sub

Private Sub CarregaEndCidadao()
If Not bExec Then Exit Sub
Dim nCodigo As Long, sTipoEnd As String
txtCodLogr.Text = ""
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
txtComplImovel.Text = ""
cmbUF.ListIndex = -1
cmbCidade.ListIndex = -1
cmbBairro.ListIndex = -1
LimpaMascara mskCEP
For x = 1 To tvProp.Nodes.Count
     If Right$(tvProp.Nodes(x).Text, 9) = "Principal" Then
        nCodigo = Mid(tvProp.Nodes(x).Key, 5, Len(tvProp.Nodes(x).Key) - 4)
        Exit For
     End If
Next

If nCodigo = 0 Then
    MsgBox "Selecione o proprietário", vbExclamation, "Atenção"
    Exit Sub
End If

Sql = "SELECT CODCIDADAO,CODBAIRRO,CODBAIRRO2 FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Val(SubNull(RdoAux!CodBairro)) > 0 Then
   sTipoEnd = "R"
Else
   If Val(SubNull(RdoAux!CodBairro2)) > 0 Then
      sTipoEnd = "C"
   Else
      sTipoEnd = "R"
   End If
End If
RdoAux.Close

If sTipoEnd = "R" Then
    Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
    Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
    Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
    Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
Else
    Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
    Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
    Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
    Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodigo
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Val(SubNull(!FCodLogradouro)) > 0 Then
            txtCodLogr.Text = !FCodLogradouro
            txtCodLogr_LostFocus
        Else
            txtCodLogr.Text = "0"
            txtNomeLogr.Text = SubNull(!FNomeLogradouro)
        End If
            txtNumImovel.Text = Val(SubNull(!fNUMIMOVEL))
            If Not IsNull(!fsiglauf) Then
                bExec = False
                For x = 0 To cmbUF.ListCount - 1
                    If Left(cmbUF.List(x), 2) = !fsiglauf Then
                        cmbUF.ListIndex = x
                        Exit For
                    End If
                Next
                bExec = True
                cmbUF_Click
                bExec = False
               
                If Not IsNull(!fCodCidade) Then
                   For x = 0 To cmbCidade.ListCount - 1
                       cmbCidade.ListIndex = x
                       If cmbCidade.ItemData(cmbCidade.ListIndex) = !fCodCidade Then
                          Exit For
                       End If
                   Next
                   bExec = True
                   cmbCidade_Click
                   If Val(SubNull(!fCodBairro)) <> 0 And Val(SubNull(!fCodBairro)) <> 999 Then
                       For x = 0 To cmbBairro.ListCount - 1
                           cmbBairro.ListIndex = x
                           If cmbBairro.ItemData(cmbBairro.ListIndex) = !fCodBairro Then
                               Exit For
                           End If
                       Next
                    End If
                    bExec = True
                Else
                    cmbBairro.ListIndex = -1
                End If
            End If
            If Not IsNull(!FCEP) Then
               mskCEP.Text = Format(!FCEP, "00000-000")
            End If

        
    End If
End With

'Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & nCodigo
'Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux2
'    If .RowCount > 0 Then
'        If Val(SubNull(!CodLogradouro)) > 0 Then
'           txtCodLogr.Text = !CodLogradouro
'           txtCodLogr_LostFocus
'        Else
'            txtCodLogr.Text = "0"
'            txtNomeLogr.Text = SubNull(!Endereco)
'        End If
'        txtNumImovel.Text = Val(SubNull(!NUMIMOVEL))
'        txtComplImovel.Text = SubNull(!Complemento)
'        If Not IsNull(!SiglaUF) Then
'           bExec = False
'           For x = 0 To cmbUF.ListCount - 1
'               cmbUF.ListIndex = x
'               If Left$(cmbUF.Text, 2) = !SiglaUF Then
'                  Exit For
'               End If
'           Next
'           bExec = True
'           cmbUF_Click
'           bExec = False
'           If Not IsNull(!CodCidade) Then
'              For x = 0 To cmbCidade.ListCount - 1
'                   cmbCidade.ListIndex = x
'                   If cmbCidade.ItemData(cmbCidade.ListIndex) = !CodCidade Then
'                      Exit For
'                   End If
'              Next
'              bExec = True
'              cmbCidade_Click
'              If Val(SubNull(!CodBairro)) <> 0 And Val(SubNull(!CodBairro)) <> 999 Then
'                  For x = 0 To cmbBairro.ListCount - 1
'                       cmbBairro.ListIndex = x
'                       If cmbBairro.ItemData(cmbBairro.ListIndex) = !CodBairro Then
'                          Exit For
'                       End If
'                  Next
'                  bExec = True
'               Else
'                cmbBairro.ListIndex = -1
'               End If
'           End If
'        End If
'        If Not IsNull(!Cep) Then
'           mskCEP.Text = Format(!Cep, "00000-000")
'        End If
'    End If
'End With
frEE.Refresh
End Sub

Private Sub CarregaEndEntrega()
If Not bExec Then Exit Sub
Dim nCodigo As Long, sCep As String
txtCodLogr.Text = ""
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
txtComplImovel.Text = ""
cmbUF.ListIndex = -1
cmbCidade.ListIndex = -1
cmbBairro.ListIndex = -1
LimpaMascara mskCEP
nCodigo = Val(Left$(lblCodReduz.Caption, 7))

With xImovel
    .CarregaImovel nCodigo
    txtCodLogr.Text = .Ee_CodLog
    If Val(txtCodLogr.Text) > 0 Then
        txtCodLogr_LostFocus
    Else
        txtNomeLogr.Text = .Ee_NomeLog
    End If
    txtNumImovel.Text = .Ee_NumImovel
    txtComplImovel.Text = .Ee_Complemento
    If Not IsNull(.Ee_Uf) Then
       bExec = False
         For x = 0 To cmbUF.ListCount - 1
             cmbUF.ListIndex = x
             If Left$(cmbUF.Text, 2) = .Ee_Uf Then
                bExec = True
                cmbUF_Click
                Exit For
             End If
         Next
    End If
    If Not IsNull(.Ee_Cidade) Then
       bExec = False
       For x = 0 To cmbCidade.ListCount - 1
            cmbCidade.ListIndex = x
            If cmbCidade.ItemData(cmbCidade.ListIndex) = .Ee_Cidade Then
               bExec = True
               cmbCidade_Click
               Exit For
            End If
       Next
    End If
    If cmbCidade.ListIndex = -1 Then
        bExec = True
        cmbUF.ListIndex = 25
        bExec = False
        For x = 0 To cmbCidade.ListCount - 1
             cmbCidade.ListIndex = x
             If cmbCidade.ItemData(cmbCidade.ListIndex) = 413 Then
                bExec = True
                cmbCidade_Click
                Exit For
             End If
        Next
    End If
    If .Ee_Bairro > 0 Then
       For x = 0 To cmbBairro.ListCount - 1
            cmbBairro.ListIndex = x
            If cmbBairro.ItemData(cmbBairro.ListIndex) = .Ee_Bairro Then
               Exit For
            End If
       Next
    Else
       cmbBairro.ListIndex = -1
    End If
    On Error Resume Next
    If Not IsNull(.Ee_Cep) Then
        sCep = .Ee_Cep
       mskCEP.Text = sCep
    End If
    
End With

frEE.Refresh
End Sub

Private Sub Limpa()
Dim z As Long
z = SendMessage(lvArea.hwnd, LVM_DELETEALLITEMS, 0, 0)

'Cabeçalho

lblIC.Caption = ""
lblCodReduz.Caption = "0"
lblDist.Caption = ""
lblSetor.Caption = ""
optM(0).value = True
txtMat.Text = ""
txtQuadra.Text = ""
txtLote.Text = ""
lblValorIPTU.Caption = "R$ 0,00"
txtSeq.Text = ""
lblCond.Caption = "Não Selecionado"
'Proprietario e Área
Inicio:
For i = 1 To tvProp.Nodes.Count
    tvProp.Nodes.Remove (i)
    GoTo Inicio
Next
Buildtree
'Local do Imovel
chkImune.value = vbUnchecked
chkConjugado.value = vbUnchecked

txtCodLogrLI.Text = "0"
txtNomeLogLI.Text = ""
txtNum.Text = "0"
txtCompl.Text = ""
cmbBairroImovel.ListIndex = -1
txtQuadras.Text = ""
txtLotes.Text = ""
lblCep.Caption = ""
'Endereço de Entrega
optTEnd(0).value = True
txtCodLogr.Text = "0"
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
txtComplImovel.Text = ""
cmbUF.ListIndex = -1
cmbCidade.Clear
cmbBairro.Clear
LimpaMascara mskCEP
'Dados do Terreno
txtAreaTerreno.Text = "0,00"
cmbUso.ListIndex = -1
cmbBenf.ListIndex = -1
cmbTopog.ListIndex = -1
cmbCatProp.ListIndex = -1
cmbSit.ListIndex = -1
cmbPedol.ListIndex = -1
'LimpaMascara mskAgua
lblQtdeEdif.Caption = "0"
grdTestada.Rows = 1
txtTestada.Text = "0,00"
txtFace.Text = "0"
txtFracaoIdeal.Text = "0"
'Historico
txtHist.Text = ""
grdHist.Rows = 1
End Sub

Public Sub AlteraCadastro()

Evento = "Alterar"
LiberaProp
LiberaLI
LiberaEE
LiberaAT
LiberaDT
LiberaTT
LiberaDC
Eventos "INCLUIR"

End Sub

Private Sub LiberaProp()
    cmdAddCid.Enabled = True
    cmdDelCid.Enabled = True
    cmdCadCid.Enabled = True
    cmdCnsCid.Enabled = True
    cmdHistProp.Enabled = True
End Sub

Private Sub LiberaLI()
    txtMat.Locked = False
    txtMat.BackColor = Branco
    txtQuadra.Locked = False
    txtQuadra.BackColor = Branco
    txtLote.Locked = False
    txtLote.BackColor = Branco
    txtSeq.Locked = False
    txtSeq.BackColor = Branco
    txtNum.Enabled = True
    txtNum.BackColor = Branco
    txtCompl.Enabled = True
    txtCompl.BackColor = Branco
    cmbBairroImovel.Enabled = True
    cmbBairroImovel.BackColor = Branco
    txtQuadras.Enabled = True
    txtQuadras.BackColor = Branco
    txtLotes.Enabled = True
    txtLotes.BackColor = Branco
    chkImune.Enabled = True
    chkConjugado.Enabled = True
End Sub

Private Sub LiberaEE()
   optTEnd(0).Enabled = True
   optTEnd(1).Enabled = True
   optTEnd(2).Enabled = True
   If optTEnd(2).value = True Then
      TravaEndereco
   End If
   
End Sub

Private Sub LiberaAT()
   txtAreaTerreno.Enabled = True
   txtAreaTerreno.BackColor = Branco
End Sub

Private Sub LiberaDT()
   cmbUso.Enabled = True
   cmbUso.BackColor = Branco
   cmbBenf.Enabled = True
   cmbBenf.BackColor = Branco
   cmbTopog.Enabled = True
   cmbTopog.BackColor = Branco
   cmbCatProp.Enabled = True
   cmbCatProp.BackColor = Branco
   cmbSit.Enabled = True
   cmbSit.BackColor = Branco
   cmbPedol.Enabled = True
   cmbPedol.BackColor = Branco
   txtFracaoIdeal.Enabled = True
   txtFracaoIdeal.BackColor = Branco

End Sub

Private Sub LiberaTT()
    txtFace.Enabled = True
    txtFace.BackColor = Branco
   txtTestada.Enabled = True
   txtTestada.BackColor = Branco
   cmdAddTestada.Enabled = True
   cmdDelTestada.Enabled = True
End Sub

Private Sub LiberaDC()
   lvArea.Enabled = True
   cmdAddArea.Enabled = True
   cmdEditArea.Enabled = True
   cmdDelArea.Enabled = True
   cmdHist.Enabled = True
End Sub

Private Sub LiberaHI()
   txtHist.Enabled = True
   txtHist.BackColor = vbWhite
End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evEsp = 11
evDel = 4
evHist = 21
bEsp = False: bDel = False: bNew = False: bEdit = False: bHist = False

cmdExcluir.Enabled = False: cmdEditHist.Enabled = False: cmdAtivar.Enabled = False

If InStr(1, sRet, Format(evEsp, "000"), vbBinaryCompare) > 0 Then bEsp = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True
If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evHist, "000"), vbBinaryCompare) > 0 Then bHist = True

cmdNovo.Enabled = bNew
cmdAlterar.Enabled = bEdit

If bDel Then
    cmdExcluir.Enabled = True
    cmdAtivar.Enabled = True
End If
If bHist Then cmdEditHist.Enabled = True

End Sub

Private Sub LoadHistImovel()

'*** ATUALIZA ORIGEM ALTERAÇÃO ***
HistImovel.Logradouro = txtNomeLogLI.Text
HistImovel.Numero = txtNum.Text
HistImovel.Lote = txtLote.Text
HistImovel.Quadra = txtQuadra.Text
HistImovel.Seq = txtSeq.Text
HistImovel.Lotes = txtLotes.Text
HistImovel.Quadras = txtQuadras.Text
HistImovel.AreaTerreno = txtAreaTerreno.Text
HistImovel.FracaoIdeal = txtFracaoIdeal.Text
HistImovel.UsoTerreno = cmbUso.Text
HistImovel.Benfeitoria = cmbBenf.Text
HistImovel.CategProp = cmbCatProp.Text
HistImovel.Pedologia = cmbPedol.Text
HistImovel.Topografia = cmbTopog.Text
HistImovel.Situacao = cmbSit.Text

End Sub
