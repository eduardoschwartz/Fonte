VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCidadao 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cidadão"
   ClientHeight    =   6300
   ClientLeft      =   7650
   ClientTop       =   1860
   ClientWidth     =   7605
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames jcFrames3 
      Height          =   1485
      Left            =   60
      Top             =   60
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   2619
      FrameColor      =   12829635
      Style           =   0
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtCod 
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
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   56
         Top             =   105
         Width           =   975
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   4
         Top             =   420
         Width           =   6300
      End
      Begin VB.TextBox txtOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4605
         MaxLength       =   25
         TabIndex        =   6
         Top             =   765
         Width           =   2790
      End
      Begin VB.OptionButton optP 
         Caption         =   "Pessoa Física"
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optP 
         Caption         =   "Pessoa Jurídica"
         Height          =   285
         Index           =   1
         Left            =   5970
         TabIndex        =   3
         Top             =   60
         Width           =   1455
      End
      Begin esMaskEdit.esMaskedEdit mskCPF 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   1110
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":0000
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
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         SelText         =   ""
         Text            =   "___.___.___-__"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskCNPJ 
         Height          =   285
         Left            =   4605
         TabIndex        =   8
         Top             =   1110
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":001C
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
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         SelText         =   ""
         Text            =   "__.___.___/____-__"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskRG 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   765
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":0038
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
         MaxLength       =   25
         SelText         =   ""
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmbFindCod 
         Height          =   270
         Left            =   2280
         TabIndex        =   0
         ToolTipText     =   "Consulta por código"
         Top             =   60
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   476
         BTYPE           =   9
         TX              =   ""
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCidadao.frx":0054
         PICN            =   "frmCidadao.frx":0070
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
         Caption         =   "Código.........:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   90
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome..........:"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   61
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ..:"
         Height          =   225
         Index           =   11
         Left            =   4020
         TabIndex        =   60
         Top             =   1125
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF.............:"
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   59
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RG..............:"
         Height          =   225
         Index           =   13
         Left            =   120
         TabIndex        =   58
         Top             =   780
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Orgão..:"
         Height          =   225
         Index           =   14
         Left            =   4020
         TabIndex        =   57
         Top             =   795
         Width           =   630
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   1695
      Index           =   0
      Left            =   60
      Top             =   1560
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   2990
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço Residencial"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.CheckBox chkEtiq 
         Alignment       =   1  'Right Justify
         Caption         =   "Etiqueta"
         Height          =   225
         Left            =   6480
         TabIndex        =   28
         Top             =   1350
         Width           =   885
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   18
         Top             =   240
         Width           =   5220
      End
      Begin VB.TextBox txtNumLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   20
         Top             =   585
         Width           =   930
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2970
         MaxLength       =   50
         TabIndex        =   22
         Top             =   585
         Width           =   4455
      End
      Begin VB.ComboBox cmbBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1305
         Width           =   2250
      End
      Begin VB.ComboBox cmbCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   930
         Width           =   3000
      End
      Begin VB.ComboBox cmbUF 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCidadao.frx":01CA
         Left            =   1080
         List            =   "frmCidadao.frx":01CC
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   915
         Width           =   2250
      End
      Begin esMaskEdit.esMaskedEdit mskCEP 
         Height          =   285
         Left            =   4425
         TabIndex        =   27
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":01CE
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
         Left            =   3345
         TabIndex        =   26
         ToolTipText     =   "Cadastrar um novo bairro"
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
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
         MICON           =   "frmCidadao.frx":01EA
         PICN            =   "frmCidadao.frx":0504
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número........:"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   48
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   47
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complem.:"
         Height          =   225
         Index           =   4
         Left            =   2175
         TabIndex        =   46
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...........:"
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   45
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   225
         Index           =   6
         Left            =   3810
         TabIndex        =   44
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...............:"
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   43
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP....:"
         Height          =   225
         Index           =   8
         Left            =   3810
         TabIndex        =   42
         Top             =   1365
         Width           =   585
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   825
      Left            =   60
      Top             =   4980
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   1455
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   40
         Top             =   450
         Width           =   2475
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         MaxLength       =   50
         TabIndex        =   41
         Top             =   450
         Width           =   2940
      End
      Begin VB.TextBox txtPais 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   39
         Top             =   90
         Width           =   2475
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone......:"
         Height          =   225
         Index           =   9
         Left            =   60
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email...:"
         Height          =   225
         Index           =   10
         Left            =   3810
         TabIndex        =   19
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "País.............:"
         Height          =   180
         Index           =   15
         Left            =   60
         TabIndex        =   17
         Top             =   135
         Width           =   990
      End
   End
   Begin prjChameleon.chameleonButton cmdCopy 
      Height          =   315
      Left            =   4575
      TabIndex        =   15
      ToolTipText     =   "Copiar dados do cidadão selecioando"
      Top             =   5910
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Co&piar"
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
      MICON           =   "frmCidadao.frx":065E
      PICN            =   "frmCidadao.frx":067A
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
      Left            =   6525
      TabIndex        =   14
      ToolTipText     =   "Gravar os Dados"
      Top             =   5910
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
      MICON           =   "frmCidadao.frx":0765
      PICN            =   "frmCidadao.frx":0781
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   5910
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmCidadao.frx":0B26
      PICN            =   "frmCidadao.frx":0B42
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
      Left            =   6525
      TabIndex        =   9
      ToolTipText     =   "Sair da Tela"
      Top             =   5910
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
      MICON           =   "frmCidadao.frx":0C9C
      PICN            =   "frmCidadao.frx":0CB8
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
      Left            =   5430
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   5910
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
      MICON           =   "frmCidadao.frx":0D26
      PICN            =   "frmCidadao.frx":0D42
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
      Left            =   2385
      TabIndex        =   11
      ToolTipText     =   "Excluir Registro"
      Top             =   5910
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCidadao.frx":0E9C
      PICN            =   "frmCidadao.frx":0EB8
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
      TabIndex        =   12
      ToolTipText     =   "Editar Registro"
      Top             =   5910
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
      MICON           =   "frmCidadao.frx":0F5A
      PICN            =   "frmCidadao.frx":0F76
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
      Left            =   195
      TabIndex        =   13
      ToolTipText     =   "Novo Registro"
      Top             =   5910
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
      MICON           =   "frmCidadao.frx":10D0
      PICN            =   "frmCidadao.frx":10EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   1695
      Index           =   1
      Left            =   60
      Top             =   3270
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   2990
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço Comercial"
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.CheckBox chkEtiq2 
         Alignment       =   1  'Right Justify
         Caption         =   "Etiqueta"
         Height          =   225
         Left            =   6480
         TabIndex        =   38
         Top             =   1350
         Width           =   885
      End
      Begin VB.ComboBox cmbUF2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCidadao.frx":1246
         Left            =   1080
         List            =   "frmCidadao.frx":1248
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   915
         Width           =   2250
      End
      Begin VB.ComboBox cmbCidade2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4425
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   930
         Width           =   3000
      End
      Begin VB.ComboBox cmbBairro2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   1305
         Width           =   2250
      End
      Begin VB.TextBox txtCompl2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2970
         MaxLength       =   50
         TabIndex        =   32
         Top             =   585
         Width           =   4455
      End
      Begin VB.TextBox txtNum2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Top             =   585
         Width           =   930
      End
      Begin VB.TextBox txtNumLog2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   240
         Width           =   930
      End
      Begin VB.TextBox txtNomeLog2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   30
         Top             =   240
         Width           =   5220
      End
      Begin esMaskEdit.esMaskedEdit mskCEP2 
         Height          =   285
         Left            =   4425
         TabIndex        =   37
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MouseIcon       =   "frmCidadao.frx":124A
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
      Begin prjChameleon.chameleonButton cmdAddBairro2 
         Height          =   270
         Left            =   3345
         TabIndex        =   36
         ToolTipText     =   "Cadastrar um novo bairro"
         Top             =   1320
         Width           =   360
         _ExtentX        =   635
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
         MICON           =   "frmCidadao.frx":1266
         PICN            =   "frmCidadao.frx":1580
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP....:"
         Height          =   225
         Index           =   22
         Left            =   3810
         TabIndex        =   55
         Top             =   1365
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...............:"
         Height          =   225
         Index           =   21
         Left            =   60
         TabIndex        =   54
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   225
         Index           =   20
         Left            =   3810
         TabIndex        =   53
         Top             =   990
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...........:"
         Height          =   225
         Index           =   19
         Left            =   60
         TabIndex        =   52
         Top             =   1365
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complem.:"
         Height          =   225
         Index           =   18
         Left            =   2175
         TabIndex        =   51
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   17
         Left            =   90
         TabIndex        =   50
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número........:"
         Height          =   225
         Index           =   16
         Left            =   60
         TabIndex        =   49
         Top             =   615
         Width           =   975
      End
   End
   Begin VB.ListBox lstNomeLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      ItemData        =   "frmCidadao.frx":16DA
      Left            =   2280
      List            =   "frmCidadao.frx":16DC
      TabIndex        =   63
      Top             =   1800
      Visible         =   0   'False
      Width           =   5205
   End
End
Attribute VB_Name = "frmCidadao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bZOrder As Boolean
Dim RdoAux As rdoResultset
Dim Sql As String, bExec As Boolean
Dim Evento As String, sEnd As String
Dim frm As frmCnsCidadao
Dim NomeForm As String, sTipoCid As String

Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Public Property Let sTipoCidadao(sValue As String)
    sTipoCid = sValue
End Property

Private Sub chkEtiq_Click()
chkEtiq2.Value = IIf(chkEtiq.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Private Sub chkEtiq2_Click()
chkEtiq.Value = IIf(chkEtiq2.Value = vbChecked, vbUnchecked, vbChecked)
End Sub

Public Sub cmbCidade_Click()
If Not bExec Then Exit Sub
If cmbCidade.ListIndex = -1 Then Exit Sub
cmbBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    cmbBairro.AddItem " "
    Do While Not .EOF
       cmbBairro.AddItem !DescBairro
       cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
       .MoveNext
    Loop
   .Close
End With

End Sub

Public Sub cmbCidade2_Click()
If Not bExec Then Exit Sub
If cmbCidade2.ListIndex = -1 Then Exit Sub
cmbBairro2.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF2.Text, 2) & "' AND CODCIDADE=" & cmbCidade2.ItemData(cmbCidade2.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    cmbBairro2.AddItem " "
    Do While Not .EOF
       cmbBairro2.AddItem !DescBairro
       cmbBairro2.ItemData(cmbBairro2.NewIndex) = !CodBairro
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbFindCod_Click()
Dim z As Variant

z = InputBox("Digite o código do cidadão.", "Entre com a informação")
If Val(z) > 0 Then
    If Val(z) > 700000 Then
        MsgBox "Código inválido.", vbCritical, "Atenção"
    Else
        txtCod.Text = Val(z)
        Le
    End If
End If

End Sub

Private Sub cmbUF_Click()

If Not bExec Then Exit Sub
cmbCidade.Clear
cmbBairro.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do While Not .EOF
       cmbCidade.AddItem !desccidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbUF2_Click()

If Not bExec Then Exit Sub
cmbCidade2.Clear
cmbBairro2.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF2.Text, 2) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do While Not .EOF
       cmbCidade2.AddItem !desccidade
       cmbCidade2.ItemData(cmbCidade2.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With

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

Set frm2 = frmBairro
frm2.FormCall = Me.Name & "R"
frm2.SiglaUF = cmbUF.Text
frm2.CodCidade = cmbCidade.ItemData(cmbCidade.ListIndex)
frmBairro.show

End Sub

Private Sub cmdAddBairro2_Click()
If cmbUF2.ListIndex = -1 Or cmbCidade2.ListIndex = -1 Then
    MsgBox "Selecione UF e Cidade do bairro a ser cadastardo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Left(cmbUF2.Text, 2) = "SP" And cmbCidade2.ItemData(cmbCidade2.ListIndex) = 413 Then
    MsgBox "Apenas bairros de fora podem ser cadastrados.", vbExclamation, "Atenção"
    Exit Sub
End If

Set frm2 = frmBairro
frm2.FormCall = Me.Name & "C"
frm2.SiglaUF = cmbUF2.Text
frm2.CodCidade = cmbCidade2.ItemData(cmbCidade2.ListIndex)
frmBairro.show

End Sub

Private Sub cmdAlterar_Click()
    If Val(txtCod.Text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"
    txtNome.SetFocus
End Sub

Private Sub cmdCancel_Click()
Dim x As Long
    If Evento = "Alterar" Then
       x = Val(txtCod.Text)
       Limpa
       txtCod.Text = x
       Le
    Else
       Limpa
    End If
    lstNomeLog.Visible = False
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdConsultar_Click()
bZOrder = False
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show
frmCnsCidadao.ZOrder 0
End Sub

Private Sub cmdCopy_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset
If Val(txtCod.Text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

z = InputBox("Digite o código do cidadão original.", "Entre com a informação")
If Val(z) < 500000 Or Val(z) > 700000 Then
    MsgBox "Código inválido.", vbCritical, "Erro"
    Exit Sub
End If

Sql = "SELECT * FROM CIDADAO WHERE CODCIDADAO=" & Val(z)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Código não cadastrado.", vbCritical, "Erro"
    Else
        If MsgBox("Deseja copiar os dados do cidadão " & !CodCidadao & "-" & !nomecidadao & " ?", vbYesNo + vbQuestion, "Confirmação") = vbYes Then
            Log Form, Me.Caption, Alteração, "Copiado cidadão " & !CodCidadao & " - " & nomecidadao & " para o código " & txtCod.Text & " - " & txtNome.Text
            Sql = "update cidadao set cpf='" & SubNull(!CPF) & "',cnpj='" & SubNull(!Cnpj) & "',codlogradouro=" & Val(SubNull(!CodLogradouro)) & ","
            Sql = Sql & "numimovel=" & Val(SubNull(!NUMIMOVEL)) & ",complemento='" & SubNull(!Complemento) & "',codbairro=" & Val(SubNull(!CodBairro)) & ","
            Sql = Sql & "codcidade=" & Val(!CodCidade) & ",siglauf='" & SubNull(!SiglaUF) & "' where codcidadao=" & Val(txtCod.Text)
            cn.Execute Sql, rdExecDirect
            z = Val(txtCod.Text)
            Limpa
            txtCod.Text = z
            Le
        End If

    End If
   .Close
End With

End Sub

Private Sub cmdExcluir_Click()
Dim x As Integer
On Error GoTo Erro
If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "RENATA" And NomeDeLogin <> "LEILA" And NomeDeLogin <> "CINTIA" Then
   MsgBox "Não é possível excluir.", vbCritical, "Atenção"
   Exit Sub
End If

If Val(txtCod.Text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
ElseIf Val(txtCod.Text) < 500000 Then
   MsgBox "Este cidadão não pode ser excluido.", vbCritical, "Atenção"
   Exit Sub
End If

Sql = "SELECT * FROM PROPRIETARIO WHERE CODCIDADAO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
   MsgBox "Não é possível excluir este Cidadão pois ele é Proprietário/Proprietário Solidário de um Imóvel.", vbExclamation, "Atenção"
   Exit Sub
End If

If MsgBox("Excluir este Cidadão ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM CIDADAO WHERE CODCIDADAO=" & txtCod.Text
   cn.Execute Sql, rdExecDirect
   Limpa
End If
    
Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
   MsgBox rdoErrors(x).Description
Next

End Sub

Private Sub cmdGravar_Click()
If bLocal Then
    Exit Sub
End If

Dim y As Integer
    If txtNome.Text = "" Then
       MsgBox "Favor digitar o Nome do Cidadão.", vbExclamation, "Atenção"
       txtNome.SetFocus
       Exit Sub
    End If
    If Val(txtNum.Text) > 10000 Or Val(txtNum2.Text) > 10000 Then
       MsgBox "Nº de Imóvel inválido.", vbExclamation, "Atenção"
       Exit Sub
    End If
    
    If Val(txtNumLog.Text) = 0 And cmbCidade.Text = "JABOTICABAL" Then
       MsgBox "Selecione o Logradouro de Jaboticabal.", vbCritical, "Erro de Validação."
       txtNumLog.SetFocus
       Exit Sub
    End If
    
    If Val(txtNumLog2.Text) = 0 And cmbCidade2.Text = "JABOTICABAL" Then
       MsgBox "Selecione o Logradouro de Jaboticabal.", vbCritical, "Erro de Validação."
       txtNumLog2.SetFocus
       Exit Sub
    End If
    
    If chkEtiq.Value = vbChecked And Val(txtNumLog.Text) = 0 And txtNomeLog.Text = "" Then
       MsgBox "Digite o Logradouro.", vbCritical, "Erro de Validação."
       txtNomeLog.SetFocus
       Exit Sub
    End If
    
    If chkEtiq2.Value = vbChecked And Val(txtNumLog2.Text) = 0 And txtNomeLog2.Text = "" Then
       MsgBox "Digite o Logradouro.", vbCritical, "Erro de Validação."
       txtNomeLog2.SetFocus
       Exit Sub
    End If
    
    If mskCPF.ClipText = "" And mskCNPJ.ClipText = "" And mskRG.ClipText = "" Then
       MsgBox "Digite CPF ou CNPJ e/ou RG.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If txtOrgao.Text <> "" And mskRG.ClipText = "" Then
       MsgBox "Digite o RG.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If mskCPF.ClipText <> "" And mskCNPJ.ClipText <> "" Then
       MsgBox "Digite CPF ou CNPJ.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If Trim(txtPais.Text) <> "" And cmbUF.ListIndex > 0 Then
       MsgBox "País não pode ter cidade/UF.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If chkEtiq.Value = vbChecked And cmbCidade.ListIndex = -1 Then
       MsgBox "Favor selecionar a cidade.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If chkEtiq2.Value = vbChecked And cmbCidade2.ListIndex = -1 Then
       MsgBox "Favor selecionar a cidade.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If txtNomeLog.Text <> "" And cmbBairro.ListIndex = -1 Then
       MsgBox "Favor selecionar o bairro residencial.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If txtNomeLog2.Text <> "" And cmbBairro2.ListIndex = -1 Then
       MsgBox "Favor selecionar o bairro comercial.", vbExclamation, "Atenção"
       Exit Sub
    End If
       
    If chkEtiq.Value = vbChecked Then
        If Left(cmbUF.Text, 2) = "SP" And cmbCidade.ItemData(cmbCidade.ListIndex) = 413 Then
        Else
            If mskCEP.ClipText = "" Then
               MsgBox "Favor cadastrar o CEP do endereço de fora.", vbExclamation, "Atenção"
               Exit Sub
            End If
        End If
    End If
       
    If chkEtiq2.Value = vbChecked Then
        If Left(cmbUF2.Text, 2) = "SP" And cmbCidade2.ItemData(cmbCidade2.ListIndex) = 413 Then
        Else
            If mskCEP2.ClipText = "" Then
               MsgBox "Favor cadastrar o CEP do endereço de fora.", vbExclamation, "Atenção"
               Exit Sub
            End If
        End If
    End If
       
       
    If mskCPF.ClipText <> "" Then
       If Not ValidaCPF(mskCPF.ClipText) Then
          MsgBox "CPF inválido.", vbExclamation, "Atenção"
          Exit Sub
       End If
       If Evento = "Novo" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO>500000 AND CPF='" & mskCPF.ClipText & "'"
       Else
            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO>500000 AND CODCIDADAO<>" & Val(txtCod.Text) & " AND CPF='" & mskCPF.ClipText & "'"
       End If
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                   MsgBox "CPF já cadastrado ---> " & Format(!CodCidadao, "00000") & " - " & !nomecidadao, vbCritical, "CPF Duplicado !!!"
                   Exit Sub
                End If
            End With
            
'       End If
    End If
    
    If mskCNPJ.ClipText <> "" Then
       If Not ValidaCGC(mskCNPJ.ClipText) Then
          MsgBox "CNPJ inválido.", vbExclamation, "Atenção"
          Exit Sub
       End If
       If Evento = "Novo" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & mskCNPJ.ClipText & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                   MsgBox "CNPJ já cadastrado ---> " & Format(!CodCidadao, "00000") & " - " & !nomecidadao, vbCritical, "CNPJ Duplicado !!!"
                   Exit Sub
                End If
            End With
       End If
    End If
    
    If mskRG.ClipText <> "" Then
       If Evento = "Novo" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE RG='" & mskRG.Text & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                   MsgBox "RG já cadastrado ---> " & Format(!CodCidadao, "00000") & " - " & !nomecidadao, vbCritical, "RG Duplicado !!!"
                   Exit Sub
                End If
            End With
       End If
    End If
    
    Grava
    Eventos "INICIAR"
    If Evento = "Novo" Then
        For x = 0 To Forms.Count - 1
            If Forms(x).Name = "frmCadImob" Then
                For y = 1 To frmCadImob.tvProp.Nodes.Count
                    On Error GoTo fim
                    If Right$(frmCadImob.tvProp.Nodes(x).Key, 5) = Format(txtCod.Text, "00000") Then
                       MsgBox "Ja existe um Proprietário ou Proprietário Solidário com este nome.", vbCritical, "Atenção"
                       GoTo fim
                    End If
                Next
                If sTipoCid = "P" Then
                   Set NodX = frmCadImob.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(txtCod.Text, "00000"), txtNome.Text, 1)
                   frmCadImob.tvProp.Nodes("PROP" & Format(txtCod.Text, "00000")).ForeColor = vbBlue
                Else
                   Set NodX = frmCadImob.tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(txtCod.Text, "00000"), txtNome.Text, 2)
                   frmCadImob.tvProp.Nodes("COMP" & Format(txtCod.Text, "00000")).ForeColor = vbBlue
                End If
                For y = 1 To frmCadImob.tvProp.Nodes.Count
                    frmCadImob.tvProp.Nodes(x).EnsureVisible
                Next
fim:
                sTipoCid = ""
                Unload frmCidadao
                frmCadImob.SetFocus
                Exit Sub
            ElseIf Forms(x).Name = "frmCadMob" Then
                For y = 1 To frmCadMob.grdProp.Rows - 1
                    If Val(frmCadMob.grdProp.TextMatrix(y, 0)) = Val(txtCod.Text) Then
                       MsgBox "Ja existe um Proprietário ou Proprietário Solidário com este nome.", vbCritical, "Atenção"
                       GoTo Fim2
                    End If
                Next
                frmCadMob.grdProp.AddItem Format(txtCod.Text, "00000") & Chr(9) & txtNome.Text & Chr(9) & IIf(mskCPF.ClipText <> "", mskCPF.Text, mskCNPJ.Text)
Fim2:
                sTipoCid = ""
                Unload frmCidadao
                frmCadMob.SetFocus
                Exit Sub
            End If
        Next
    End If
End Sub

Public Sub cmdNovo_Click()
    Limpa
    cmbUF.Text = "SP-SÃO PAULO"
    cmbUF_Click
    cmbCidade.Text = "JABOTICABAL"
    cmbCidade_Click
    Eventos "INCLUIR"
    Evento = "Novo"
    txtNome.SetFocus
End Sub

Private Sub cmdSair_Click()
On Error Resume Next
Unload Me
End Sub

Private Sub Form_Activate()
On Error Resume Next
If CodCidadao > 0 Then
    Limpa
    txtCod.Text = Format(CodCidadao, "00000")
    Le
End If
CodCidadao = 0

If NomeForm = "frmProcesso" Then
    'bZOrder = True
    If frmProcesso.cmdGravar.Visible = False Then
        cmdNovo.Enabled = False
        cmdAlterar.Enabled = False
        cmdExcluir.Enabled = False
        cmdConsultar.Enabled = False
    End If
End If

Liberado

End Sub

Private Sub Form_Load()
bZOrder = True
Centraliza Me
sRet = RetEventUserForm(Me.Name)
Limpa
sEnd = "R"
bExec = False
cmbUF.AddItem " "
Sql = "SELECT SIGLAUF,DESCUF FROM UF"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
Do While Not RdoAux.EOF
   cmbUF.AddItem RdoAux!SiglaUF & "-" & RdoAux!DESCUF
   RdoAux.MoveNext
Loop
cmbUF2.AddItem " "
Sql = "SELECT SIGLAUF,DESCUF FROM UF"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
Do While Not RdoAux.EOF
   cmbUF2.AddItem RdoAux!SiglaUF & "-" & RdoAux!DESCUF
   RdoAux.MoveNext
Loop

bExec = True
RdoAux.Close

Eventos "INICIAR"

End Sub

Private Sub Le()
Dim RdoS As rdoResultset

Sql = "SELECT * FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    If .RowCount = 0 Then Exit Sub
    txtCod.Text = !CodCidadao
    If Not IsNull(!CPF) Then mskCPF.Text = Format(Trim(!CPF), "000\.000\.000-00")
    If Not IsNull(!Cnpj) Then mskCNPJ.Text = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
    If Not IsNull(!rg) Then mskRG.Text = !rg
    txtNome.Text = !nomecidadao
    If !CodLogradouro > 0 Then
       txtNumLog.Text = !CodLogradouro
       txtNumLog_LostFocus
    Else
       txtNumLog.Text = 0
       txtNomeLog.Text = SubNull(!NomeLogradouro)
    End If
    
    If !CodLogradouro2 > 0 Then
       txtNumLog2.Text = !CodLogradouro2
       txtNumLog2_LostFocus
    Else
       txtNumLog2.Text = 0
       txtNomeLog2.Text = SubNull(!NomeLogradouro2)
    End If
    
    txtNum.Text = SubNull(!NUMIMOVEL)
    txtCompl.Text = SubNull(!Complemento)
    
    txtNum2.Text = SubNull(!NUMIMOVEL2)
    txtCompl2.Text = SubNull(!Complemento2)
    
    If Trim$(SubNull(!SiglaUF)) <> "" Then
       bExec = False
       Sql = "SELECT DESCUF FROM UF WHERE SIGLAUF='" & !SiglaUF & "'"
       Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
       cmbUF.Text = !SiglaUF & "-" & RdoS!DESCUF
       bExec = True
       cmbUF_Click
       If !CodCidade > 0 Then
          bExec = False
            Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !SiglaUF & "' AND CODCIDADE=" & !CodCidade
            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
            cmbCidade.Text = RdoS!desccidade
            bExec = True
           cmbCidade_Click
           If !CodBairro > 0 Then
                Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !SiglaUF & "' AND CODCIDADE=" & !CodCidade & " AND CODBAIRRO=" & !CodBairro
                Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                If RdoS!DescBairro <> "" Then
                    cmbBairro.Text = RdoS!DescBairro
                End If
          End If
       End If
    End If
    
    
    If Trim$(SubNull(!SiglaUF2)) <> "" Then
       bExec = False
       Sql = "SELECT DESCUF FROM UF WHERE SIGLAUF='" & !SiglaUF2 & "'"
       Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
       cmbUF2.Text = !SiglaUF2 & "-" & RdoS!DESCUF
       bExec = True
       cmbUF2_Click
       If !CodCidade2 > 0 Then
          bExec = False
            Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !SiglaUF2 & "' AND CODCIDADE=" & !CodCidade2
            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
            cmbCidade2.Text = RdoS!desccidade
            bExec = True
           cmbCidade2_Click
           If !CodBairro2 > 0 Then
                Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !SiglaUF2 & "' AND CODCIDADE=" & !CodCidade2 & " AND CODBAIRRO=" & !CodBairro2
                Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                If RdoS!DescBairro <> "" Then
                    cmbBairro2.Text = RdoS!DescBairro
                End If
          End If
       End If
    End If
    
    If Not IsNull(!Cep) Then
       mskCEP.Text = Format(!Cep, "00000-000")
    End If
    If Not IsNull(!Cep2) Then
       mskCEP2.Text = Format(!Cep2, "00000-000")
    End If
    txtFone.Text = SubNull(!TELEFONE)
    txtEmail.Text = SubNull(!EMAIL)
    txtOrgao.Text = SubNull(!ORGAO)
    txtPais.Text = SubNull(!PAIS)
    optP(0).Value = True
    If Not IsNull(!JURIDICA) Then
        If !JURIDICA Then
            optP(1).Value = True
        End If
    End If
    If IsNull(!etiqueta) Then
        chkEtiq.Value = vbChecked
    Else
        If !etiqueta = "R" Then
            chkEtiq.Value = vbChecked
        Else
            chkEtiq2.Value = vbChecked
        End If
    End If
   .Close
End With

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdCopy.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdConsultar.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmCidadao
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
           Ct.Enabled = False
       End If
   Next
   mskCPF.Enabled = False
   mskCPF.BackColor = Kde
   mskRG.Enabled = False
   mskRG.BackColor = Kde
   mskCEP.Enabled = False
   mskCEP.BackColor = Kde
   mskCEP2.Enabled = False
   mskCEP2.BackColor = Kde
   mskCNPJ.Enabled = False
   mskCNPJ.BackColor = Kde
   chkEtiq.Enabled = False
   chkEtiq2.Enabled = False
   optP(0).Enabled = False
   optP(1).Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdCopy.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdConsultar.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmCidadao
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   mskCPF.Enabled = True
   mskCPF.BackColor = vbWhite
   mskRG.Enabled = True
   mskRG.BackColor = vbWhite
   mskCNPJ.Enabled = True
   mskCNPJ.BackColor = vbWhite
   mskCEP.Enabled = True
   mskCEP.BackColor = vbWhite
   mskCEP2.Enabled = True
   mskCEP2.BackColor = vbWhite
   txtCod.Locked = True
   txtCod.BackColor = Kde
   chkEtiq.Enabled = True
   chkEtiq2.Enabled = True
   optP(0).Enabled = True
   optP(1).Enabled = True
End If

FormHagana

End Sub

Private Sub Limpa()

txtCod.Text = ""
LimpaMascara mskCPF
LimpaMascara mskCNPJ
LimpaMascara mskRG
txtNome.Text = ""
txtNumLog.Text = 0
txtNomeLog.Text = ""
txtNum.Text = 0
txtCompl.Text = ""
cmbBairro.ListIndex = -1
cmbCidade.ListIndex = -1
cmbUF.ListIndex = -1
LimpaMascara mskCEP
chkEtiq.Value = vbChecked

txtNumLog2.Text = 0
txtNomeLog2.Text = ""
txtNum2.Text = 0
txtCompl2.Text = ""
cmbBairro2.ListIndex = -1
cmbCidade2.ListIndex = -1
cmbUF2.ListIndex = -1
LimpaMascara mskCEP2
chkEtiq2.Value = vbUnchecked

txtFone.Text = ""
txtEmail.Text = ""
txtPais.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

On Error Resume Next
If NomeForm = "frmProcesso" And Val(txtCod.Text) < 500000 Then
    MsgBox "Selecione apenas cidadão acima de 500000.", vbExclamation, "Atenção"
    frmProcesso.HabilitaPainelPrincipal
    frmProcesso.cmdEditCid.Enabled = True
    frmProcesso.cmdEditDoc.Enabled = True
    frmProcesso.cmdEditEnd.Enabled = True
    Exit Sub
End If
If NomeForm = "frmProcesso" Then
    frmProcesso.lblCodCid.Caption = txtCod.Text
    frmProcesso.lblNomeCid.Caption = txtNome.Text
    frmProcesso.HabilitaPainelPrincipal
    If sTipoCid = "A" Then
        frmProcesso.cmbOrigem.BackColor = Kde
        frmProcesso.cmbAssunto.BackColor = Kde
        frmProcesso.chkFisico.BackColor = Kde
        frmProcesso.chkInterno.BackColor = Kde
        frmProcesso.txtCompl.BackColor = Kde
        frmProcesso.cmbOrigem.Enabled = False
        frmProcesso.cmbAssunto.Enabled = False
        frmProcesso.chkFisico.Enabled = False
        frmProcesso.chkInterno.Enabled = False
        frmProcesso.txtCompl.Locked = True
    End If
    frmProcesso.cmdEditCid.Enabled = True
    frmProcesso.cmdEditDoc.Enabled = True
    frmProcesso.cmdEditEnd.Enabled = True
    frmProcesso.SetFocus
    NomeForm = ""
End If
If NomeForm = "frmCPProcesso" And Val(txtCod.Text) < 500000 Then
    MsgBox "Selecione apenas cidadão acima de 500000.", vbExclamation, "Atenção"
    frmCPProcesso.HabilitaPainelPrincipal
    frmCPProcesso.cmdEditCid.Enabled = True
    frmCPProcesso.cmdEditDoc.Enabled = True
    frmCPProcesso.cmdEditEnd.Enabled = True
    Exit Sub
End If
If NomeForm = "frmCPProcesso" Then
    frmCPProcesso.lblCodCid.Caption = txtCod.Text
    frmCPProcesso.lblNomeCid.Caption = txtNome.Text
    frmCPProcesso.HabilitaPainelPrincipal
    If sTipoCid = "A" Then
        frmCPProcesso.cmbOrigem.BackColor = Kde
        frmCPProcesso.cmbAssunto.BackColor = Kde
        frmCPProcesso.chkFisico.BackColor = Kde
        frmCPProcesso.chkInterno.BackColor = Kde
        frmCPProcesso.txtCompl.BackColor = Kde
        frmCPProcesso.cmbOrigem.Enabled = False
        frmCPProcesso.cmbAssunto.Enabled = False
        frmCPProcesso.chkFisico.Enabled = False
        frmCPProcesso.chkInterno.Enabled = False
        frmCPProcesso.txtCompl.Locked = True
    End If
    frmCPProcesso.cmdEditCid.Enabled = True
    frmCPProcesso.cmdEditDoc.Enabled = True
    frmCPProcesso.cmdEditEnd.Enabled = True
    frmCPProcesso.SetFocus
    NomeForm = ""
End If
If NomeForm = "frmRequerimento" Then
    frmRequerimento.lblCodRequerente.Caption = txtCod.Text
    frmRequerimento.lblRequerente.Caption = txtNome.Text
ElseIf NomeForm = "frmRequerIPTU" Then
    frmRequerIPTU.lblRequerente.Tag = txtCod.Text
    frmRequerIPTU.lblRequerente.Caption = txtNome.Text
ElseIf NomeForm = "frmRequerIPTU2" Then
    frmRequerIPTU.lblRazao.Tag = txtCod.Text
    frmRequerIPTU.lblRazao.Caption = txtNome.Text
    frmRequerIPTU.lblCNPJ.Caption = mskCNPJ.Text
    frmRequerIPTU.txtEnd.Text = txtNomeLog.Text & ", " & txtNum.Text
ElseIf NomeForm = "frmDeclaraIsento" Then
    frmDeclaraIsento.lblRequerente.Tag = txtCod.Text
    frmDeclaraIsento.lblRequerente.Caption = txtNome.Text
    frmDeclaraIsento.lblCPF.Caption = mskCPF.Text
    frmDeclaraIsento.lblRG.Caption = mskRG.Text
ElseIf NomeForm = "frmRegistroAtendimento" Then
    frmRegistroAtendimento.txtCidadao.Text = Format(txtCod.Text, "000000") & " - " & txtNome.Text
    frmRegistroAtendimento.txtEnd.Text = txtNomeLog.Text & ", " & txtNum.Text
    frmRegistroAtendimento.txtBairro = cmbBairro.Text
    frmRegistroAtendimento.txtCompl.Text = txtCompl.Text
    frmRegistroAtendimento.txtFone.Text = txtFone.Text
End If

'CodCidadao = Val(txtCod.Text)
Unload frmCnsCidadao
Unload frmCidadao
End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub mskCEP_GotFocus()
mskCEP.SelStart = 0
mskCEP.SelLength = Len(mskCEP.Text)
End Sub

Private Sub mskCEP2_GotFocus()
mskCEP2.SelStart = 0
mskCEP2.SelLength = Len(mskCEP2.Text)
End Sub

Private Sub mskCNPJ_GotFocus()
mskCNPJ.SelStart = 0
mskCNPJ.SelLength = Len(mskCNPJ.Text)
'mskCNPJ.SetFocus
End Sub

Private Sub mskCPF_GotFocus()
mskCPF.SelStart = 0
mskCPF.SelLength = Len(mskCPF.Text)
'mskCPF.SetFocus
End Sub

Private Sub mskRG_GotFocus()
mskRG.SelStart = 0
mskRG.SelLength = Len(mskRG.Text)
'mskRG.SetFocus
End Sub

Private Sub txtCompl_GotFocus()
txtCompl.SelStart = 0
txtCompl.SelLength = Len(txtCompl)
End Sub

Private Sub txtCompl2_GotFocus()
txtCompl2.SelStart = 0
txtCompl2.SelLength = Len(txtCompl2)
End Sub

Private Sub txtEmail_GotFocus()
txtEmail.SelStart = 0
txtEmail.SelLength = Len(txtEmail)
End Sub

Private Sub txtFone_GotFocus()
txtFone.SelStart = 0
txtFone.SelLength = Len(txtFone)
End Sub

Private Sub txtNome_GotFocus()
txtNome.SelStart = 0
txtNome.SelLength = Len(txtNome)
End Sub

Private Sub txtNum_GotFocus()
txtNum.SelStart = 0
txtNum.SelLength = Len(txtNum)
End Sub

Private Sub txtNum2_GotFocus()
txtNum2.SelStart = 0
txtNum2.SelLength = Len(txtNum2)
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 44 Or KeyAscii = 45) Then
       KeyAscii = 0
    End If
End Sub

Private Sub txtNum_LostFocus()
Dim s As String

LimpaMascara mskCEP
If cmbCidade.Text <> "JABOTICABAL" Then
    Exit Sub
End If


If txtNomeLog.Text <> "" Then
    s = RetornaCEP(Val(txtNumLog.Text), Val(txtNum.Text))
    mskCEP.Text = s
End If

End Sub

Private Sub txtNum2_LostFocus()
Dim s As String

LimpaMascara mskCEP2
If cmbCidade2.Text <> "JABOTICABAL" Then
    Exit Sub
End If

If txtNomeLog2.Text <> "" Then
    s = RetornaCEP(Val(txtNumLog2.Text), Val(txtNum2.Text))
    mskCEP2.Text = s
End If

End Sub

Private Sub Grava()

Dim MaxCod As Long, nBairro As Integer, nCidade As Integer, nPessoa As Integer, nBairro2 As Integer, nCidade2 As Integer, sEtiq As String

Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!MAXIMO) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!MAXIMO + 1
End If
RdoAux.Close

If optP(0).Value = True Then
    nPessoa = 0
Else
    nPessoa = 1
End If

If chkEtiq.Value = vbChecked Then
    sEtiq = "R"
Else
    sEtiq = "C"
End If

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

If cmbBairro2.ListIndex > -1 Then
    nBairro2 = cmbBairro2.ItemData(cmbBairro2.ListIndex)
Else
    nBairro2 = 0
End If
If cmbCidade2.ListIndex > -1 Then
    nCidade2 = cmbCidade2.ItemData(cmbCidade2.ListIndex)
Else
    nCidade2 = 0
End If


If Evento = "Novo" Then
    Sql = "INSERT CIDADAO(CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO,NUMIMOVEL,COMPLEMENTO,CODBAIRRO,CODCIDADE,"
    Sql = Sql & "SIGLAUF,CEP,TELEFONE,EMAIL,RG,NOMELOGRADOURO,ORGAO,JURIDICA,PAIS,CODLOGRADOURO2,NUMIMOVEL2,COMPLEMENTO2,CODBAIRRO2,CODCIDADE2,"
    Sql = Sql & "SIGLAUF2,CEP2,NOMELOGRADOURO2,ETIQUETA) VALUES(" & MaxCod & ",'" & Mask(txtNome.Text) & "','"
    Sql = Sql & IIf(mskCPF.ClipText = "", "", mskCPF.ClipText) & "','" & IIf(mskCNPJ.ClipText = "", "", mskCNPJ.ClipText) & "',"
    Sql = Sql & IIf(Val(txtNumLog.Text) > 0, Val(txtNumLog.Text), "Null") & "," & Val(txtNum.Text) & ",'"
    Sql = Sql & Mask(txtCompl.Text) & "'," & IIf(nBairro > 0, nBairro, "Null") & ","
    Sql = Sql & IIf(nCidade > 0, nCidade, "Null") & ",'" & Left$(cmbUF.Text, 2) & "',"
    Sql = Sql & IIf(Trim(mskCEP.ClipText) = "", "Null", "'" & mskCEP.ClipText & "'") & ",'" & Mask(txtFone.Text) & "','"
    Sql = Sql & Mask(txtEmail.Text) & "','" & IIf(mskRG.ClipText = "", "", mskRG.ClipText) & "'," & IIf(Val(txtNumLog.Text) > 0, "Null", "'" & txtNomeLog.Text & "'") & ",'"
    Sql = Sql & Mask(txtOrgao.Text) & "'," & nPessoa & ",'" & Mask(txtPais.Text) & "'," & IIf(Val(txtNumLog2.Text) > 0, Val(txtNumLog2.Text), "Null") & "," & Val(txtNum2.Text) & ",'"
    Sql = Sql & Mask(txtCompl2.Text) & "'," & IIf(nBairro2 > 0, nBairro2, "Null") & "," & IIf(nCidade2 > 0, nCidade2, "Null") & ",'"
    Sql = Sql & Left$(cmbUF2.Text, 2) & "'," & IIf(Trim(mskCEP2.ClipText) = "", "Null", "'" & mskCEP2.ClipText & "'") & "," & IIf(Val(txtNumLog2.Text) > 0, "Null", "'" & txtNomeLog2.Text & "'") & ",'" & sEtiq & "')"
Else
    Sql = "UPDATE CIDADAO SET NOMECIDADAO='" & Mask(txtNome.Text) & "',CPF='" & IIf(mskCPF.ClipText = "", "", mskCPF.ClipText) & "',"
    Sql = Sql & "CNPJ='" & IIf(mskCNPJ.ClipText = "", "", mskCNPJ.ClipText) & "',CODLOGRADOURO=" & IIf(Val(txtNumLog.Text) > 0, Val(txtNumLog.Text), "Null") & ","
    Sql = Sql & "NUMIMOVEL=" & Val(txtNum.Text) & ",COMPLEMENTO='" & Mask(txtCompl.Text) & "',CODBAIRRO=" & IIf(nBairro > 0, nBairro, "Null") & ","
    Sql = Sql & "CODCIDADE=" & IIf(nCidade > 0, nCidade, "Null") & ",SIGLAUF='" & Left$(cmbUF.Text, 2) & "',"
    Sql = Sql & "CEP=" & IIf(Trim(mskCEP.ClipText) = "", "Null", "'" & mskCEP.ClipText & "'") & ",TELEFONE='" & Mask(txtFone.Text) & "',"
    Sql = Sql & "EMAIL='" & Mask(txtEmail.Text) & "',RG='" & IIf(mskRG.ClipText = "", "", mskRG.ClipText) & "',NOMELOGRADOURO=" & IIf(Val(txtNumLog.Text) > 0, "Null", "'" & txtNomeLog.Text & "'") & ","
    Sql = Sql & "ORGAO='" & Mask(txtOrgao.Text) & "',JURIDICA=" & nPessoa & ",PAIS='" & Mask(txtPais.Text) & "',CODLOGRADOURO2=" & IIf(Val(txtNumLog2.Text) > 0, Val(txtNumLog2.Text), "Null") & ","
    Sql = Sql & "NUMIMOVEL2=" & Val(txtNum2.Text) & ",COMPLEMENTO2='" & Mask(txtCompl2.Text) & "',CODBAIRRO2=" & IIf(nBairro2 > 0, nBairro2, "Null") & ",CODCIDADE2=" & IIf(nCidade2 > 0, nCidade2, "Null") & ","
    Sql = Sql & "SIGLAUF2='" & Left$(cmbUF2.Text, 2) & "',CEP2=" & IIf(Trim(mskCEP2.ClipText) = "", "Null", "'" & mskCEP2.ClipText & "'") & ",NOMELOGRADOURO2=" & IIf(Val(txtNumLog2.Text) > 0, "Null", "'" & txtNomeLog2.Text & "'") & ","
    Sql = Sql & "ETIQUETA='" & sEtiq & "' Where CodCidadao = " & Val(txtCod.Text)
End If
 cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   txtCod.Text = MaxCod
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

Private Sub txtNomeLog_Change()
If Trim$(txtNomeLog) = "" Then
   txtNumLog.Text = 0
End If
End Sub

Private Sub txtNomeLog2_Change()
If Trim$(txtNomeLog2) = "" Then
   txtNumLog2.Text = 0
End If
End Sub

Private Sub txtNomeLog_GotFocus()
txtNomeLog.SelStart = 0
txtNomeLog.SelLength = Len(txtNomeLog)
End Sub

Private Sub txtNomeLog2_GotFocus()
txtNomeLog2.SelStart = 0
txtNomeLog2.SelLength = Len(txtNomeLog2)
End Sub

Private Sub txtNomeLog_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLog.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLog.Text) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Left = 2250
             lstNomeLog.Top = 1800
             sEnd = "R"
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder 0
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro não encontrado.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLog.SetFocus
          End If
      End With
   End If
Else
   txtNumLog.Text = 0
End If

End Sub

Private Sub txtNomeLog2_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLog2.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLog2) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Left = 2250
             lstNomeLog.Top = 3510
             sEnd = "C"
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder 0
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro não encontrado.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLog.SetFocus
          End If
      End With
   End If
Else
   txtNumLog2.Text = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
    If sEnd = "R" Then
        txtNumLog.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
        txtNumLog_LostFocus
        txtNum.SetFocus
    Else
        txtNumLog2.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
        txtNumLog2_LostFocus
        txtNum2.SetFocus
    End If
End If

lstNomeLog.Visible = False
End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
        If sEnd = "R" Then
            txtNumLog.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
            txtNumLog_LostFocus
            txtNum.SetFocus
        Else
            txtNumLog2.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
            txtNumLog2_LostFocus
            txtNum2.SetFocus
        End If
    End If
    lstNomeLog.Visible = False
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
End If

End Sub

Private Sub txtNumLog_LostFocus()
If Val(txtNumLog.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtNumLog.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLog.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLog.Text = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtNumLog.SetFocus
       End If
      .Close
   End With
End If

End Sub

Private Sub txtNumLog2_LostFocus()
If Val(txtNumLog2.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtNumLog2.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLog2.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLog2.Text = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtNumLog2.SetFocus
       End If
      .Close
   End With
End If

End Sub


