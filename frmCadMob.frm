VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmCadMob 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro mobiliário"
   ClientHeight    =   5700
   ClientLeft      =   7560
   ClientTop       =   2685
   ClientWidth     =   11490
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   11490
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   309
      Top             =   4095
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Documentos"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   305
      Top             =   3735
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Valor Adicional"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   245
      Top             =   3360
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "ISS Eletrônico"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   735
      Left            =   0
      Top             =   4470
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1296
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
      Begin VB.Label lblSusp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SUSPENSA (12/05/2003)"
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
         Height          =   495
         Left            =   180
         TabIndex        =   239
         Top             =   90
         Visible         =   0   'False
         Width           =   1635
      End
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   315
      Left            =   10260
      TabIndex        =   58
      ToolTipText     =   "Cancelar Edição"
      Top             =   5310
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
      MICON           =   "frmCadMob.frx":094E
      PICN            =   "frmCadMob.frx":096A
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
      Left            =   10260
      TabIndex        =   51
      ToolTipText     =   "Sair da Tela"
      Top             =   5310
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
      MICON           =   "frmCadMob.frx":0AC4
      PICN            =   "frmCadMob.frx":0AE0
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
      Left            =   5040
      TabIndex        =   52
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   5310
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
      MICON           =   "frmCadMob.frx":0B4E
      PICN            =   "frmCadMob.frx":0B6A
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
      Left            =   9180
      TabIndex        =   53
      ToolTipText     =   "Gravar os Dados"
      Top             =   5310
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
      MICON           =   "frmCadMob.frx":0CC4
      PICN            =   "frmCadMob.frx":0CE0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   6150
      TabIndex        =   54
      ToolTipText     =   "Imprimir esta Tela"
      Top             =   5310
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmCadMob.frx":1085
      PICN            =   "frmCadMob.frx":10A1
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
      Left            =   2250
      TabIndex        =   55
      ToolTipText     =   "EXCLUSÃO DA EMPRESA"
      Top             =   5310
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "EXCLUSÃO DA EMPRESA"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadMob.frx":11FB
      PICN            =   "frmCadMob.frx":1217
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
      Left            =   1170
      TabIndex        =   56
      ToolTipText     =   "Editar Registro"
      Top             =   5310
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
      MICON           =   "frmCadMob.frx":12B9
      PICN            =   "frmCadMob.frx":12D5
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
      Left            =   90
      TabIndex        =   57
      ToolTipText     =   "Novo Registro"
      Top             =   5310
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
      MICON           =   "frmCadMob.frx":142F
      PICN            =   "frmCadMob.frx":144B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1980
      Left            =   105
      TabIndex        =   50
      Top             =   6855
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   3493
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   $"frmCadMob.frx":15A5
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Dados Gerais"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":162C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Localização"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":1946
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Proprietário"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":1C60
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Endereço Entrega"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":1F7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Atividades"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":2294
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   1935
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Histórico"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":25AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Outros"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":28C8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   2640
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Informativo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":2BE2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btMenu 
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   3000
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      BTYPE           =   8
      TX              =   "Livro"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4194304
      BCOLO           =   16744576
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   13026246
      MPTR            =   99
      MICON           =   "frmCadMob.frx":2EFC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExtrato 
      Height          =   315
      Left            =   7245
      TabIndex        =   273
      ToolTipText     =   "Consulta Extrato"
      Top             =   5310
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Extrato"
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
      MICON           =   "frmCadMob.frx":3216
      PICN            =   "frmCadMob.frx":3232
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   2
      Left            =   2040
      TabIndex        =   94
      Top             =   45
      Width           =   9405
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Proprietários e Sócios"
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
         Height          =   2565
         Left            =   150
         TabIndex        =   109
         Top             =   240
         Width           =   8985
         Begin VB.CommandButton cmdCadCid 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Cadastrar"
            Height          =   315
            Left            =   2520
            Style           =   1  'Graphical
            TabIndex        =   112
            ToolTipText     =   "Cadastrar novo cidadão"
            Top             =   2010
            Width           =   1155
         End
         Begin VB.CommandButton cmdDelCid 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Remover"
            Height          =   315
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   111
            ToolTipText     =   "Remover proprietário da empresa"
            Top             =   2010
            Width           =   1155
         End
         Begin VB.CommandButton cmdAddCid 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Adicionar"
            Height          =   315
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   110
            ToolTipText     =   "Adicionar Proprietário a Empresa"
            Top             =   2010
            Width           =   1155
         End
         Begin MSFlexGridLib.MSFlexGrid grdProp 
            Height          =   1515
            Left            =   120
            TabIndex        =   113
            Top             =   330
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   2672
            _Version        =   393216
            Rows            =   1
            Cols            =   3
            FixedCols       =   0
            BackColorFixed  =   15658734
            BackColorSel    =   8388608
            BackColorBkg    =   15658734
            FocusRect       =   0
            GridLinesFixed  =   1
            SelectionMode   =   1
            Appearance      =   0
            FormatString    =   $"frmCadMob.frx":331D
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Contato"
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
         Height          =   2085
         Left            =   150
         TabIndex        =   95
         Top             =   2970
         Width           =   8985
         Begin VB.TextBox txtEmailNF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            MaxLength       =   100
            TabIndex        =   103
            Top             =   1680
            Width           =   6975
         End
         Begin VB.TextBox txtFoneNF 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6210
            MaxLength       =   12
            TabIndex        =   101
            Top             =   1020
            Width           =   2505
         End
         Begin VB.TextBox txtDDDNF 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5610
            MaxLength       =   2
            TabIndex        =   100
            Top             =   1020
            Width           =   555
         End
         Begin VB.TextBox txtNomeContato 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            MaxLength       =   50
            TabIndex        =   96
            Top             =   330
            Width           =   6975
         End
         Begin VB.TextBox txtCargo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            MaxLength       =   25
            TabIndex        =   97
            Top             =   675
            Width           =   2595
         End
         Begin VB.TextBox txtFax 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            MaxLength       =   25
            TabIndex        =   99
            Top             =   1020
            Width           =   2595
         End
         Begin VB.TextBox txtFone 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5610
            MaxLength       =   40
            TabIndex        =   98
            Top             =   675
            Width           =   3105
         End
         Begin VB.TextBox txtEmail 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1740
            MaxLength       =   100
            TabIndex        =   102
            Top             =   1350
            Width           =   6975
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Email NF...........:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   66
            Left            =   150
            TabIndex        =   359
            Top             =   1710
            Width           =   1485
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone NF:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   65
            Left            =   4440
            TabIndex        =   358
            Top             =   1020
            Width           =   1125
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Email................:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   45
            Left            =   150
            TabIndex        =   108
            Top             =   1380
            Width           =   1485
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome Completo..:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   48
            Left            =   150
            TabIndex        =   107
            Top             =   360
            Width           =   1665
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Cargo................:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   47
            Left            =   150
            TabIndex        =   106
            Top             =   690
            Width           =   1665
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "FAX..................:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   46
            Left            =   150
            TabIndex        =   105
            Top             =   1020
            Width           =   1665
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Telefone....:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   44
            Left            =   4440
            TabIndex        =   104
            Top             =   720
            Width           =   1065
         End
      End
   End
   Begin VB.Frame frCNAE 
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      Height          =   2715
      Left            =   2700
      TabIndex        =   214
      Top             =   1575
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox txtQtdeCnae 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3780
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   223
         Top             =   1740
         Width           =   915
      End
      Begin VB.ComboBox cmbCriterio 
         Height          =   315
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   219
         Top             =   1230
         Width           =   6315
      End
      Begin prjChameleon.chameleonButton cmdSelect 
         Height          =   345
         Left            =   4200
         TabIndex        =   215
         ToolTipText     =   "Pesquisar"
         Top             =   2115
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "S&elecionar"
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
         MICON           =   "frmCadMob.frx":33C0
         PICN            =   "frmCadMob.frx":33DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskCodCNAE 
         Height          =   285
         Left            =   1590
         TabIndex        =   216
         Top             =   210
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":3536
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
         Mask            =   "9999-9/99"
         SelText         =   ""
         Text            =   "____-_/__"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdSairCNAE 
         Height          =   345
         Left            =   5490
         TabIndex        =   224
         ToolTipText     =   "Sair da Tela"
         Top             =   2115
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
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
         MICON           =   "frmCadMob.frx":3552
         PICN            =   "frmCadMob.frx":356E
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde....:"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   53
         Left            =   3060
         TabIndex        =   222
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label lblValorCnae 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1650
         TabIndex        =   221
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor em Reais.:"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   52
         Left            =   390
         TabIndex        =   220
         Top             =   1800
         Width           =   1275
      End
      Begin VB.Label lblDescCNAE 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCadMob.frx":35DC
         ForeColor       =   &H0080FFFF&
         Height          =   465
         Left            =   360
         TabIndex        =   218
         Top             =   630
         Width           =   6405
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Código CNAE..:"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   51
         Left            =   360
         TabIndex        =   217
         Top             =   240
         Width           =   1275
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00FFFFFF&
         Height          =   2520
         Left            =   30
         Top             =   45
         Width           =   6885
      End
   End
   Begin Tributacao.jcFrames frIE 
      Height          =   3885
      Left            =   6390
      Top             =   1125
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   6853
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextBoxHeight   =   20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   2
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton cmdIEReduz 
         Height          =   240
         Left            =   540
         TabIndex        =   282
         Top             =   90
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   423
         BTYPE           =   14
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
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   13026246
         MPTR            =   99
         MICON           =   "frmCadMob.frx":36AC
         PICN            =   "frmCadMob.frx":39C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdIEExpande 
         Height          =   240
         Left            =   90
         TabIndex        =   283
         Top             =   90
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   423
         BTYPE           =   14
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
         COLTYPE         =   4
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   13026246
         MPTR            =   99
         MICON           =   "frmCadMob.frx":3B20
         PICN            =   "frmCadMob.frx":3E3A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   -1  'True
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin vbAcceleratorSGrid6.vbalGrid grdMain 
         Height          =   3390
         Left            =   45
         TabIndex        =   284
         Top             =   450
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   5980
         NoHorizontalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         HighlightBackColor=   128
         HighlightForeColor=   16777215
         GroupRowForeColor=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ScrollBarStyle  =   1
         DisableIcons    =   -1  'True
         DrawFocusRectangle=   0   'False
         GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   1
      Left            =   2040
      TabIndex        =   71
      Top             =   90
      Width           =   9405
      Begin VB.ComboBox cmbBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   1380
         Width           =   5000
      End
      Begin VB.ComboBox cmbCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1020
         Width           =   5000
      End
      Begin VB.ComboBox cmbUF 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   660
         Width           =   5000
      End
      Begin VB.CommandButton cmdFoto 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Consultar &Foto do Imóvel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3780
         Width           =   2415
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   42
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtNomeLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2490
         MaxLength       =   50
         TabIndex        =   43
         Top             =   1920
         Width           =   5325
      End
      Begin VB.TextBox txtNumero 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   45
         Top             =   2270
         Width           =   975
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   60
         TabIndex        =   46
         Top             =   2270
         Width           =   3735
      End
      Begin VB.TextBox txtHP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   100
         TabIndex        =   48
         Top             =   2970
         Width           =   6405
      End
      Begin esMaskEdit.esMaskedEdit mskCEP 
         Height          =   285
         Left            =   1440
         TabIndex        =   47
         Top             =   2620
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":3F94
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
      Begin VB.ListBox lstNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   2370
         ItemData        =   "frmCadMob.frx":3FB0
         Left            =   2480
         List            =   "frmCadMob.frx":3FB2
         TabIndex        =   44
         Top             =   1920
         Visible         =   0   'False
         Width           =   5370
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Seq.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   7
         Left            =   4890
         TabIndex        =   93
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblSeq 
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   5340
         TabIndex        =   92
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SubUnidade.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   6
         Left            =   7050
         TabIndex        =   91
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblSubUnid 
         BackStyle       =   0  'Transparent
         Caption         =   "000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   8100
         TabIndex        =   90
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Unidade.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   5
         Left            =   5850
         TabIndex        =   89
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblUnidade 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   6630
         TabIndex        =   88
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   4
         Left            =   3690
         TabIndex        =   87
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lblLote 
         BackStyle       =   0  'Transparent
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   4170
         TabIndex        =   86
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   3
         Left            =   2430
         TabIndex        =   85
         Top             =   240
         Width           =   645
      End
      Begin VB.Label lblQuadra 
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3120
         TabIndex        =   84
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   2
         Left            =   1500
         TabIndex        =   83
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblSetor 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   2070
         TabIndex        =   82
         Top             =   240
         Width           =   285
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito.:"
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   81
         Top             =   240
         Width           =   645
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
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1260
         TabIndex        =   80
         Top             =   240
         Width           =   345
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Index           =   27
         Left            =   2580
         TabIndex        =   79
         Top             =   2310
         Width           =   1365
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   24
         Left            =   180
         TabIndex        =   78
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro.........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   25
         Left            =   210
         TabIndex        =   77
         Top             =   1440
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   26
         Left            =   210
         TabIndex        =   76
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   31
         Left            =   180
         TabIndex        =   75
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número.......:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   33
         Left            =   210
         TabIndex        =   74
         Top             =   2340
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP............:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   28
         Left            =   210
         TabIndex        =   73
         Top             =   2700
         Width           =   1155
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Home Page..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   50
         Left            =   210
         TabIndex        =   72
         Top             =   3090
         Width           =   1185
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   9
      Left            =   2025
      TabIndex        =   244
      Top             =   90
      Width           =   9420
      Begin VB.OptionButton optRec 
         Caption         =   "Recebidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   332
         Top             =   1755
         Width           =   1275
      End
      Begin VB.OptionButton optRec 
         Caption         =   "Emitidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   0
         Left            =   1800
         TabIndex        =   331
         Top             =   1755
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.ComboBox cmbPagto 
         Height          =   315
         ItemData        =   "frmCadMob.frx":3FB4
         Left            =   6165
         List            =   "frmCadMob.frx":3FC1
         Style           =   2  'Dropdown List
         TabIndex        =   330
         Top             =   270
         Width           =   1050
      End
      Begin VB.CheckBox chkHabilitar 
         Caption         =   "Habilitar"
         Height          =   195
         Left            =   315
         TabIndex        =   326
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox txtNomeISS 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   301
         TabStop         =   0   'False
         Top             =   675
         Width           =   6225
      End
      Begin VB.TextBox txtCodIss 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   299
         Top             =   630
         Width           =   870
      End
      Begin VB.ComboBox cmbAnoISS 
         Height          =   315
         ItemData        =   "frmCadMob.frx":3FD8
         Left            =   2160
         List            =   "frmCadMob.frx":3FDA
         Style           =   2  'Dropdown List
         TabIndex        =   247
         Top             =   1080
         Width           =   1365
      End
      Begin VB.Frame frISS 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Alteração da Nota de "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1275
         Left            =   9045
         TabIndex        =   274
         Top             =   270
         Visible         =   0   'False
         Width           =   4110
         Begin VB.CheckBox chkSemMov 
            BackColor       =   &H00FFC0C0&
            Caption         =   "SEM MOVIMENTO"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   270
            TabIndex        =   277
            Top             =   900
            Width           =   1950
         End
         Begin VB.TextBox txtValorNF 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1080
            TabIndex        =   276
            Top             =   405
            Width           =   1365
         End
         Begin prjChameleon.chameleonButton cmdCancelNF 
            Height          =   315
            Left            =   2925
            TabIndex        =   279
            ToolTipText     =   "Cancelar Edição"
            Top             =   810
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
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
            MPTR            =   1
            MICON           =   "frmCadMob.frx":3FDC
            PICN            =   "frmCadMob.frx":3FF8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdGravarNF 
            Height          =   315
            Left            =   2385
            TabIndex        =   278
            ToolTipText     =   "Gravar os Dados"
            Top             =   810
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
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
            MICON           =   "frmCadMob.frx":4152
            PICN            =   "frmCadMob.frx":416E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelNF 
            Height          =   315
            Left            =   3465
            TabIndex        =   280
            ToolTipText     =   "Excluir Registro"
            Top             =   810
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   556
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
            MPTR            =   1
            MICON           =   "frmCadMob.frx":4513
            PICN            =   "frmCadMob.frx":452F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblMesNF 
            Caption         =   "0"
            Height          =   240
            Left            =   3060
            TabIndex        =   281
            Top             =   405
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor...:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   275
            Top             =   450
            Width           =   735
         End
      End
      Begin Tributacao.XP_ProgressBar PBar 
         Height          =   195
         Left            =   1440
         TabIndex        =   327
         Top             =   270
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   16777215
         Scrolling       =   1
         ShowText        =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdPrintIssAno 
         Height          =   315
         Left            =   270
         TabIndex        =   328
         ToolTipText     =   "Imprimir extrato anual"
         Top             =   1395
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Imprimir"
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
         MICON           =   "frmCadMob.frx":45D1
         PICN            =   "frmCadMob.frx":45ED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label11 
         Caption         =   "Guias com pagamento:"
         Height          =   195
         Left            =   4410
         TabIndex        =   329
         Top             =   315
         Width           =   1725
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Código a Consultar..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   300
         Top             =   675
         Width           =   1815
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DEZEMBRO..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   12
         Left            =   135
         TabIndex        =   298
         Top             =   4455
         Width           =   1410
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   13
         Left            =   3015
         TabIndex        =   297
         Top             =   1980
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   14
         Left            =   3015
         TabIndex        =   296
         Top             =   2205
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   15
         Left            =   3015
         TabIndex        =   295
         Top             =   2430
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   16
         Left            =   3015
         TabIndex        =   294
         Top             =   2655
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   17
         Left            =   3015
         TabIndex        =   293
         Top             =   2880
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   18
         Left            =   3015
         TabIndex        =   292
         Top             =   3105
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   19
         Left            =   3015
         TabIndex        =   291
         Top             =   3330
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   20
         Left            =   3015
         TabIndex        =   290
         Top             =   3555
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   21
         Left            =   3015
         TabIndex        =   289
         Top             =   3780
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   22
         Left            =   3015
         TabIndex        =   288
         Top             =   4005
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   23
         Left            =   3015
         TabIndex        =   287
         Top             =   4230
         Width           =   1230
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   24
         Left            =   3015
         TabIndex        =   286
         Top             =   4455
         Width           =   1230
      End
      Begin VB.Label lblTotISSR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3015
         TabIndex        =   285
         Top             =   4680
         Width           =   1230
      End
      Begin VB.Label lblTotISSE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1530
         TabIndex        =   272
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   1530
         TabIndex        =   271
         Top             =   4455
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   11
         Left            =   1530
         TabIndex        =   270
         Top             =   4230
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   10
         Left            =   1530
         TabIndex        =   269
         Top             =   4005
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   9
         Left            =   1530
         TabIndex        =   268
         Top             =   3780
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   8
         Left            =   1530
         TabIndex        =   267
         Top             =   3555
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   7
         Left            =   1530
         TabIndex        =   266
         Top             =   3330
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   6
         Left            =   1530
         TabIndex        =   265
         Top             =   3105
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   5
         Left            =   1530
         TabIndex        =   264
         Top             =   2880
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   1530
         TabIndex        =   263
         Top             =   2655
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   1530
         TabIndex        =   262
         Top             =   2430
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1530
         TabIndex        =   261
         Top             =   2205
         Width           =   1500
      End
      Begin VB.Label lblIss 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   1530
         TabIndex        =   260
         Top             =   1980
         Width           =   1500
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   12
         Left            =   135
         TabIndex        =   259
         Top             =   4680
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NOVEMBRO..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   11
         Left            =   135
         TabIndex        =   258
         Top             =   4230
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "OUTUBRO...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   10
         Left            =   135
         TabIndex        =   257
         Top             =   4005
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SETEMBRO..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   9
         Left            =   135
         TabIndex        =   256
         Top             =   3780
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AGOSTO....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   8
         Left            =   135
         TabIndex        =   255
         Top             =   3555
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JULHO.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   254
         Top             =   3330
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JUNHO.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   253
         Top             =   3105
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MAIO......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   252
         Top             =   2880
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ABRIL.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   251
         Top             =   2655
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MARÇO.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   250
         Top             =   2430
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "FEVEREIRO.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   249
         Top             =   2205
         Width           =   1410
      End
      Begin VB.Label lblMes 
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "JANEIRO...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   248
         Top             =   1980
         Width           =   1410
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano da Declaração..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   246
         Top             =   1125
         Width           =   1860
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   7
      Left            =   2040
      TabIndex        =   193
      Top             =   60
      Width           =   9405
      Begin MSFlexGridLib.MSFlexGrid grdNF 
         Height          =   2235
         Left            =   90
         TabIndex        =   194
         Top             =   465
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   3942
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedCols       =   0
         BackColorSel    =   8388608
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "^Seq     |<Série    |^Núm.Inicial   |^Núm.Final       |<Nº de Autorização           |^Data Autoriz. |<Usuario               "
      End
      Begin MSFlexGridLib.MSFlexGrid grdProc 
         Height          =   1920
         Left            =   90
         TabIndex        =   234
         Top             =   3015
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   3387
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388608
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmCadMob.frx":4747
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas autorizadas"
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
         Index           =   1
         Left            =   135
         TabIndex        =   236
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Processos relacionados - (duplo clique na linha do grid abre a tela do processo)"
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
         Index           =   0
         Left            =   135
         TabIndex        =   235
         Top             =   2790
         Width           =   7755
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   8
      Left            =   2040
      TabIndex        =   195
      Top             =   60
      Width           =   9405
      Begin VB.Frame PnLivro 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   4155
         Left            =   4980
         TabIndex        =   197
         Top             =   300
         Width           =   3345
         Begin VB.ComboBox cmbModelo 
            Height          =   315
            ItemData        =   "frmCadMob.frx":47D5
            Left            =   1650
            List            =   "frmCadMob.frx":47DF
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   1350
            Width           =   1035
         End
         Begin esMaskEdit.esMaskedEdit mskAb 
            Height          =   285
            Left            =   1650
            TabIndex        =   199
            Top             =   1740
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            MouseIcon       =   "frmCadMob.frx":47EB
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
            MaxLength       =   10
            Mask            =   "99/99/9999"
            SelText         =   ""
            Text            =   "__/__/____"
            HideSelection   =   -1  'True
         End
         Begin prjChameleon.chameleonButton cmdExcluir3 
            Height          =   315
            Left            =   2220
            TabIndex        =   200
            ToolTipText     =   "Excluir Livro Fiscal"
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "frmCadMob.frx":4807
            PICN            =   "frmCadMob.frx":4823
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAlterar3 
            Height          =   315
            Left            =   1200
            TabIndex        =   201
            ToolTipText     =   "Alterar Livro Fiscal"
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "frmCadMob.frx":48C5
            PICN            =   "frmCadMob.frx":48E1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdNovo3 
            Height          =   315
            Left            =   180
            TabIndex        =   202
            ToolTipText     =   "Novo Livro Fiscal"
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "frmCadMob.frx":4A3B
            PICN            =   "frmCadMob.frx":4A57
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin esMaskEdit.esMaskedEdit mskEn 
            Height          =   285
            Left            =   1650
            TabIndex        =   203
            Top             =   2130
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   503
            MouseIcon       =   "frmCadMob.frx":4BB1
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
            MaxLength       =   10
            Mask            =   "99/99/9999"
            SelText         =   ""
            Text            =   "__/__/____"
            HideSelection   =   -1  'True
         End
         Begin prjChameleon.chameleonButton cmdGravar3 
            Height          =   315
            Left            =   1230
            TabIndex        =   204
            ToolTipText     =   "Gravar Livro"
            Top             =   3600
            Width           =   975
            _ExtentX        =   1720
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
            MICON           =   "frmCadMob.frx":4BCD
            PICN            =   "frmCadMob.frx":4BE9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdCancelar3 
            Height          =   315
            Left            =   2250
            TabIndex        =   205
            ToolTipText     =   "Cancelar Edição"
            Top             =   3630
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "&Cancel"
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
            MICON           =   "frmCadMob.frx":4F8E
            PICN            =   "frmCadMob.frx":4FAA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblNovo2 
            Caption         =   "0"
            Height          =   225
            Left            =   600
            TabIndex        =   211
            Top             =   3150
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Encerramento:"
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
            Left            =   270
            TabIndex        =   210
            Top             =   2190
            Width           =   1335
         End
         Begin VB.Shape Shape2 
            BorderColor     =   &H00C00000&
            BorderWidth     =   3
            FillColor       =   &H00C00000&
            Height          =   4125
            Left            =   0
            Top             =   0
            Width           =   3345
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Sequência....:"
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
            Left            =   270
            TabIndex        =   209
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Modelo.........:"
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
            Left            =   270
            TabIndex        =   208
            Top             =   1425
            Width           =   1275
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Abertura........:"
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
            Left            =   270
            TabIndex        =   207
            Top             =   1815
            Width           =   1335
         End
         Begin VB.Label lblSeq2 
            BackStyle       =   0  'Transparent
            Height          =   225
            Left            =   1350
            TabIndex        =   206
            Top             =   1050
            Width           =   555
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdLivro 
         Height          =   4125
         Left            =   120
         TabIndex        =   196
         Top             =   300
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   7276
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388608
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "^Seq     |^Modelo    |^Data Abertura        |^Data Encerramento "
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   0
      Left            =   2025
      TabIndex        =   59
      Top             =   30
      Width           =   9405
      Begin VB.CheckBox chkLiberadoVRE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Liberado VRE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   7710
         TabIndex        =   356
         ToolTipText     =   "Os dados da empresa já foram liberados"
         Top             =   3780
         Width           =   1515
      End
      Begin VB.CheckBox chkDanfe 
         Caption         =   "Check1"
         Height          =   195
         Left            =   8700
         TabIndex        =   355
         Top             =   5430
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkMEI 
         Caption         =   "Check1"
         Height          =   195
         Left            =   8130
         TabIndex        =   354
         Top             =   5370
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkRE 
         Caption         =   "Check1"
         Height          =   195
         Left            =   8670
         TabIndex        =   353
         Top             =   5700
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkEmiteNF 
         Caption         =   "Check1"
         Height          =   195
         Left            =   8130
         TabIndex        =   352
         Top             =   5670
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.CheckBox chkEmpInd 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Empresa Individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   348
         Top             =   4230
         Width           =   2055
      End
      Begin VB.CheckBox chkSubstitutoTributario 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Substituto tributário do ISSQN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   3300
         Width           =   3165
      End
      Begin Tributacao.jcFrames frMei 
         Height          =   330
         Left            =   90
         Top             =   4605
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         FrameColor      =   192
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
         Begin prjChameleon.chameleonButton btDataMei 
            Height          =   255
            Left            =   1620
            TabIndex        =   346
            ToolTipText     =   "Consultar períodos do Simples Nacional"
            Top             =   45
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   450
            BTYPE           =   5
            TX              =   "..."
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
            BCOL            =   -2147483633
            BCOLO           =   -2147483633
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCadMob.frx":5104
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
            Alignment       =   2  'Center
            Caption         =   "Optante do MEI"
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
            Height          =   195
            Index           =   8
            Left            =   90
            TabIndex        =   345
            Top             =   60
            Width           =   1455
         End
      End
      Begin VB.TextBox txtSIL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4080
         MaxLength       =   80
         TabIndex        =   33
         Top             =   2925
         Visible         =   0   'False
         Width           =   5145
      End
      Begin VB.TextBox txtImovel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8160
         MaxLength       =   5
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1065
      End
      Begin VB.CheckBox chkBombon 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Bombonieri"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   5535
         TabIndex        =   26
         Top             =   2565
         Width           =   1455
      End
      Begin VB.CheckBox chkIsentoTaxa 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Isento Taxa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   1530
         TabIndex        =   31
         Top             =   3930
         Width           =   1410
      End
      Begin VB.CheckBox chk24horas 
         BackColor       =   &H00EEEEEE&
         Caption         =   "24 horas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   4140
         TabIndex        =   25
         Top             =   2565
         Width           =   1185
      End
      Begin VB.CheckBox chkInscTemp 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Inscrição Temporária"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2990
         Width           =   2805
      End
      Begin VB.Frame Frame1 
         Height          =   960
         Left            =   3465
         TabIndex        =   302
         Top             =   4050
         Width           =   3975
         Begin VB.TextBox txtNumProcIE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2850
            MaxLength       =   15
            TabIndex        =   38
            Top             =   540
            Width           =   1065
         End
         Begin VB.CheckBox chkIE 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Dispensado do ISS Eletrônico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   195
            Left            =   90
            TabIndex        =   36
            ToolTipText     =   "Obrigatoriedade da Declaração pelo ISS Eletrônico"
            Top             =   180
            Width           =   2955
         End
         Begin esMaskEdit.esMaskedEdit mskDataIE 
            Height          =   285
            Left            =   690
            TabIndex        =   37
            Top             =   540
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   503
            MouseIcon       =   "frmCadMob.frx":5120
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
            MaxLength       =   10
            Mask            =   "99/99/9999"
            SelText         =   ""
            Text            =   "__/__/____"
            HideSelection   =   -1  'True
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Data.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   59
            Left            =   90
            TabIndex        =   304
            Top             =   585
            Width           =   585
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Processo.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   240
            Index           =   4
            Left            =   1860
            TabIndex        =   303
            Top             =   585
            Width           =   900
         End
      End
      Begin VB.TextBox txtOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8415
         MaxLength       =   25
         TabIndex        =   17
         Top             =   1260
         Width           =   810
      End
      Begin VB.CheckBox chkIsentoISS 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Isento ISS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   3920
         Width           =   1275
      End
      Begin VB.CheckBox chkAlvara 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Alvará Automático"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   120
         TabIndex        =   29
         Top             =   3610
         Width           =   2175
      End
      Begin VB.ComboBox cmbHorario 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1995
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2505
         Width           =   1995
      End
      Begin VB.TextBox txtNumProcE 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4995
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2100
         Width           =   1425
      End
      Begin VB.TextBox txtNumProcA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4995
         MaxLength       =   15
         TabIndex        =   19
         Top             =   1740
         Width           =   1425
      End
      Begin VB.TextBox txtInscEst 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1260
         Width           =   1995
      End
      Begin VB.TextBox txtFantasia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         MaxLength       =   60
         TabIndex        =   14
         Top             =   900
         Width           =   7245
      End
      Begin VB.TextBox txtRazao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         MaxLength       =   200
         TabIndex        =   13
         Top             =   540
         Width           =   7245
      End
      Begin VB.TextBox txtCodEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1980
         MaxLength       =   9
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   1065
      End
      Begin esMaskEdit.esMaskedEdit mskDataAb 
         Height          =   285
         Left            =   1980
         TabIndex        =   18
         Top             =   1740
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":513C
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
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskDataEn 
         Height          =   285
         Left            =   1980
         TabIndex        =   21
         Top             =   2100
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":5158
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
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskDataPAb 
         Height          =   285
         Left            =   8025
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         BackColor       =   15658734
         MouseIcon       =   "frmCadMob.frx":5174
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
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Locked          =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskDataPEn 
         Height          =   285
         Left            =   8025
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         BackColor       =   15658734
         MouseIcon       =   "frmCadMob.frx":5190
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
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Locked          =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdGerar 
         Height          =   345
         Left            =   5760
         TabIndex        =   34
         ToolTipText     =   "Consultar períodos do Simples Nacional"
         Top             =   3255
         Width           =   375
         _ExtentX        =   661
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
         MICON           =   "frmCadMob.frx":51AC
         PICN            =   "frmCadMob.frx":51C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdISSEletro 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   5760
         TabIndex        =   35
         ToolTipText     =   "Consultar períodos do Iss Eletrônico"
         Top             =   3675
         Width           =   375
         _ExtentX        =   661
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
         MICON           =   "frmCadMob.frx":5322
         PICN            =   "frmCadMob.frx":533E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskCPF 
         Height          =   285
         Left            =   3750
         TabIndex        =   10
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":5498
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
         Left            =   5820
         TabIndex        =   11
         Top             =   180
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":54B4
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
         Mask            =   "99999999999999"
         SelText         =   ""
         Text            =   "______________"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskRG 
         Height          =   285
         Left            =   5280
         TabIndex        =   16
         Top             =   1260
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":54D0
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
      Begin esMaskEdit.esMaskedEdit mskDataAP 
         Height          =   285
         Left            =   7860
         TabIndex        =   32
         Top             =   3300
         Visible         =   0   'False
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":54EC
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
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdAP 
         Height          =   360
         Left            =   8850
         TabIndex        =   307
         ToolTipText     =   "Imprimir alvará provisório"
         Top             =   3270
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   635
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
         MPTR            =   1
         MICON           =   "frmCadMob.frx":5508
         PICN            =   "frmCadMob.frx":5524
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSil 
         Height          =   285
         Left            =   6990
         TabIndex        =   351
         ToolTipText     =   "Cadastrar SIL"
         Top             =   2520
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "SIL"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":567E
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
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "SIL..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   0
         Left            =   3510
         TabIndex        =   335
         Top             =   2970
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Imóvel..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   63
         Left            =   7335
         TabIndex        =   334
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Alvará Provisório:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   60
         Left            =   6225
         TabIndex        =   306
         Top             =   3345
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de RG..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   13
         Left            =   4200
         TabIndex        =   243
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Orgão..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   14
         Left            =   7560
         TabIndex        =   242
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   3
         Left            =   5115
         TabIndex        =   241
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "CPF..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   6
         Left            =   3150
         TabIndex        =   240
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ISS Eletrônico......:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   58
         Left            =   3510
         TabIndex        =   238
         Top             =   3810
         Width           =   1755
      End
      Begin VB.Label lblISSEletro 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
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
         Height          =   225
         Left            =   5280
         TabIndex        =   237
         Top             =   3780
         Width           =   435
      End
      Begin VB.Label lblSN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
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
         Height          =   225
         Left            =   5280
         TabIndex        =   232
         Top             =   3315
         Width           =   435
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Simples Nacional..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   56
         Left            =   3510
         TabIndex        =   231
         Top             =   3330
         Width           =   1755
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Processo.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   12
         Left            =   6540
         TabIndex        =   70
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Processo.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   11
         Left            =   6540
         TabIndex        =   69
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   10
         Left            =   3420
         TabIndex        =   68
         Top             =   2160
         Width           =   1485
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   9
         Left            =   3420
         TabIndex        =   67
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Data Encerramento.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   8
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   1845
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Abertura....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   65
         Top             =   1800
         Width           =   1845
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Estadual..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   64
         Top             =   1305
         Width           =   1845
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Nome Fantasia.......:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   945
         Width           =   1845
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Razão Social..........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   62
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000040&
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição Municipal.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   1845
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Horário Funcionam..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   19
         Left            =   120
         TabIndex        =   60
         Top             =   2550
         Width           =   1845
      End
   End
   Begin VB.Frame Tela 
      Height          =   5160
      Index           =   11
      Left            =   2010
      TabIndex        =   310
      Top             =   90
      Width           =   9405
      Begin Tributacao.jcFrames jcFrames2 
         Height          =   1950
         Left            =   4140
         Top             =   2340
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3440
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
         Begin VB.TextBox txtNumProc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3810
            TabIndex        =   319
            Top             =   540
            Width           =   1185
         End
         Begin VB.ComboBox cmbAno 
            Height          =   315
            Left            =   1575
            Style           =   2  'Dropdown List
            TabIndex        =   318
            Top             =   540
            Width           =   1050
         End
         Begin VB.TextBox txtArq 
            Height          =   285
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   320
            Top             =   945
            Width           =   2940
         End
         Begin VB.ComboBox cmbTipoDoc 
            Height          =   315
            ItemData        =   "frmCadMob.frx":569A
            Left            =   1575
            List            =   "frmCadMob.frx":569C
            Style           =   2  'Dropdown List
            TabIndex        =   317
            Top             =   135
            Width           =   3435
         End
         Begin prjChameleon.chameleonButton cmdOpenPic 
            Height          =   285
            Left            =   4590
            TabIndex        =   321
            ToolTipText     =   "Localiza documento a importar"
            Top             =   945
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
            MICON           =   "frmCadMob.frx":569E
            PICN            =   "frmCadMob.frx":56BA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdGravarPic 
            Height          =   315
            Left            =   3960
            TabIndex        =   324
            ToolTipText     =   "Gravar o documento"
            Top             =   1440
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
            MICON           =   "frmCadMob.frx":5741
            PICN            =   "frmCadMob.frx":575D
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
            Caption         =   "Nº Processo.:"
            Height          =   225
            Index           =   62
            Left            =   2745
            TabIndex        =   325
            Top             =   570
            Width           =   1080
         End
         Begin VB.Label Label9 
            Caption         =   "Arquivo Importar...:"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   323
            Top             =   990
            Width           =   1365
         End
         Begin VB.Label Label9 
            Caption         =   "Ano Documento...:"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   322
            Top             =   585
            Width           =   1365
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo Documento..:"
            Height          =   195
            Left            =   135
            TabIndex        =   316
            Top             =   180
            Width           =   1320
         End
      End
      Begin prjChameleon.chameleonButton cmdD1 
         Height          =   285
         Left            =   5040
         TabIndex        =   312
         ToolTipText     =   "Documento anterior"
         Top             =   585
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
         MICON           =   "frmCadMob.frx":5B02
         PICN            =   "frmCadMob.frx":5B1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdD2 
         Height          =   285
         Left            =   7515
         TabIndex        =   313
         ToolTipText     =   "Próximo documento"
         Top             =   585
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
         MICON           =   "frmCadMob.frx":5C78
         PICN            =   "frmCadMob.frx":5C94
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAbrirPic 
         Height          =   540
         Left            =   5625
         TabIndex        =   315
         ToolTipText     =   "Abrir documento"
         Top             =   1395
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   953
         BTYPE           =   3
         TX              =   " &Abrir"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCadMob.frx":5DEE
         PICN            =   "frmCadMob.frx":5E0A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExcluirDoc 
         Height          =   315
         Left            =   7605
         TabIndex        =   333
         ToolTipText     =   "Exclusão de documento"
         Top             =   1620
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Excluir"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":5FB9
         PICN            =   "frmCadMob.frx":5FD5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image imgTmp2 
         Height          =   495
         Left            =   8595
         Picture         =   "frmCadMob.frx":6077
         Stretch         =   -1  'True
         Top             =   4500
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label lblTipoDoc 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   314
         Top             =   945
         Width           =   2940
      End
      Begin VB.Label lblPagDoc 
         Alignment       =   2  'Center
         Caption         =   "Documento 0 de 0"
         Height          =   240
         Left            =   5490
         TabIndex        =   311
         Top             =   630
         Width           =   1995
      End
      Begin VB.Image imgTmp 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   4650
         Left            =   315
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3660
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   6
      Left            =   2040
      TabIndex        =   172
      Top             =   90
      Width           =   9405
      Begin VB.Frame Frame7 
         Caption         =   "Veículos"
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
         Height          =   2115
         Left            =   7245
         TabIndex        =   337
         Top             =   1440
         Width           =   2040
         Begin VB.TextBox txtPonto 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   90
            MaxLength       =   40
            TabIndex        =   349
            Top             =   1740
            Width           =   1845
         End
         Begin esMaskEdit.esMaskedEdit mskPlaca 
            Height          =   285
            Left            =   180
            TabIndex        =   339
            Top             =   1215
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   503
            MouseIcon       =   "frmCadMob.frx":C2BF
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
            MaxLength       =   8
            Mask            =   "???-9999"
            SelText         =   ""
            Text            =   "___-____"
            HideSelection   =   -1  'True
         End
         Begin VB.ListBox lstPlaca 
            Appearance      =   0  'Flat
            Height          =   810
            Left            =   180
            TabIndex        =   338
            Top             =   315
            Width           =   1680
         End
         Begin prjChameleon.chameleonButton btAddPlaca 
            Height          =   285
            Left            =   1125
            TabIndex        =   340
            ToolTipText     =   "Selecionar um Responsável"
            Top             =   1215
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
            MICON           =   "frmCadMob.frx":C2DB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton btDelPlaca 
            Height          =   285
            Left            =   1485
            TabIndex        =   341
            ToolTipText     =   "Selecionar um Responsável"
            Top             =   1215
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
            MICON           =   "frmCadMob.frx":C2F7
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
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Ponto/Agência"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   64
            Left            =   90
            TabIndex        =   350
            Top             =   1530
            Width           =   1515
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Escritório Contabil Responsável"
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
         Height          =   2055
         Left            =   90
         TabIndex        =   181
         Top             =   1440
         Width           =   7080
         Begin VB.TextBox txtEmailCont 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   228
            Top             =   1140
            Width           =   5325
         End
         Begin VB.TextBox txtFoneCont 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3195
            Locked          =   -1  'True
            MaxLength       =   200
            TabIndex        =   226
            Top             =   360
            Width           =   2235
         End
         Begin VB.ComboBox cmbNomeEsc 
            Height          =   315
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   183
            Top             =   750
            Width           =   5895
         End
         Begin VB.TextBox txtCodEsc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   990
            MaxLength       =   6
            TabIndex        =   182
            Top             =   360
            Width           =   855
         End
         Begin prjChameleon.chameleonButton cmdEmail 
            Height          =   345
            Left            =   6390
            TabIndex        =   230
            ToolTipText     =   "Abrir Email"
            Top             =   1125
            Width           =   495
            _ExtentX        =   873
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
            MICON           =   "frmCadMob.frx":C313
            PICN            =   "frmCadMob.frx":C32F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAbrir 
            Height          =   330
            Left            =   5715
            TabIndex        =   336
            ToolTipText     =   "Selecionar um Responsável"
            Top             =   315
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   582
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
            MICON           =   "frmCadMob.frx":C3C2
            PICN            =   "frmCadMob.frx":C3DE
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
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email....:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   55
            Left            =   150
            TabIndex        =   229
            Top             =   1200
            Width           =   795
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fone..:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   54
            Left            =   2475
            TabIndex        =   227
            Top             =   405
            Width           =   660
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nome....:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   18
            Left            =   150
            TabIndex        =   185
            Top             =   810
            Width           =   795
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código..:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   13
            Left            =   150
            TabIndex        =   184
            Top             =   420
            Width           =   795
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Profissional Responsável"
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
         Height          =   1245
         Left            =   90
         TabIndex        =   173
         Top             =   180
         Width           =   9195
         Begin VB.TextBox txtTipoConselho 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   50
            TabIndex        =   176
            Top             =   780
            Width           =   2985
         End
         Begin VB.TextBox txtNumRegistro 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6300
            MaxLength       =   15
            TabIndex        =   175
            Top             =   780
            Width           =   2260
         End
         Begin VB.TextBox txtNomeProf 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1770
            MaxLength       =   40
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   390
            Width           =   6795
         End
         Begin prjChameleon.chameleonButton cmdAddResp 
            Height          =   285
            Left            =   8670
            TabIndex        =   177
            ToolTipText     =   "Selecionar um Responsável"
            Top             =   390
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
            MICON           =   "frmCadMob.frx":C538
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDelResp 
            Height          =   285
            Left            =   8640
            TabIndex        =   347
            ToolTipText     =   "Limpar campo do Responsável"
            Top             =   765
            Width           =   360
            _ExtentX        =   635
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
            MICON           =   "frmCadMob.frx":C554
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
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Conselho:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   22
            Left            =   120
            TabIndex        =   180
            Top             =   810
            Width           =   1635
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº Registro..:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   23
            Left            =   5010
            TabIndex        =   179
            Top             =   810
            Width           =   1155
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Código e Nome...:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   315
            Index           =   21
            Left            =   120
            TabIndex        =   178
            Top             =   450
            Width           =   1635
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Diversos"
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
         Height          =   1485
         Left            =   90
         TabIndex        =   186
         Top             =   3495
         Width           =   9195
         Begin VB.TextBox txtSenha 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7470
            MaxLength       =   15
            TabIndex        =   190
            Top             =   270
            Width           =   1425
         End
         Begin VB.TextBox txtHorarioExt 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   1485
            MaxLength       =   300
            MultiLine       =   -1  'True
            TabIndex        =   192
            Top             =   675
            Width           =   7455
         End
         Begin VB.TextBox txtNumFunc 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2040
            MaxLength       =   6
            TabIndex        =   187
            Top             =   270
            Width           =   810
         End
         Begin VB.TextBox txtCapital 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4515
            MaxLength       =   15
            TabIndex        =   189
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Senha ISS Elet.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   61
            Left            =   5985
            TabIndex        =   308
            Top             =   330
            Width           =   1470
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Horário p/Extenso..:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   465
            Index           =   57
            Left            =   195
            TabIndex        =   233
            Top             =   735
            Width           =   1215
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nº de Funcionários.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   35
            Left            =   180
            TabIndex        =   191
            Top             =   315
            Width           =   1755
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Capital Social...:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   20
            Left            =   2985
            TabIndex        =   188
            Top             =   330
            Width           =   1515
         End
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   5
      Left            =   2040
      TabIndex        =   168
      Top             =   60
      Width           =   9405
      Begin VB.TextBox txtHist 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   795
         Left            =   90
         MaxLength       =   5000
         MultiLine       =   -1  'True
         TabIndex        =   169
         Top             =   3720
         Width           =   9225
      End
      Begin MSFlexGridLib.MSFlexGrid grdHist 
         Height          =   3345
         Left            =   60
         TabIndex        =   170
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388608
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   $"frmCadMob.frx":C570
      End
      Begin prjChameleon.chameleonButton cmdEditHist 
         Height          =   315
         Left            =   120
         TabIndex        =   171
         ToolTipText     =   "Editar Histórico"
         Top             =   4695
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
         MICON           =   "frmCadMob.frx":C617
         PICN            =   "frmCadMob.frx":C633
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
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   4
      Left            =   2040
      TabIndex        =   134
      Top             =   60
      Width           =   9405
      Begin prjChameleon.chameleonButton cmdAddCnae1 
         Height          =   300
         Left            =   4260
         TabIndex        =   153
         ToolTipText     =   "Adicionar/alterar CNAE Principal"
         Top             =   4740
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":C78D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cmbCnae 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   156
         Top             =   4740
         Width           =   1425
      End
      Begin esMaskEdit.esMaskedEdit mskCnae 
         Height          =   285
         Left            =   2910
         TabIndex        =   152
         Top             =   4740
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         MouseIcon       =   "frmCadMob.frx":C7A9
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
         Mask            =   "9999-9/99"
         SelText         =   ""
         Text            =   "____-_/__"
         HideSelection   =   -1  'True
      End
      Begin VB.TextBox txtValorAliq 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   141
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txtAtiv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2070
         Locked          =   -1  'True
         MaxLength       =   300
         TabIndex        =   140
         Top             =   540
         Width           =   6645
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5595
         MaxLength       =   12
         TabIndex        =   139
         Top             =   885
         Width           =   1245
      End
      Begin VB.TextBox txtAtivExt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2070
         MaxLength       =   1000
         TabIndex        =   136
         Top             =   1230
         Width           =   6645
      End
      Begin VB.TextBox txtQtde 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7950
         MaxLength       =   12
         TabIndex        =   135
         Top             =   885
         Width           =   765
      End
      Begin VB.CheckBox chkVistoria 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Vistoria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   150
         Top             =   4770
         Width           =   1005
      End
      Begin MSFlexGridLib.MSFlexGrid grdAtiv 
         Height          =   1215
         Left            =   90
         TabIndex        =   142
         Top             =   1920
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedCols       =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmCadMob.frx":C7C5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdVS 
         Height          =   1215
         Left            =   90
         TabIndex        =   145
         Top             =   3420
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmCadMob.frx":C86F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid grdVS2 
         Height          =   1215
         Left            =   90
         TabIndex        =   225
         Top             =   3420
         Visible         =   0   'False
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   2143
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   8388608
         ForeColorSel    =   16777215
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmCadMob.frx":C916
      End
      Begin prjChameleon.chameleonButton cmdAddAtiv 
         Height          =   300
         Left            =   8820
         TabIndex        =   137
         ToolTipText     =   "Adicionar Atividade Principal"
         Top             =   510
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":C9B8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAddISS 
         Height          =   300
         Left            =   8880
         TabIndex        =   143
         Top             =   2220
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":C9D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAddVS 
         Height          =   300
         Left            =   8880
         TabIndex        =   147
         Top             =   3900
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":C9F0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAddCnae2 
         Height          =   300
         Left            =   7980
         TabIndex        =   157
         ToolTipText     =   "Adicionar  outros códigos CNAE"
         Top             =   4740
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CA0C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelISS 
         Height          =   300
         Left            =   8880
         TabIndex        =   144
         Top             =   2580
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CA28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelVS 
         Height          =   300
         Left            =   8880
         TabIndex        =   148
         Top             =   4260
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CA44
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelCnae 
         Height          =   300
         Left            =   8340
         TabIndex        =   158
         ToolTipText     =   "Remover código CNAE"
         Top             =   4740
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CA60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelCnaeP 
         Height          =   300
         Left            =   4620
         TabIndex        =   154
         ToolTipText     =   "Remover CNAE Principal"
         Top             =   4740
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CA7C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdVSOld 
         Height          =   300
         Left            =   8880
         TabIndex        =   146
         ToolTipText     =   "Exibir as Atividades antigas de VS"
         Top             =   3540
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   ">"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CA98
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdEdtCnae 
         Height          =   300
         Left            =   8700
         TabIndex        =   159
         ToolTipText     =   "Alterar código CNAE"
         Top             =   4740
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CAB4
         PICN            =   "frmCadMob.frx":CAD0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOutrasAtiv 
         Height          =   300
         Left            =   8820
         TabIndex        =   138
         ToolTipText     =   "Outras Atividades"
         Top             =   870
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   529
         BTYPE           =   3
         TX              =   ">"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CC2A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btAtivExtenso 
         Height          =   270
         Left            =   8760
         TabIndex        =   357
         ToolTipText     =   "Preencher a atividade por Extenso a partir dos Cnaes cadastrados."
         Top             =   1240
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "Cnae"
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
         FCOL            =   128
         FCOLO           =   128
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCadMob.frx":CC46
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
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CNAE outros..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   38
         Left            =   5100
         TabIndex        =   213
         Top             =   4830
         Width           =   1365
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CNAE principal..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   37
         Left            =   1380
         TabIndex        =   212
         Top             =   4800
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   3315
         TabIndex        =   167
         Top             =   930
         Width           =   570
      End
      Begin VB.Label lblAliq 
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
         Height          =   285
         Left            =   3960
         TabIndex        =   166
         Top             =   930
         Width           =   315
      End
      Begin VB.Label lblTipoISS 
         BackStyle       =   0  'Transparent
         Caption         =   "-->"
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
         Left            =   3150
         TabIndex        =   165
         Top             =   1650
         Visible         =   0   'False
         Width           =   4035
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Área m².....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   49
         Left            =   4515
         TabIndex        =   164
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Aliquota......:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   43
         Left            =   120
         TabIndex        =   163
         Top             =   930
         Width           =   2000
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Atividades para pagamento de ISS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   42
         Left            =   90
         TabIndex        =   162
         Top             =   1650
         Width           =   3705
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade Principal....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   16
         Left            =   120
         TabIndex        =   161
         Top             =   600
         Width           =   2000
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade p/Extenso..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   17
         Left            =   120
         TabIndex        =   160
         Top             =   1290
         Width           =   2000
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade Principal / Taxa de Licença de Funcionamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Index           =   14
         Left            =   90
         TabIndex        =   155
         Top             =   180
         Width           =   6105
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BackStyle       =   0  'Transparent
         Caption         =   "Vigilância Sanitária"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   15
         Left            =   90
         TabIndex        =   151
         Top             =   3180
         Width           =   7755
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   285
         Index           =   32
         Left            =   7215
         TabIndex        =   149
         Top             =   930
         Width           =   660
      End
   End
   Begin VB.Frame Tela 
      BackColor       =   &H00EEEEEE&
      Height          =   5160
      Index           =   3
      Left            =   2040
      TabIndex        =   114
      Top             =   90
      Width           =   9405
      Begin VB.Frame Frame8 
         Caption         =   "Tipo de Endereço"
         ForeColor       =   &H00000080&
         Height          =   1005
         Left            =   6255
         TabIndex        =   342
         Top             =   450
         Visible         =   0   'False
         Width           =   2265
         Begin VB.OptionButton OptEE 
            Caption         =   "Entrega de Carnê"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   344
            Top             =   630
            Width           =   1995
         End
         Begin VB.OptionButton OptEE 
            Caption         =   "Endereço de Entrega"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   343
            Top             =   315
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.ListBox lstEENomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   1785
         ItemData        =   "frmCadMob.frx":CC62
         Left            =   2690
         List            =   "frmCadMob.frx":CC69
         TabIndex        =   126
         Top             =   2460
         Visible         =   0   'False
         Width           =   5400
      End
      Begin VB.CheckBox chkEnd 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Utilizar a localização do imóvel como endereço de entrega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   195
         Left            =   330
         TabIndex        =   125
         Top             =   420
         Width           =   5445
      End
      Begin VB.CommandButton cmdRefreshBairro 
         BackColor       =   &H00EEEEEE&
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Atualizar lista de Bairros"
         Top             =   1680
         Width           =   315
      End
      Begin VB.CommandButton cmdRefreshCity 
         BackColor       =   &H00EEEEEE&
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5490
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Atualizar lista de Cidades"
         Top             =   1320
         Width           =   315
      End
      Begin VB.ComboBox cmbEEBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   122
         Top             =   1680
         Width           =   3915
      End
      Begin VB.ComboBox cmbEECidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   1320
         Width           =   3915
      End
      Begin VB.ComboBox cmbEEUf 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   119
         Top             =   960
         Width           =   3915
      End
      Begin VB.TextBox txtEECompl 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   118
         Top             =   3180
         Width           =   6375
      End
      Begin VB.TextBox txtEENumero 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   117
         Top             =   2820
         Width           =   945
      End
      Begin VB.TextBox txtEENomeLogr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2685
         MaxLength       =   50
         TabIndex        =   116
         Top             =   2460
         Width           =   5385
      End
      Begin VB.TextBox txtEECodLogr 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   115
         Top             =   2460
         Width           =   945
      End
      Begin esMaskEdit.esMaskedEdit mskEECep 
         Height          =   300
         Left            =   1680
         TabIndex        =   120
         Top             =   3555
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   529
         MouseIcon       =   "frmCadMob.frx":CC7B
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
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "CEP..............:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   41
         Left            =   300
         TabIndex        =   133
         Top             =   3630
         Width           =   1400
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   40
         Left            =   270
         TabIndex        =   132
         Top             =   3270
         Width           =   1400
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Número..........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   39
         Left            =   270
         TabIndex        =   131
         Top             =   2880
         Width           =   1400
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade.........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   36
         Left            =   270
         TabIndex        =   130
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro.....:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   34
         Left            =   270
         TabIndex        =   129
         Top             =   2520
         Width           =   1400
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro..........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   30
         Left            =   270
         TabIndex        =   128
         Top             =   1740
         Width           =   1245
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00808000&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado.........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Index           =   29
         Left            =   270
         TabIndex        =   127
         Top             =   1020
         Width           =   1245
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "Print"
      Visible         =   0   'False
      Begin VB.Menu mnuFicha 
         Caption         =   "Ficha Cadastral"
      End
      Begin VB.Menu mnuDECA 
         Caption         =   "DECA - Frente"
      End
      Begin VB.Menu mnuDecaV 
         Caption         =   "DECA - Verso"
      End
   End
End
Attribute VB_Name = "frmCadMob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuISS As cPopupMenu
Attribute m_cMenuISS.VB_VarHelpID = -1

Private Type tPic
    sTipo As String
    sTipoExt As String
    sAno As String
    sArq As String
    sExt As String
End Type

Private Type tTipoDoc
    nCod As Integer
    sNome As String
End Type

Private Type tLOG
    sDataAb As String
    sDataEn As String
    sNumProcAb As String
    sNumProcEn As String
    sDataProcAb As String
    sDataProcEn As String
End Type

Private Type NOTAS
    IdentificaPrestador As String
    TipoPrestador As Integer
    TipoNota As Integer
    NumeroNota As Double
    Serie As String
    DataEmissao As String
    MesRef As Integer
    AnoRef As Integer
    StatusNota As Integer
    DataCancel As String
    Natureza As String
    ValorTotal As Double
    ValorServico As Double
    ValorImposto As Double
    Recolhimento As Integer
    Atividade As Integer
    Aliquota As Double
    RazaoPrestador As String
    CidadePrestador As String
    UFPrestador As String
    LocalPrestador As String
    IdentificaTomador As String
    TipoTomador As String
    RazaoTomador As String
    CidadeTomador As String
    UFTomador As String
    LocalTomador As String
    NumGuia As String
    Pago As String
    CNPJTomador As String
    CNPJPrestador As String
End Type

Private Type ISSELETRO
    nAno As Integer
    nMes As Integer
    nValorEmitida As Double
    nValorRecebida As Double
    bSemMovimento As Boolean
End Type

Private Type VALORADICIONAL
    nMes As Integer
    nBaseCalculoE As Double
    nValorNTribE As Double
    nOutrasE As Double
    nBaseCalculoS As Double
    nValorNTribS As Double
    nOutrasS As Double
End Type

Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String, frAtivo As Integer, bExec As Boolean
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean
Dim sCH As String, TipoISS As String, aLog() As tLOG, aNota() As NOTAS
Dim bDispensadoIE As Boolean
Dim bIsentoTaxa As Boolean, nPointer As Integer, bIsentoIss As Boolean
Dim bAlvaraAutomatico As Boolean, aPic() As tPic, aTipoDoc() As tTipoDoc

Private Sub btAddPlaca_Click()
Dim x As Integer, bFind As Boolean
bFind = False
If mskPlaca.ClipText = "" Then
    MsgBox "Digite o nº da placa.", vbCritical, "Erro"
Else
    For x = 0 To lstPlaca.ListCount - 1
        If lstPlaca.List(x) = mskPlaca.Text Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        lstPlaca.AddItem mskPlaca.Text
        LimpaMascara mskPlaca
    Else
        MsgBox "Placa já cadastrada.", vbCritical, "Erro"
    End If
End If
End Sub

Private Sub btAtivExtenso_Click()
Dim Sql As String, RdoAux As rdoResultset, sDescricao As String, aCnae() As String, x As Integer

ReDim aCnae(0)
If Trim(txtAtivExt.Text) <> "" Then
    If MsgBox("Já existe uma atividade por extenso preenchida, se clicar em sim a atividade por extenso será substituida pela descrição a partir dos Cnaes." & vbCrLf & "Você deseja continuar?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
        Exit Sub
    End If
End If

If mskCnae.ClipText = "" Then
    MsgBox "Selecione o Cnae principal.", vbCritical, "Erro"
    Exit Sub
End If

sCnae = mskCnae.ClipText

Sql = "select * from cnae where cnae='" & sCnae & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Cnae não cadastrado. (" & sCnae & ")", vbCritical, "Erro"
    RdoAux.Close
    Exit Sub
End If
RdoAux.Close

aCnae(0) = sCnae
For x = 0 To cmbCnae.ListCount - 1
    ReDim Preserve aCnae(UBound(aCnae) + 1)
    aCnae(UBound(aCnae)) = RetornaNumero(cmbCnae.List(x))
Next

sDescricao = ""
For x = 0 To UBound(aCnae)
    Sql = "select descricao from cnae where cnae='" & aCnae(x) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        sDescricao = sDescricao & RdoAux!descricao & ";"
    Else
        MsgBox "Cnae: " & aCnae(x) & " não cadastrado.", vbCritical, "Erro"
    End If
Next
txtAtivExt.Text = UCase(sDescricao)

End Sub

Private Sub btDelPlaca_Click()
If lstPlaca.ListIndex = -1 Then
    MsgBox "Selecione a placa a ser excluída.", vbCritical, "Erro"
Else
    lstPlaca.RemoveItem (lstPlaca.ListIndex)
End If
End Sub

Private Sub btMenu_Click(Index As Integer)
If Index = 10 Then
    VALORADICIONAL
Else
    AtivaTela (Index)
End If
End Sub



Private Sub chkEnd_Click()
If Not bExec Then Exit Sub
If chkEnd.value = vbChecked Then
   cmbEEUf.Enabled = False
   cmbEEUf.BackColor = Kde
   cmbEEBairro.Enabled = False
   cmbEEBairro.BackColor = Kde
   cmbEECidade.Enabled = False
   cmbEECidade.BackColor = Kde
   cmdRefreshBairro.Enabled = False
   cmdRefreshCity.Enabled = False
   txtEECodLogr.Enabled = False
   txtEECodLogr.BackColor = Kde
   txtEENomeLogr.Enabled = False
   txtEENomeLogr.BackColor = Kde
   txtEENumero.Enabled = False
   txtEENumero.BackColor = Kde
   txtEECompl.Enabled = False
   txtEECompl.BackColor = Kde
   mskEECep.Enabled = False
   mskEECep.BackColor = Kde
Else
   cmbEEUf.Enabled = True
   cmbEEUf.BackColor = Branco
   cmbEEBairro.Enabled = True
   cmbEEBairro.BackColor = Branco
   cmbEECidade.Enabled = True
   cmbEECidade.BackColor = Branco
   cmdRefreshBairro.Enabled = True
   cmdRefreshCity.Enabled = True
   txtEECodLogr.Enabled = True
   txtEECodLogr.BackColor = Branco
   txtEENomeLogr.Enabled = True
   txtEENomeLogr.BackColor = Branco
   txtEENumero.Enabled = True
   txtEENumero.BackColor = Branco
   txtEECompl.Enabled = True
   txtEECompl.BackColor = Branco
   mskEECep.Enabled = True
   mskEECep.BackColor = Branco
End If

End Sub

Private Sub chkHabilitar_Click()
cmbAnoISS_Click
lblIss_Click (1)
End Sub

Private Sub chkIE_Click()
If chkIE.value = vbUnchecked Then
    LimpaMascara mskDataIE
    txtNumProcIE.Text = ""
End If
End Sub

Private Sub cmbAnoISS_Click()
If cmbAnoISS.ListIndex < 0 Then Exit Sub
LimpaISS
lblIss_Click (1)
If chkHabilitar.value = vbChecked Then
    CarregaISS
End If
'CarregaISS
End Sub

Private Sub cmbBairro_GotFocus()
    Me.KeyPreview = False
End Sub

Private Sub cmbBairro_LostFocus()
    Me.KeyPreview = True
End Sub

Private Sub cmbCidade_Click()

If Not bExec Then Exit Sub
If cmbCidade.ListIndex = -1 Or Not bExec Then
    cmbBairro.Clear
    Exit Sub
End If

cmbBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(Right$(cmbUF.Text, 3), 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbBairro.AddItem !DescBairro
       cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbCidade_GotFocus()
If cmbUF.ListIndex = -1 Then cmbCidade.Clear
Me.KeyPreview = False
End Sub

Private Sub cmbCidade_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmbCriterio_Click()
Dim sCnae As String, nDivisao As Integer, nGrupo As Integer, sClasse As String, nClasse As Integer, nSubClasse As Integer
lblValorCnae.Caption = "0,00"
sCnae = RetornaNumero(mskCodCNAE.Text)
nDivisao = Val(Left(sCnae, 2))
nGrupo = Val(Mid(sCnae, 3, 1))
sClasse = Mid(sCnae, 4, 3)
sClasse = Left(sClasse, 1) & Right(sClasse, 1)
nClasse = Val(sClasse)
nSubClasse = Val(Right(sCnae, 2))
If cmbCriterio.ListIndex > -1 Then
'    Sql = "SELECT cnaecriterio.valor From "
'    Sql = Sql & "cnaecriteriodesc INNER JOIN cnaecriterio ON (cnaecriteriodesc.criterio = cnaecriterio.criterio) WHERE CNAE='" & sCnae & "' AND cnaecriterio.CRITERIO=" & cmbCriterio.ItemData(cmbCriterio.ListIndex)
    Sql = "select valor from cnaecriteriodesc where criterio=" & cmbCriterio.ItemData(cmbCriterio.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            lblValorCnae.Caption = Format(!Valor, "#0.0000")
        End If
       .Close
    End With
End If

End Sub

Private Sub cmbEEBairro_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cmbEEBairro_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmbEECidade_Click()
If cmbEECidade.ListIndex = -1 Or Not bExec Then
    cmbEEBairro.Clear
    Exit Sub
End If

cmbEEBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(Right$(cmbEEUf.Text, 3), 2) & "' AND CODCIDADE=" & cmbEECidade.ItemData(cmbEECidade.ListIndex) & " ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbEEBairro.AddItem !DescBairro
       cmbEEBairro.ItemData(cmbEEBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbEECidade_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cmbEECidade_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmbEEUF_Click()

If cmbEEUf.ListIndex = -1 Or Not bExec Then
    cmbEECidade.Clear
    Exit Sub
End If

cmbEECidade.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(Right$(cmbEEUf.Text, 3), 2) & "' ORDER BY DESCCIDADE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbEECidade.AddItem !descCidade
       cmbEECidade.ItemData(cmbEECidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbEEUF_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cmbEEUF_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmbHorario_GotFocus()
Me.KeyPreview = False
End Sub

Private Sub cmbHorario_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmbNomeEsc_Click()

If Not bExec Then Exit Sub
txtFoneCont.Text = ""
txtEmailCont.Text = ""
If cmbNomeEsc.ListIndex > -1 Then
    txtCodEsc.Text = cmbNomeEsc.ItemData(cmbNomeEsc.ListIndex)
    Sql = "SELECT * FROM ESCRITORIOCONTABIL WHERE CODIGOESC=" & Val(txtCodEsc.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount >= 0 Then
            txtFoneCont.Text = SubNull(!telefone)
            txtEmailCont.Text = SubNull(!Email)
        End If
       .Close
    End With
Else
   txtCodEsc.Text = 0
End If

End Sub



Private Sub cmbPagto_Click()
If chkHabilitar.value = vbChecked Then
    CarregaISS
End If

End Sub

Private Sub cmbUF_Click()
If Not bExec Then Exit Sub
If cmbUF.ListIndex = -1 Or Not bExec Then
    cmbCidade.Clear
    Exit Sub
End If

cmbCidade.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(Right$(cmbUF.Text, 3), 2) & "' ORDER BY DESCCIDADE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbCidade.AddItem !descCidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With
End Sub

Private Sub cmbUF_GotFocus()
Me.KeyPreview = False
End Sub


Private Sub cmbUF_LostFocus()
Me.KeyPreview = True
End Sub

Private Sub cmdAbrir_Click()
frmEscContab.show
frmEscContab.ZOrder 0
End Sub

Private Sub cmdAbrirPic_Click()
Dim ret As Long, sPathOrigem As String, sAno As String, sFile As String, sFullArq As String
If Val(txtCodEmpresa.Text) = 0 Then Exit Sub
If nPointer > 0 Then
    sPathOrigem = "\\192.168.200.130\atualizagti\Documentos\"
    sAno = aPic(nPointer).sAno
    sFile = aPic(nPointer).sArq
    sFullArq = sPathOrigem & sAno & "\" & sFile
    Call ShellExecute(0&, vbNullString, sFullArq, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox "Não existe arquivo a ser aberto.", vbCritical, "Atenção"
End If
End Sub

Private Sub cmdAddAtiv_Click()

Set frm = frmAtiv
frm.sForm = "frmCadMob"
frmAtiv.show vbModeless

End Sub

Private Sub cmdAddCid_Click()

Set frm = frmCnsCidadao
frm.sForm = Me.Name
frm.sTipoCidadao = "P"
frmCnsCidadao.show

End Sub

Private Sub cmdAddCnae1_Click()

If mskCnae.ClipText = "" And cmdNovo.Visible = True Then Exit Sub
Set frm = frmCnaeNovo
frm.sForm = "frmCadMob"
frmCnaeNovo.show vbModeless
frmCnaeNovo.ZOrder 0

End Sub

Private Sub cmdAddCnae2_Click()
If cmbCnae.ListIndex = -1 And cmdNovo.Visible = True Then Exit Sub
Set frm = frmCnaeNovo
frm.sForm = "frmCadMob1"
frmCnaeNovo.show vbModeless
frmCnaeNovo.ZOrder 0

End Sub

Private Sub cmdAddISS_Click()
lIndex = m_cMenuISS.ShowPopupMenu(cmdAddISS.Left, cmdAddISS.Top, cmdAddISS.Left, cmdAddISS.Top, Me.ScaleWidth - cmdAddISS.Left - cmdAddISS.Width, cmdAddISS.Top + cmdAddISS.Height, False)
End Sub

Private Sub cmdAddResp_Click()
If cmdGravar.Visible = False Then Exit Sub
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frm.sTipoCidadao = "R"
frmCnsCidadao.show

End Sub

Private Sub cmdAddVS_Click()

frCNAE.Visible = True
frCNAE.ZOrder 0
cmdAddAtiv.Enabled = False
cmdAddISS.Enabled = False
cmdAddVS.Enabled = False
cmdDelISS.Enabled = False
cmdDelVS.Enabled = False
cmdGravar.Enabled = False
cmdCancel.Enabled = False
cmdOutrasAtiv.Enabled = False
cmdCancel.Enabled = False
cmdAddCnae1.Enabled = False
cmdAddCnae2.Enabled = False
cmdDelCnae.Enabled = False
cmdDelCnaeP.Enabled = False
cmdEdtCnae.Enabled = False
btAtivExtenso.Enabled = False
Tela(4).Enabled = False
lblDescCNAE.Caption = ""
LimpaMascara mskCodCNAE
cmbCriterio.Clear
lblValorCnae.Caption = "0,00"
txtQtdeCnae.Text = ""
mskCodCNAE.SetFocus
'Set frm = frmVigSanitaria
'frm.sForm = "frmCadMob"
'frmVigSanitaria.show vbModeless

End Sub

Private Sub cmdAlterar_Click()
    If Val(txtCodEmpresa.Text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    Evento = "Alterar"
    Eventos "INCLUIR"

End Sub

Private Sub cmdAlterar3_Click()

lblNovo2.Caption = "0"
With grdLivro
    If .Rows = 1 Then
        MsgBox "Não existem registros.", vbExclamation, "Atenção"
        Exit Sub
    Else
        lblSeq2.Caption = Format(.TextMatrix(.Row, 0), "0000")
        cmbModelo.Text = .TextMatrix(.Row, 1)
        If IsDate(.TextMatrix(.Row, 2)) Then
           mskAb.Text = .TextMatrix(.Row, 2)
        Else
           LimpaMascara mskAb
        End If
        If IsDate(.TextMatrix(.Row, 3)) Then
           mskEn.Text = .TextMatrix(.Row, 3)
        Else
           LimpaMascara mskEn
        End If
    End If
End With

cmbModelo.BackColor = Branco
mskAb.BackColor = Branco
mskEn.BackColor = Branco
cmbModelo.Enabled = True
mskAb.Enabled = True
mskEn.Enabled = True
cmdNovo3.Visible = False
cmdAlterar3.Visible = False
cmdExcluir3.Visible = False
cmdGravar3.Visible = True
cmdCancelar3.Visible = True

End Sub

Private Sub cmdAP_Click()

If Val(Left(txtCodEmpresa.Text, 7)) = 0 Then
    MsgBox "Selecione uma empresa.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataAP.Text) Then
    MsgBox "Digite a data do alvará provisório", vbExclamation, "Atenção"
    Exit Sub
End If

If Abs(DateDiff("d", Now, CDate(mskDataAP.Text))) > 31 Then
    MsgBox "Alvará provisório fora do prazo de 30 dias", vbExclamation, "Atenção"
    Exit Sub
End If

frmAlvara.show
frmAlvara.txtCodigo.Text = Val(Left(txtCodEmpresa.Text, 7))
frmAlvara.Carrega
'frmAlvara.txtAlvara.SetFocus
End Sub

Private Sub cmdCadCid_Click()
On Error GoTo Erro:
   Set frm2 = frmCidadao
   frm2.sForm = Me.Name
   frmCidadao.show
   Exit Sub
   
Erro:
   MsgBox "Clique na Árvore para selecionar Proprietário ou Proprietário Solidário.", vbExclamation, "Atenção"

End Sub

Private Sub cmdCancel_Click()
    If Evento = "Novo" Then Limpa
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdCancelar3_Click()
cmbModelo.Enabled = False
mskAb.Enabled = False
mskEn.Enabled = False
cmbModelo.BackColor = PnLivro.BackColor
mskAb.BackColor = PnLivro.BackColor
mskEn.BackColor = PnLivro.BackColor
cmdNovo3.Visible = True
cmdAlterar3.Visible = True
cmdExcluir3.Visible = True
cmdGravar3.Visible = False
cmdCancelar3.Visible = False

End Sub

Private Sub cmdCancelNF_Click()
frISS.Visible = False
End Sub

Private Sub cmdConsultar_Click()
sFormMob = "CM"
frmCnsMob.show vbModeless
frmCnsMob.ZOrder 0
End Sub

Private Sub cmdD1_Click()
nPointer = nPointer - 1
LoadPic
End Sub

Private Sub cmdD2_Click()
nPointer = nPointer + 1
LoadPic
End Sub

Private Sub cmdDelCid_Click()

If grdProp.Rows = 1 Then
   MsgBox "Selecione um proprietário.", vbExclamation, "Atenção"
Else
   If grdProp.Rows > 2 Then
      grdProp.RemoveItem (grdProp.Row)
   Else
      grdProp.Rows = 1
   End If
End If

End Sub

Private Sub cmdDelCnae_Click()

If cmdNovo.Visible = True Then Exit Sub
If cmbCnae.ListIndex = -1 Then Exit Sub
cmbCnae.RemoveItem (cmbCnae.ListIndex)

End Sub

Private Sub cmdDelCnaeP_Click()
If cmdNovo.Visible = True Then Exit Sub
If mskCnae.ClipText = "" Then Exit Sub
LimpaMascara mskCnae

End Sub

Private Sub cmdDelISS_Click()
If grdAtiv.Rows > 1 Then
   If grdAtiv.Rows = 2 Then
      lblTipoISS.Caption = ""
      grdAtiv.Rows = 1
   Else
      If grdAtiv.Row < 1 Then
         MsgBox "Selecione um item.", vbExclamation, "Atenção"
      Else
         grdAtiv.RemoveItem (grdAtiv.Row)
      End If
   End If
End If

End Sub

Private Sub cmdDelNF_Click()

If MsgBox("Excluir esta nota?", vbYesNo, "Atenção") = vbYes Then
    Sql = "DELETE FROM DECLARACAOISS WHERE CODIGO=" & Val(txtCodEmpresa.Text) & " AND ANO=" & Val(cmbAnoISS.Text) & " AND MES=" & Val(lblMesNF.Caption)
    cn.Execute Sql, rdExecDirect
    lblIss(lblMesNF.Caption).Caption = "0,00"
End If
frISS.Visible = False

End Sub

Private Sub cmdDelResp_Click()
If cmdGravar.Visible = False Then Exit Sub
txtNomeProf.Text = ""
End Sub

Private Sub cmdDelVS_Click()

If grdVS.Rows > 1 Then
   If grdVS.Rows = 2 Then
      grdVS.Rows = 1
   Else
      If grdVS.Row < 1 Then
         MsgBox "Selecione um item.", vbExclamation, "Atenção"
      Else
         grdVS.RemoveItem (grdVS.Row)
      End If
   End If
End If

End Sub

Private Sub cmdEditHist_Click()

Set frm = frmEditHist
frm.sForm = Me.Name
frmEditHist.show 1

End Sub

Private Sub cmdEdtCnae_Click()
If cmbCnae.ListIndex = -1 Then Exit Sub
Exit Sub
Set frm = frmCnae
frm.sForm = "frmCadMob2"
frmCnae.show vbModeless
frmCnae.ZOrder 0
End Sub

Private Sub cmdEmail_Click()
If Trim(txtEmailCont.Text) <> "" Then
    Call ShellExecute(0&, vbNullString, "mailto: " & txtEmailCont.Text, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox "Sem email.", vbCritical, "ERRO"
End If

End Sub

Private Sub cmdExcluir_Click()
Dim nCodReduz As Long

'MsgBox "FUNÇÃO BLOQUEADA", vbCritical, "AVISO DE SEGURANÇA"
'Exit Sub

nCodReduz = Val(txtCodEmpresa.Text)

If nCodReduz = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

If MsgBox("A Exclusão desta empresa apagará todos os dados cadastrias bem como todos os débitos dela." & vbCrLf & vbCrLf & "Você confirma ???", vbQuestion + vbYesNo + vbDefaultButton2, "CONFIRMAÇÃO DE EXCLUSÃO") = vbNo Then Exit Sub

If MsgBox("Voce tem certeza que quer excluir mesmo ????", vbQuestion + vbYesNo + vbDefaultButton2, "CONFIRMAÇÃO DE EXCLUSÃO") = vbNo Then Exit Sub

Log Form, Me.Caption, Exclusão, "Excluída empresa " & CStr(nCodReduz) & " - " & txtRazao.Text
Ocupado

Sql = "DELETE FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIONF WHERE CODIGOMOB=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOLIVRO WHERE CODIGOMOB=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM SIL WHERE CODIGO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOINF WHERE CODMOB=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOPROPRIETARIO WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOCNAE WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOENDENTREGA WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOATIVIDADEVS WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOATIVIDADETL WHERE CODIGOMOB=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & nCodReduz
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM MOBILIARIO WHERE CODIGOMOB=" & nCodReduz
cn.Execute Sql, rdExecDirect

Liberado
MsgBox "Empresa excluida.", vbExclamation, "Atenção"
Limpa

End Sub


Private Sub cmdExcluir3_Click()
If grdLivro.Rows = 1 Then
    MsgBox "Não existem registros.", vbExclamation, "Atenção"
    Exit Sub
Else
    If MsgBox("Excluir este Livro?", vbExclamation + vbYesNo, "Confirmação") = vbYes Then
        If grdLivro.Rows = 2 Then
            grdLivro.Rows = 1
        Else
            grdLivro.RemoveItem (grdLivro.Row)
        End If
    End If
End If

End Sub

Private Sub cmdExcluirDoc_Click()
Dim sArq As String
If Val(txtCodEmpresa.Text) = 0 Then Exit Sub
If MsgBox("Confirma a exclusão deste documento???", vbQuestion + vbYesNo + vbDefaultButton2, "CONFIRMAÇÃO DE EXCLUSÃO") = vbNo Then Exit Sub

sArq = aPic(nPointer).sArq


Sql = "delete from documentopic where codigo=" & Val(txtCodEmpresa.Text) & " and documento='" & sArq & "'"
cn.Execute Sql, rdExecDirect

ReDim aPic(0)
'DOCUMENTOPIC
Sql = "select documento from documentopic where codigo=" & Val(txtCodEmpresa.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sTmp = !Documento
        ReDim Preserve aPic(UBound(aPic) + 1)
        aPic(UBound(aPic)).sTipo = Left(sTmp, 2)
        aPic(UBound(aPic)).sAno = Mid(sTmp, 3, 4)
        aPic(UBound(aPic)).sArq = sTmp
        aPic(UBound(aPic)).sExt = Right(sTmp, 3)
        For x = 1 To UBound(aTipoDoc)
            If aTipoDoc(x).nCod = Val(Left(sTmp, 2)) Then
                aPic(UBound(aPic)).sTipoExt = aTipoDoc(x).sNome
                Exit For
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

If UBound(aPic) > 0 Then
    lblPagDoc.Caption = "Documento 1 de " & UBound(aPic)
    cmdD2.Enabled = True
    nPointer = 1
    LoadPic
End If

End Sub

Private Sub cmdExtrato_Click()
If Val(txtCodEmpresa.Text) = 0 Then Exit Sub
CodEmpresa = Val(txtCodEmpresa.Text)
frmDebitoImob.show: frmDebitoImob.ZOrder 0
End Sub

Private Sub cmdFoto_Click()
If Val(txtCodEmpresa.Text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If
If Val(lblDist.Caption) = 0 Then
   MsgBox "Não foi possível relacionar esta empresa com nenhum imóvel da cidade." & vbCrLf & "Não existe nenhum imóvel cadastrado no logradouro " & txtNomeLogr.Text & " nº " & Val(txtNumero.Text) & " conforme especificados nesta empresa.", vbExclamation, "Atenção"
Else
    sFormFoto = "M"
'    frmImageImovel.show
'    frmImageImovel.ZOrder 0
End If

End Sub

Private Sub cmdGerar_Click()

If Val(txtCodEmpresa.Text) > 0 Then frmPeriodoSN.show vbModal
SNCheck
End Sub

Private Sub cmdGravar_Click()
If bLocal Then
    Exit Sub
End If
    
    If Not (Valida) Then Exit Sub
    Grava
    Eventos "INICIAR"
    Evento = ""

End Sub

Private Sub Grava()
On Error GoTo Erro
Dim qd As New rdoQuery, sCnae As String, sSecao As String, nDivisao As Integer, nGrupo As Integer, sClasse As String, nClasse As Integer, nSubClasse As Integer
Dim MinCod As Long, MaxCod As Long, bLogChange As Boolean, sLog As String, nTipoEE As Integer
Dim nSeq As Long, nTipoIss As Integer, nHorario As Integer, nBairro As Integer

If chkIE.value = vbChecked And Not bDispensadoIE Then
    grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & "Dispensado do Iss eletrônico por " & NomeDeLogin & Chr(9) & "GTI"
    bDispensadoIE = True
End If
If chkIE.value = vbUnchecked And bDispensadoIE Then
    grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & "Removido a dispensa do Iss eletrônico por " & NomeDeLogin & Chr(9) & "GTI"
    bDispensadoIE = False
End If

If chkIsentoISS.value = vbChecked And Not bIsentoIss Then
    grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & "Inserido X em Isento ISS por " & NomeDeLogin & Chr(9) & "GTI"
    bIsentoIss = True
End If
If chkIsentoISS.value = vbUnchecked And bIsentoIss Then
    grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & "Removido X do Isento ISS por " & NomeDeLogin & Chr(9) & "GTI"
    bIsentoIss = False
End If

If chkIsentoTaxa.value = vbChecked And Not bIsentoTaxa Then
    grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & "Inserido X em Isento Taxa por " & NomeDeLogin & Chr(9) & "GTI"
    bIsentoTaxa = True
End If
If chkIsentoTaxa.value = vbUnchecked And bIsentoTaxa Then
    grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & "Removido X do Isento Taxa por " & NomeDeLogin & Chr(9) & "GTI"
    bIsentoTaxa = False
End If

If Evento = "Novo" Then
    Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB>100000 and CODIGOMOB<200000 ORDER BY CODIGOMOB"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
           If MinCod = 0 Then
              MinCod = !codigomob
           Else
              MaxCod = !codigomob
              If MaxCod - MinCod > 1 Then
                  MaxCod = MinCod + 1
                  Exit Do
              Else
                  MinCod = MaxCod
              End If
           End If
          .MoveNext
        Loop
       .Close
    End With
Else
    If InStr(1, txtCodEmpresa.Text, "-", vbBinaryCompare) > 0 Then
       MaxCod = Left$(txtCodEmpresa.Text, Len(txtCodEmpresa.Text) - 2)
    Else
       MaxCod = txtCodEmpresa.Text
    End If
    
    'verifica log
    ReDim aLog(0)
    bLogChange = False
    If aLog(0).sDataAb <> mskDataAb.Text And aLog(0).sDataAb <> "" Then
        bLogChange = True
        sLog = "Alterada Data de Abertura de " & aLog(0).sDataAb & " para " & mskDataAb.Text & " pelo usuário: " & NomeDeLogin
    End If
    If aLog(0).sDataEn <> mskDataEn.Text And aLog(0).sDataEn <> "" Then
        bLogChange = True
        sLog = "Alterada Data de Encerramento de " & aLog(0).sDataEn & " para " & mskDataEn.Text & " pelo usuário: " & NomeDeLogin
    End If
    If aLog(0).sDataProcAb <> mskDataPAb.Text And aLog(0).sDataProcAb <> "" Then
        bLogChange = True
        sLog = "Alterada Data de Abertura do Processo de " & aLog(0).sDataProcAb & " para " & mskDataPAb.Text & " pelo usuário: " & NomeDeLogin
    End If
    If aLog(0).sDataProcEn <> mskDataPEn.Text And aLog(0).sDataProcEn <> "" Then
        bLogChange = True
        sLog = "Alterada Data de Encerramento do Processo de " & aLog(0).sDataProcEn & " para " & mskDataPEn.Text & " pelo usuário: " & NomeDeLogin
    End If
    
    If bLogChange Then
        grdHist.AddItem Format(Now, "dd/mm/yyyy") & Chr(9) & grdHist.Rows & Chr(9) & sLog & Chr(9) & "GTI"
    End If
End If

'Grava na Tabela Mobiliário
GRAVAMOBILIARIO:
If Val(txtCapital.Text) = 0 Then txtCapital.Text = 0
If cmbHorario.ListIndex = -1 Then nHorario = 2 Else nHorario = cmbHorario.ItemData(cmbHorario.ListIndex)
If cmbBairro.ListIndex = -1 Then nBairro = 999 Else nBairro = cmbBairro.ItemData(cmbBairro.ListIndex)

If Evento = "Novo" Then
    Sql = "INSERT MOBILIARIO(CODIGOMOB,DVMOB,RAZAOSOCIAL,NOMEFANTASIA,CPF,CNPJ,INSCESTADUAL,DATAABERTURA,NUMPROCESSO,"
    Sql = Sql & "DATAPROCESSO,DATAENCERRAMENTO,NUMPROCENCERRAMENTO,DATAPROCENCERRAMENTO,HORARIO,CODLOGRADOURO,COMPLEMENTO,"
    Sql = Sql & "NUMERO,SIGLAUF,CODCIDADE,CODBAIRRO,CEP,HOMEPAGE,NOMECONTATO,FONECONTATO,FAXCONTATO,CARGOCONTATO,"
    Sql = Sql & "EMAILCONTATO,RESPCONTABIL,CAPITALSOCIAL,QTDEEMPREGADO,CODPROFRESP,NOMEORGAO,NUMREGISTRORESP,CODATIVIDADE,"
    Sql = Sql & "ATIVEXTENSO,CODIGOALIQ,AREATL,QTDEPROF,RG,ORGAO,NOMELOGRADOURO,VISTORIA,REGESPECIAL,ALVARA,ISENTOTAXA,MEI,"
    Sql = Sql & "HORARIOEXT,ISSELETRO,DISPENSAIEDATA,DISPENSAIEPROC,DTALVARAPROVISORIO,SENHAISS,INSCTEMP,HORAS24,ISENTOISS,"
    Sql = Sql & "BOMBONIERI,EMITENF,DANFE,IMOVEL,SIL,substituto_tributario_issqn,INDIVIDUAL,ponto_agencia,DDD_NF,TELEFONE_NF,EMAIL_NF) VALUES(" & MaxCod & "," & RetornaDVCodReduzido(MaxCod) & ",'"
    Sql = Sql & Mask(txtRazao.Text) & "','" & Mask(txtFantasia.Text) & "'," & IIf(mskCPF.ClipText <> "", "'" & mskCPF.ClipText & "'", "Null") & ","
    Sql = Sql & IIf(mskCNPJ.ClipText <> "", "'" & mskCNPJ.ClipText & "'", "Null") & "," & IIf(IsNumeric(txtInscEst.Text), "'" & txtInscEst.Text & "'", "Null") & "," & IIf(IsDate(mskDataAb.Text), "'" & Format(mskDataAb.Text, "mm/dd/yyyy") & "'", "Null") & ",'"
    Sql = Sql & txtNumProcA.Text & "'," & IIf(IsDate(mskDataPAb.Text), "'" & Format(mskDataPAb.Text, "mm/dd/yyyy") & "'", "Null") & "," & IIf(IsDate(mskDataEn.Text), "'" & Format(mskDataEn.Text, "mm/dd/yyyy") & "'", "Null") & ",'" & txtNumProcE.Text & "'," & IIf(IsDate(mskDataPEn.Text), "'" & Format(mskDataPEn.Text, "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & nHorario & "," & Val(txtCodLogr.Text) & ",'" & Mask(txtCompl.Text) & "'," & Val(txtNumero.Text) & ",'" & Left$(Right$(cmbUF.Text, 3), 2) & "',"
    Sql = Sql & cmbCidade.ItemData(cmbCidade.ListIndex) & "," & nBairro & ",'" & mskCEP.ClipText & "','" & Mask(txtHP.Text) & "','" & Mask(txtNomeContato.Text) & "','"
    Sql = Sql & Mask(txtFone.Text) & "','" & Mask(txtFax.Text) & "','" & Mask(txtCargo.Text) & "','" & LCase(Mask(Left(txtEmail.Text, 100))) & "'," & IIf(Val(txtCodEsc.Text) > 0, Val(txtCodEsc.Text), "Null") & "," & Virg2Ponto(RemovePonto(txtCapital.Text)) & ","
    Sql = Sql & Val(txtNumFunc.Text) & "," & IIf(Trim$(txtNomeProf.Text) <> "", Val(Left$(txtNomeProf.Text, 5)), 0) & ",'" & Mask(txtTipoConselho.Text) & "','" & Mask(txtNumRegistro.Text) & "',"
    Sql = Sql & Val(Left$(txtAtiv.Text, 5)) & ",'" & Mask(txtAtivExt.Text) & "'," & Val(lblAliq.Caption) & "," & Virg2Ponto(RemovePonto(txtArea.Text)) & "," & Val(txtQtde.Text) & ",'" & mskRG.Text & "','" & txtOrgao.Text & "','"
    Sql = Sql & IIf(Val(txtCodLogr.Text) > 0, Null, txtNomeLogr.Text) & "'," & IIf((chkVistoria.value = 0), 0, 1) & "," & chkRE.value & "," & chkAlvara.value & "," & chkIsentoTaxa.value & "," & 0 & ",'"
    Sql = Sql & Mask(txtHorarioExt.Text) & "'," & IIf((chkIE.value = 0), 0, 1) & "," & IIf(IsDate(mskDataIE.Text), "'" & Format(mskDataIE.Text, "mm/dd/yyyy") & "'", "Null") & "," & IIf(IsDate(mskDataIE.Text), "'" & txtNumProcIE.Text & "'", "Null") & "," & IIf(IsDate(mskDataAP.Text), "'" & Format(mskDataAP.Text, "mm/dd/yyyy") & "'", "Null") & ",'" & Mask(txtSenha.Text) & "',"
    Sql = Sql & IIf((chkInscTemp.value = 0), 0, 1) & "," & IIf((chk24horas.value = 0), 0, 1) & "," & chkIsentoISS.value & "," & IIf((chkBombon.value = 0), 0, 1) & "," & IIf((chkEmiteNF.value = 0), 0, 1) & "," & IIf((chkDanfe.value = 0), 0, 1) & "," & Val(txtImovel.Text) & ",'" & Mask(txtSIL.Text) & "'," & IIf((chkSubstitutoTributario.value = 0), 0, 1) & ","
    Sql = Sql & IIf((chkEmpInd.value = 0), 0, 1) & ",'" & Mask(txtPonto.Text) & "','" & txtDDDNF.Text & "','" & txtFoneNF.Text & "','" & LCase(Mask(Trim(txtEmailNF.Text))) & "')"
Else
    Sql = "UPDATE MOBILIARIO SET RAZAOSOCIAL='" & Mask(txtRazao.Text) & "',NOMEFANTASIA='" & Mask(txtFantasia.Text) & "',CPF=" & IIf(mskCPF.ClipText <> "", "'" & mskCPF.ClipText & "'", "Null") & ",CNPJ=" & IIf(mskCNPJ.ClipText <> "", "'" & mskCNPJ.ClipText & "'", "Null") & ","
    Sql = Sql & "INSCESTADUAL=" & IIf(IsNumeric(txtInscEst.Text), "'" & txtInscEst.Text & "'", "Null") & ",DATAABERTURA=" & IIf(IsDate(mskDataAb.Text), "'" & Format(mskDataAb.Text, "mm/dd/yyyy") & "'", "Null") & ",NUMPROCESSO='" & txtNumProcA.Text & "',"
    Sql = Sql & "DATAPROCESSO=" & IIf(IsDate(mskDataPAb.Text), "'" & Format(mskDataPAb.Text, "mm/dd/yyyy") & "'", "Null") & ",DATAENCERRAMENTO=" & IIf(IsDate(mskDataEn.Text), "'" & Format(mskDataEn.Text, "mm/dd/yyyy") & "'", "Null") & ",NUMPROCENCERRAMENTO='" & txtNumProcE.Text & "',"
    Sql = Sql & "DATAPROCENCERRAMENTO=" & IIf(IsDate(mskDataPEn.Text), "'" & Format(mskDataPEn.Text, "mm/dd/yyyy") & "'", "Null") & ",HORARIO=" & nHorario & ",CODLOGRADOURO=" & Val(txtCodLogr.Text) & ",COMPLEMENTO='" & Mask(txtCompl.Text) & "',"
    Sql = Sql & "NUMERO=" & Val(txtNumero.Text) & ",SIGLAUF='" & Left$(Right$(cmbUF.Text, 3), 2) & "',CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex) & ",CODBAIRRO=" & nBairro & ",CEP='" & mskCEP.ClipText & "',HOMEPAGE='" & txtHP.Text & "',"
    Sql = Sql & "NOMECONTATO='" & txtNomeContato.Text & "',FONECONTATO='" & Mask(txtFone.Text) & "',FAXCONTATO='" & txtFax.Text & "',CARGOCONTATO='" & txtCargo.Text & "',EMAILCONTATO='" & LCase(Mask(txtEmail.Text)) & "',"
    Sql = Sql & "RESPCONTABIL=" & IIf(Val(txtCodEsc.Text) > 0, Val(txtCodEsc.Text), "Null") & ",CAPITALSOCIAL=" & Virg2Ponto(RemovePonto(txtCapital.Text)) & ",QTDEEMPREGADO=" & Val(txtNumFunc.Text) & ","
    Sql = Sql & "CODPROFRESP=" & IIf(Trim$(txtNomeProf.Text) <> "", Val(Left$(txtNomeProf.Text, 6)), 0) & ",NOMEORGAO='" & txtTipoConselho.Text & "',NUMREGISTRORESP='" & txtNumRegistro.Text & "',CODATIVIDADE=" & Val(Left$(txtAtiv.Text, 5)) & ","
    Sql = Sql & "ATIVEXTENSO='" & Mask(txtAtivExt.Text) & "',CODIGOALIQ=" & Val(lblAliq.Caption) & ",AREATL=" & Virg2Ponto(RemovePonto(txtArea.Text)) & ",QTDEPROF=" & Val(txtQtde.Text) & ",RG='" & mskRG.Text & "',ORGAO='" & txtOrgao.Text & "',NOMELOGRADOURO=" & IIf(Val(txtCodLogr.Text) > 0, "Null", "'" & txtNomeLogr.Text & "'") & ","
    Sql = Sql & "VISTORIA=" & IIf((chkVistoria.value = 0), 0, 1) & ",MEI=" & 0 & ",ISSELETRO=" & IIf((chkIE.value = 0), 0, 1) & ",ISENTOTAXA=" & chkIsentoTaxa.value & ",REGESPECIAL=" & chkRE.value & ",ALVARA=" & chkAlvara.value & ",HORARIOEXT='" & Mask(txtHorarioExt.Text) & "',"
    Sql = Sql & "DISPENSAIEDATA=" & IIf(IsDate(mskDataIE.Text), "'" & Format(mskDataIE.Text, "mm/dd/yyyy") & "'", "Null") & ",DISPENSAIEPROC=" & IIf(IsDate(mskDataIE.Text), "'" & txtNumProcIE.Text & "'", "Null") & ",DTALVARAPROVISORIO=" & IIf(IsDate(mskDataAP.Text), "'" & Format(mskDataAP.Text, "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & "substituto_tributario_issqn=" & IIf((chkSubstitutoTributario.value = 0), 0, 1) & ",INDIVIDUAL=" & IIf((chkEmpInd.value = 0), 0, 1) & ",LIBERADO_VRE=" & IIf((chkLiberadoVRE.value = 0), 0, 1) & ",DDD_NF='" & txtDDDNF.Text & "',TELEFONE_NF='" & txtFoneNF.Text & "',EMAIL_NF='" & Mask(Trim(LCase(txtEmailNF.Text))) & "'"
    If txtSenha.Text <> "** Restrito **" Then
        Sql = Sql & ",SENHAISS='" & Mask(txtSenha.Text) & "'"
    End If
    Sql = Sql & ",INSCTEMP=" & IIf((chkInscTemp.value = 0), 0, 1) & ",HORAS24=" & IIf((chk24horas.value = 0), 0, 1) & ",ISENTOISS=" & chkIsentoISS.value & ",BOMBONIERI=" & IIf((chkBombon.value = 0), 0, 1) & ",EMITENF=" & IIf((chkEmiteNF.value = 0), 0, 1) & ",DANFE=" & chkDanfe.value & ",IMOVEL=" & Val(txtImovel.Text) & ",SIL='" & Mask(txtSIL.Text) & "', PONTO_AGENCIA='" & Mask(txtPonto.Text) & "'"
    Sql = Sql & " WHERE  CODIGOMOB = " & MaxCod
End If
cn.Execute Sql, rdExecDirect

'Grava cnae
'cnae principal
If mskCnae.ClipText = "" Then GoTo prop
Sql = "DELETE FROM MOBILIARIOCNAE WHERE CODMOBILIARIO=" & MaxCod
cn.Execute Sql, rdExecDirect

sCnae = mskCnae.Text
sSecao = ""
nDivisao = Val(Left(sCnae, 2))
nGrupo = Val(Mid(sCnae, 3, 1))
sClasse = Mid(sCnae, 4, 3)
sClasse = Left(sClasse, 1) & Right(sClasse, 1)
nClasse = Val(sClasse)
nSubClasse = Val(Right(sCnae, 2))

Sql = "INSERT MOBILIARIOCNAE(CODMOBILIARIO,SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,PRINCIPAL,CNAE) VALUES("
Sql = Sql & MaxCod & ",'" & sSecao & "'," & nDivisao & "," & nGrupo & "," & nClasse & "," & nSubClasse & "," & 1 & ",'" & sCnae & "')"
cn.Execute Sql, rdExecDirect

'cnae outros
With cmbCnae
    If .ListCount > 0 Then
        For x = 0 To .ListCount - 1
            sCnae = cmbCnae.List(x)
            sSecao = ""
            nDivisao = Val(Left(sCnae, 2))
            nGrupo = Val(Mid(sCnae, 3, 1))
            sClasse = Mid(sCnae, 4, 3)
            sClasse = Left(sClasse, 1) & Right(sClasse, 1)
            nClasse = Val(sClasse)
            nSubClasse = Val(Right(sCnae, 2))
            
        
            
            Sql = "INSERT MOBILIARIOCNAE(CODMOBILIARIO,SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,PRINCIPAL,CNAE) VALUES("
            Sql = Sql & MaxCod & ",'" & sSecao & "'," & nDivisao & "," & nGrupo & "," & nClasse & "," & nSubClasse & "," & 0 & ",'" & sCnae & "')"
            cn.Execute Sql, rdExecDirect
        Next
    End If
End With

prop:
'Grava Mobiliario Proprietario
Sql = "DELETE FROM MOBILIARIOPROPRIETARIO WHERE CODMOBILIARIO=" & MaxCod
cn.Execute Sql, rdExecDirect
With grdProp
    If .Rows > 1 Then
        For x = 1 To .Rows - 1
            Sql = "INSERT MOBILIARIOPROPRIETARIO(CODMOBILIARIO,CODCIDADAO) VALUES("
            Sql = Sql & MaxCod & "," & .TextMatrix(x, 0) & ")"
            cn.Execute Sql, rdExecDirect
        Next
    End If
End With


If OptEE(0).value = True Then
    nTipoEE = 1
Else
    nTipoEE = 2
End If

'Grava Mobiliario Endereço de Entrega
Sql = "DELETE FROM MOBILIARIOENDENTREGA WHERE CODMOBILIARIO=" & MaxCod
cn.Execute Sql, rdExecDirect

If chkEnd.value = vbUnchecked Then
   Sql = "INSERT MOBILIARIOENDENTREGA(CODMOBILIARIO,TIPO,CODLOGRADOURO,NOMELOGRADOURO,"
   Sql = Sql & "NUMIMOVEL,COMPLEMENTO,UF,CODCIDADE,CODBAIRRO,CEP) VALUES("
   Sql = Sql & MaxCod & "," & nTipoEE & "," & Val(txtEECodLogr.Text) & ",'" & txtEENomeLogr.Text & "',"
   Sql = Sql & Val(txtEENumero.Text) & ",'" & txtEECompl.Text & "','" & Left$(Right$(cmbEEUf.Text, 3), 2) & "',"
   If cmbEEBairro.ListIndex = -1 Then
      Sql = Sql & cmbEECidade.ItemData(cmbEECidade.ListIndex) & "," & 999 & ",'"
   Else
      Sql = Sql & cmbEECidade.ItemData(cmbEECidade.ListIndex) & "," & cmbEEBairro.ItemData(cmbEEBairro.ListIndex) & ",'"
   End If
   Sql = Sql & mskEECep.Text & "')"
   cn.Execute Sql, rdExecDirect
End If

'Grava Mobiliario ISS
Sql = "DELETE FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & MaxCod
cn.Execute Sql, rdExecDirect
If grdAtiv.Rows > 1 Then
    For x = 1 To grdAtiv.Rows - 1
        If grdAtiv.TextMatrix(x, 0) = "F" Then
            nTipoIss = 11
        ElseIf grdAtiv.TextMatrix(x, 0) = "E" Then
            nTipoIss = 12
        ElseIf grdAtiv.TextMatrix(x, 0) = "V" Then
            nTipoIss = 13
        End If
    
       Sql = "INSERT MOBILIARIOATIVIDADEISS(CODMOBILIARIO,CODTRIBUTO,CODATIVIDADE,SEQ,QTDEISS,VALORISS) VALUES("
       Sql = Sql & MaxCod & "," & nTipoIss & "," & Val(Left$(grdAtiv.TextMatrix(x, 1), 4)) & "," & x & ","
       Sql = Sql & Virg2Ponto(grdAtiv.TextMatrix(x, 2)) & "," & Virg2Ponto(RemovePonto(grdAtiv.TextMatrix(x, 3))) & ")"
       cn.Execute Sql, rdExecDirect
    Next
End If

Sql = "DELETE FROM MOBILIARIOVS WHERE CODIGO=" & MaxCod
cn.Execute Sql, rdExecDirect
If grdVS.Rows > 1 Then
    With grdVS
        For x = 1 To .Rows - 1
            sCnae = .TextMatrix(x, 0)
            Sql = "INSERT MOBILIARIOVS(CODIGO,CNAE,CRITERIO,QTDE) VALUES("
            Sql = Sql & MaxCod & ",'" & RetornaNumero(sCnae) & "'," & Val(.TextMatrix(x, 1)) & "," & Val(.TextMatrix(x, 3)) & ")"
            cn.Execute Sql, rdExecDirect
        Next
    End With
End If



'Grava Mobiliario TL
Sql = "DELETE FROM MOBILIARIOATIVIDADETL WHERE CODIGOMOB=" & MaxCod
cn.Execute Sql, rdExecDirect
If grdTemp.Rows > 1 Then
    For x = 1 To grdTemp.Rows - 1
       Sql = "INSERT MOBILIARIOATIVIDADETL(CODIGOMOB,CODATIVIDADE,CODIGOALIQ,AREA,QTDE) VALUES("
       Sql = Sql & MaxCod & "," & Val(grdTemp.TextMatrix(x, 0)) & "," & Val(grdTemp.TextMatrix(x, 3)) & "," & Virg2Ponto(grdTemp.TextMatrix(x, 4)) & "," & Val(grdTemp.TextMatrix(x, 5)) & ")"
       cn.Execute Sql, rdExecDirect
    Next
End If

'Grava Mobiliario NF

Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQNOT'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nSeq = Format(!valparam + 1, "000000")
   .Close
End With

If nSeq = 187 And Year(Now) = 2007 Then
    nSeq = 188
End If


Sql = "DELETE FROM MOBILIARIONF WHERE CODIGOMOB=" & MaxCod
cn.Execute Sql, rdExecDirect
If grdNF.Rows > 1 Then
    For x = 1 To grdNF.Rows - 1
       If grdNF.TextMatrix(x, 4) = "Definição Automática" Then
          grdNF.TextMatrix(x, 4) = nSeq
          Sql = "UPDATE PARAMETROS SET VALPARAM=" & nSeq & " WHERE NOMEPARAM='SEQNOT'"
          cn.Execute Sql, rdExecDirect
          nSeq = nSeq + 1
       End If
       grdNF.TextMatrix(x, 4) = Trim$(grdNF.TextMatrix(x, 4))
       Sql = "INSERT MOBILIARIONF(CODIGOMOB,SEQ,SERIE,NUMINI,NUMFIM,NUMAUT,DATAAUT,CANCEL,USUARIO) VALUES("
       Sql = Sql & MaxCod & "," & Val(grdNF.TextMatrix(x, 0)) & ",'" & grdNF.TextMatrix(x, 1) & "'," & Val(grdNF.TextMatrix(x, 2)) & "," & Val(grdNF.TextMatrix(x, 3)) & ",'"
       If IsNumeric(grdNF.TextMatrix(x, 4)) Then
           Sql = Sql & grdNF.TextMatrix(x, 4) & "'," & IIf(grdNF.TextMatrix(x, 5) = "", "NULL,0", "'" & Format(grdNF.TextMatrix(x, 5), "mm/dd/yyyy") & "',0") & ",'" & grdNF.TextMatrix(x, 6) & "')"
       Else
          If InStr(1, grdNF.TextMatrix(x, 4), "Cancelado", vbBinaryCompare) = 0 And InStr(1, grdNF.TextMatrix(x, 4), "XXX", vbBinaryCompare) = 0 Then
             Sql = Sql & CStr(nSeq) & "'," & IIf(grdNF.TextMatrix(x, 5) = "", "NULL,0", "'" & Format(grdNF.TextMatrix(x, 5), "mm/dd/yyyy") & "',0") & ",'" & grdNF.TextMatrix(x, 6) & "')"
          Else
             If InStr(1, grdNF.TextMatrix(x, 4), "Cancelado", vbBinaryCompare) > 0 Then
                Sql = Sql & Left$(grdNF.TextMatrix(x, 4), InStr(1, grdNF.TextMatrix(x, 4), " (C")) & "'," & IIf(grdNF.TextMatrix(x, 5) = "", "NULL,1", "'" & Format(grdNF.TextMatrix(x, 5), "mm/dd/yyyy") & "',1") & ",'" & grdNF.TextMatrix(x, 6) & "')"
             Else
                Sql = Sql & grdNF.TextMatrix(x, 4) & "'," & IIf(grdNF.TextMatrix(x, 5) = "", "NULL,0", "'" & Format(grdNF.TextMatrix(x, 5), "mm/dd/yyyy") & "',0") & ",'" & grdNF.TextMatrix(x, 6) & "')"
             End If
          End If
       End If
       cn.Execute Sql, rdExecDirect
    Next
End If

'Grava Mobiliario livro
Sql = "DELETE FROM MOBILIARIOLIVRO WHERE CODIGOMOB=" & MaxCod
cn.Execute Sql, rdExecDirect
If grdLivro.Rows > 1 Then
    For x = 1 To grdLivro.Rows - 1
       Sql = "INSERT MOBILIARIOLIVRO(CODIGOMOB,SEQ,MODELO,DATAAB,DATAEN) VALUES("
       Sql = Sql & MaxCod & "," & Val(grdLivro.TextMatrix(x, 0)) & ",'" & grdLivro.TextMatrix(x, 1) & "','" & Format(grdLivro.TextMatrix(x, 2), "mm/dd/yyyy") & "'," & IIf(grdLivro.TextMatrix(x, 3) = "", "NULL", "'" & Format(grdLivro.TextMatrix(x, 3), "mm/dd/yyyy") & "'") & ")"
       cn.Execute Sql, rdExecDirect
    Next
End If

'grava placa
Sql = "DELETE FROM MOBILIARIOPLACA WHERE CODIGO=" & MaxCod
cn.Execute Sql, rdExecDirect

For x = 0 To lstPlaca.ListCount - 1
   Sql = "INSERT MOBILIARIOPLACA(CODIGO,PLACA) VALUES(" & MaxCod & ",'" & lstPlaca.List(x) & "')"
   cn.Execute Sql, rdExecDirect
Next



'Grava historico
Sql = "DELETE FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & MaxCod
cn.Execute Sql, rdExecDirect
With grdHist
    If .Rows > 1 Then
        For x = 1 To .Rows - 1
'           Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USUARIO) VALUES("
'           Sql = Sql & MaxCod & "," & x & ",'" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "','" & Mask(.TextMatrix(x, 2)) & "','" & Mask(.TextMatrix(x, 3)) & "')"
           Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
           Sql = Sql & MaxCod & "," & x & ",'" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "','" & Mask(.TextMatrix(x, 2)) & "'," & RetornaUsuarioID(Mask(.TextMatrix(x, 3))) & ")"
           cn.Execute Sql, rdExecDirect
        Next
    End If
End With


fim:
'Atualiza Dados
If Evento = "Novo" Then
   txtCodEmpresa.Text = MaxCod
   'Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtDesc.Text
 ElseIf Evento = "Alterar" Then
   'Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.text, "000") & " de " & sOldDesc & " para " & txtDesc.text
End If


'Integração_Eicon
Sql = "select codigo from eicon_empresa where codigo=" & MaxCod
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    Sql = "insert eicon_empresa(codigo) values(" & MaxCod & ")"
    cn.Execute Sql, rdExecDirect
End If
RdoAux.Close

Exit Sub
Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Function Valida() As Boolean
Dim Sql As String, RdoAux As rdoResultset

Valida = False

If Val(txtQtde.Text) = 0 Then txtQtde.Text = "1"

If Trim$(txtRazao.Text) = "" Then
    MsgBox "Digite a Razão Social.", vbExclamation, "Erro de Validação"
    Exit Function
End If
If mskCPF.ClipText <> "" Then
    If Not ValidaCPF(mskCPF.ClipText) Then
        MsgBox "CPF Invalido.", vbExclamation, "Atenção"
        Exit Function
    End If
End If
If mskCNPJ.ClipText <> "" Then
    If Not ValidaCGC(mskCNPJ.ClipText) Then
       MsgBox "CNPJ Invalido.", vbExclamation, "Atenção"
       Exit Function
    End If
End If
If cmbUF.ListIndex = -1 Then
    MsgBox "Selecione a UF da localização da empresa .", vbExclamation, "Atenção"
    Exit Function
End If
If cmbCidade.ListIndex = -1 Then
    MsgBox "Selecione a Cidade da localização da empresa .", vbExclamation, "Atenção"
    Exit Function
End If
If Len(mskCEP.ClipText) > 0 And Len(mskCEP.ClipText) < 8 Then
   MsgBox "Cép inválido.", vbExclamation, "Atenção"
   Exit Function
End If
If Val(txtEECodLogr.Text) = 0 And cmbEECidade.Text = "JABOTICABAL" And chkEnd.value = 0 Then
   MsgBox "Selecione o Logradouro de Jaboticabal.", vbCritical, "Erro de Validação."
   Exit Function
End If
If cmbCidade.ItemData(cmbCidade.ListIndex) <> 413 And Val(txtEECodLogr.Text) = 0 And txtNomeLogr.Text = "" Then
   MsgBox "Digite o Logradouro.", vbCritical, "Erro de Validação."
   Exit Function
End If

If chkEnd.value = vbUnchecked Then
    If cmbEEUf.ListIndex = -1 Then
        MsgBox "Selecione a UF do endereço de entrega .", vbExclamation, "Atenção"
        Exit Function
    End If
    If cmbEECidade.ListIndex = -1 Then
        MsgBox "Selecione a Cidade do endereço de entrega .", vbExclamation, "Atenção"
        Exit Function
    End If
    If Val(txtEECodLogr.Text) = 0 And Trim$(txtEENomeLogr.Text) = "" Then
       MsgBox "Digite o nome do logradouro de entrega.", vbExclamation, "Atenção"
       Exit Function
    End If
    If Len(mskEECep.ClipText) > 0 And Len(mskEECep.ClipText) < 8 Then
       MsgBox "Cép inválido.", vbExclamation, "Atenção"
       Exit Function
    End If
End If

If Trim(txtEmail.Text) <> "" Then
    If Trim(txtEmailNF.Text) = "" Then
        MsgBox "Digite o Email para Nota Fiscal.", vbExclamation, "Erro de Validação"
        Exit Function
    End If
End If

If Trim(txtFone.Text) <> "" Then
    If Trim(txtDDDNF.Text) = "" Or Trim(txtFoneNF.Text) = "" Then
        MsgBox "Digite o DDD e Telefone para Nota Fiscal.", vbExclamation, "Erro de Validação"
        Exit Function
    End If
End If

If cmbCnae.ListCount > 0 And mskCnae.ClipText = "" Then
    MsgBox "CNAE Secundária sem CNAE Principal cadastrada.", vbCritical, "ERRO"
    Exit Function
End If


If Evento = "Novo" Then
    If mskCNPJ.ClipText <> "" Then
        Sql = "select codigomob,razaosocial from mobiliario where cnpj='" & mskCNPJ.ClipText & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            If MsgBox("CNPJ já cadastrado para a empresa: " & vbCrLf & RdoAux!codigomob & "-" & RdoAux!razaosocial & vbCrLf & "Deseja gravar assim mesmo?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
                Exit Function
            End If
        End If
    End If
End If

If chkIsentoISS.value = vbChecked And chkIsentoTaxa.value = vbChecked Then
Else
If chkAlvara.value = vbChecked Then
End If
End If
If Val(txtArea.Text) = 0 Then txtArea.Text = 0

If chkIE.value = vbChecked Then
    If Not IsDate(mskDataIE.Text) Or txtNumProcIE.Text = "" Then
        MsgBox "Digite data e número do processo de dispensa de ISS Eletrônico.", vbExclamation, "Atenção"
        Exit Function
    End If
End If

If chkIE.value = vbChecked And UCase(lblISSEletro.Caption) = "SIM" Then
    MsgBox "Para dispensar o contribuinte do ISS Eletrônico, insira a data final no ISS Eletrônico.", vbExclamation, "Atenção"
    Exit Function
End If

Valida = True

End Function


Private Sub cmdGravar3_Click()

Dim sData1 As String, sData2 As String

If IsDate(mskAb.Text) Then
    sData1 = mskAb.Text
Else
    If mskAb.ClipText = "" Then
        MsgBox "Data de Abertura inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

If IsDate(mskEn.Text) Then
    sData2 = mskEn.Text
Else
    If mskEn.ClipText <> "" Then
        MsgBox "Data de Encerramento inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
    sData2 = ""
End If

If cmbModelo.ListIndex = -1 Then
    MsgBox "Selecione o Modelo.", vbExclamation, "Atenção"
    Exit Sub
Else
    If lblNovo2.Caption = "1" Then
        grdLivro.AddItem lblSeq2.Caption & Chr(9) & cmbModelo.Text & Chr(9) & sData1 & Chr(9) & sData2
'        If grdLivro.Rows > 1 Then
'           grdLivro.TextMatrix(grdLivro.Rows - 1, 3) = sData1
'        End If
    Else
        With grdLivro
            .TextMatrix(.Row, 1) = cmbModelo.Text
            .TextMatrix(.Row, 2) = sData1
            .TextMatrix(.Row, 3) = sData2
        End With
    End If
End If

Sql = "UPDATE PARAMETROS SET VALPARAM= VALPARAM + 1  WHERE NOMEPARAM='SEQLIV'"
cn.Execute Sql, rdExecDirect


cmbModelo.Enabled = False
mskAb.Enabled = False
mskEn.Enabled = False
cmbModelo.BackColor = PnLivro.BackColor
mskAb.BackColor = PnLivro.BackColor
mskEn.BackColor = PnLivro.BackColor

cmdNovo3.Visible = True
cmdAlterar3.Visible = True
cmdExcluir3.Visible = True
cmdGravar3.Visible = False
cmdCancelar3.Visible = False

End Sub


Private Sub cmdGravarNF_Click()
Dim Sql As String, RdoAux As rdoResultset

If Val(txtValorNF.Text) > 0 And chkSemMov.value = vbChecked Then
    MsgBox "Declaração sem movimento não pode ter valor.", vbCritical, "Atenção"
    Exit Sub
End If

If Trim(txtValorNF.Text) = "" Then txtValorNF.Text = "0"
If Trim(txtValorNF.Text) = "SEM MOVIMENTO" Then txtValorNF.Text = "0"
If Trim(txtValorNF.Text) = "SEM DECLARAÇÃO" Then txtValorNF.Text = "0"
Sql = "SELECT * from declaracaoiss where codigo=" & Val(txtCodEmpresa.Text)
Sql = Sql & " AND ANO=" & Val(cmbAnoISS.Text) & " AND MES=" & Val(lblMesNF.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        Sql = "INSERT DECLARACAOISS(CODIGO,ANO,MES,VALOR,SMOV) VALUES(" & Val(txtCodEmpresa.Text) & ","
        Sql = Sql & Val(cmbAnoISS.Text) & "," & Val(lblMesNF.Caption) & "," & Virg2Ponto(RemovePonto(txtValorNF.Text)) & ","
        Sql = Sql & chkSemMov.value & ")"
    Else
        Sql = "UPDATE DECLARACAOISS SET VALOR=" & Virg2Ponto(RemovePonto(txtValorNF.Text)) & ",SMOV=" & chkSemMov.value & " WHERE "
        Sql = Sql & "CODIGO=" & Val(txtCodEmpresa.Text) & " AND ANO=" & Val(cmbAnoISS.Text) & " AND MES=" & Val(lblMesNF.Caption)
    End If
    cn.Execute Sql, rdExecDirect
   .Close
End With

If chkSemMov.value = vbChecked Then
    lblIss(lblMesNF.Caption).Caption = "SEM MOVIMENTO"
Else
    lblIss(lblMesNF.Caption).Caption = FormatNumber(txtValorNF.Text, 2)
End If

frISS.Visible = False
End Sub

Private Sub cmdGravarPic_Click()
Dim ret As Long, sPathDestino As String, sAno As String, sFile As String, sFullArq As String, sTipo As String, sSeq As String, sCod As String
Dim FS As FileSystemObject, sExt As String
If Val(txtCodEmpresa.Text) = 0 Then Exit Sub
Set FS = New FileSystemObject
If cmbTipoDoc.ListIndex = -1 Then
    MsgBox "Selecione um tipo de documento.", vbCritical, "Atenção"
    Exit Sub
End If

If txtArq.Text = "" Then
    MsgBox "Selecione um arquivo.", vbCritical, "Atenção"
    Exit Sub
End If

If txtCodEmpresa.Text = "" Then
    MsgBox "É necessário que a empresa esteja criada antes de poder salvar documentos para ela.", vbCritical, "Atenção"
    Exit Sub
End If

If MsgBox("Deseja importar o arquivo " & txtArq.Text & "?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
sPathDestino = "\\192.168.200.130\atualizagti\Documentos\"
sAno = cmbAno.Text

If Not (FS.FolderExists(sPathDestino & sAno)) Then
    FS.CreateFolder (sPathDestino & sAno)
End If

sTipo = Format(cmbTipoDoc.ItemData(cmbTipoDoc.ListIndex), "00")
sSeq = Format(UBound(aPic) + 1, "00")
sCod = Mid(txtCodEmpresa.Text, 2, 6)
sExt = Right(txtArq.Text, 3)

sFile = sTipo & sAno & sSeq & sCod & "." & sExt
sFullArq = sPathDestino & sAno & "\" & sFile

FS.CopyFile txtArq.Text, sFullArq, True

Sql = "select max(seq) as maximo from documentopic"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
sSeq = Val(RdoAux!maximo) + 1

Sql = "insert documentopic(seq,codigo,documento) values(" & Val(sSeq) & "," & Val(sCod) & ",'" & sFile & "')"
cn.Execute Sql, rdExecDirect

ReDim Preserve aPic(UBound(aPic) + 1)
aPic(UBound(aPic)).sTipo = sTipo
aPic(UBound(aPic)).sArq = sFile
aPic(UBound(aPic)).sExt = sExt
aPic(UBound(aPic)).sTipoExt = cmbTipoDoc.Text
aPic(UBound(aPic)).sAno = cmbAno.Text
nPointer = UBound(aPic)
LoadPic

End Sub

Private Sub cmdIEExpande_Click()
frIE.Width = 10900
frIE.Left = 270
grdMain.Width = frIE.Width - 80
End Sub

Private Sub cmdIEReduz_Click()
frIE.Left = 6345
frIE.Width = 4965
grdMain.Width = frIE.Width - 80
End Sub

Private Sub cmdISSEletro_Click()
If Val(txtCodEmpresa.Text) > 0 Then frmPeriodoIE.show vbModal
IECheck

End Sub

Private Sub cmdNovo_Click()
    AtivaTela (0)
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
    cmbUF.ListIndex = 24
    cmbCidade.Text = "JABOTICABAL"
    mskCPF.SetFocus
End Sub

Private Sub cmdNovo3_Click()

lblNovo2.Caption = "1"


Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQLIV'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nSeq = Format(!valparam + 1, "000000")
   .Close
End With
lblSeq2.Caption = Format(nSeq, "0000")

cmbModelo.ListIndex = -1
LimpaMascara mskEn
LimpaMascara mskAb

cmbModelo.Enabled = True
mskAb.Enabled = True
mskEn.Enabled = True
cmbModelo.BackColor = Branco
mskAb.BackColor = Branco
mskEn.BackColor = Branco
cmdNovo3.Visible = False
cmdAlterar3.Visible = False
cmdExcluir3.Visible = False
cmdGravar3.Visible = True
cmdCancelar3.Visible = True

End Sub

Private Sub cmdOpenPic_Click()

Dim fName As String, cc As cCommonDlg
Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Todos os Arquivos|*.*", , App.Path & "\Bin", "Selecione um arquivo para importação", , Me.HWND, OFN_HIDEREADONLY, False
If fName <> "" Then
    txtArq.Text = fName
End If

End Sub

Private Sub cmdOutrasAtiv_Click()
If txtAtiv.Text = "" Then
   MsgBox "Selecione a atividade principal.", vbExclamation, "Atenção"
   Exit Sub
Else
   frmOutraAtividade.show
End If
End Sub

Private Sub cmdPrint_Click()
PopupMenu mnuPrint

End Sub

Private Sub cmdPrintIssAno_Click()
Dim Sql As String, RdoAux As rdoResultset, nPosRow As Long, nTotalRows As Long, sIdPrestador As String, nNumdoc As Long, nValorPago As Double
Dim sTipoNota As String, sStatusNota As String, sNatureza As String, sRecolhe As String, aNotaNum() As Long, nFindNota As Long
Dim sAtiv As String, sAliq As String, sTipoTomador As String, sIdTomador, bHasRows As Boolean, sCNPJPrestador As String
Dim sPago As String, sGuia As String, x As Integer, RdoAux2 As rdoResultset, sCNPJTomador As String

'If grdMain.Rows = 0 Then
'    MsgBox "Nada a imprimir.", vbExclamation, "Atenção"
'    Exit Sub
'End If

ReDim aNota(0): ReDim aNotaNum(0)
bHasRows = False
nPosRow = 1
Ocupado
Me.Refresh
Sql = "delete from extratonf where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "delete from extratonfresumo where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT nfisseletro2.identificaprestador, nfisseletro2.tipoprestador, nfisseletro2.tiponota, nfisseletro2.numeronota, nfisseletro2.serie, nfisseletro2.dataemissao,"
Sql = Sql & "nfisseletro2.mesref, nfisseletro2.anoref, nfisseletro2.statusnota, nfisseletro2.datacancel, nfisseletro2.natureza, nfisseletro2.valortotal, nfisseletro2.valorservico,"
Sql = Sql & "nfisseletro2.valorimposto, nfisseletro2.recolhimento, nfisseletro2.atividade, nfisseletro2.aliquota, nfisseletro2.razaoprestador, nfisseletro2.cidadeprestador,"
Sql = Sql & "nfisseletro2.ufprestador, nfisseletro2.localprestador, nfisseletro2.identificatomador, nfisseletro2.tipotomador, nfisseletro2.razaotomador, nfisseletro2.cidadetomador,"
Sql = Sql & "nfisseletro2.UFTomador , nfisseletro2.LocalTomador, nfisseletro2.NumDoc, nfisseletro2.simplesnac from nfisseletro2 "
Sql = Sql & "WHERE ((IDENTIFICAPRESTADOR='" & txtCodIss.Text & "' OR IDENTIFICAPRESTADOR='" & Format(txtCodIss.Text, "00000000000000") & "') or (IDENTIFICATOMADOR='" & txtCodIss.Text & "' OR IDENTIFICATOMADOR='" & Format(txtCodIss.Text, "00000000000000") & "'))"
Sql = Sql & "and AnoRef = " & Val(cmbAnoISS.Text)

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotalRows = .RowCount
    nPosRow = 1
    Do Until .EOF
        bHasRows = True
        If nPosRow Mod 20 = 0 Then
            CallPb nPosRow, nTotalRows
        End If
        If Val(!IdentificaTomador) = Val(txtCodIss.Text) Then
            If !Recolhimento <> 2 And !Recolhimento <> 4 Then
                GoTo proximo
            Else
                If !TipoNota = 1 Then
                    GoTo proximo
                End If
            End If
        Else
            If !TipoNota = 2 Then
                GoTo proximo
            End If
        
        End If

        Sql = "select codigomob,cnpj from mobiliario where codigomob=" & Val(!IdentificaTomador)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sCNPJTomador = Format(SubNull(RdoAux2!Cnpj), "0#\.###\.###/####-##")
        Else
            If Len(SubNull(!IdentificaTomador)) = 14 Then
                sCNPJTomador = Format(!IdentificaTomador, "0#\.###\.###/####-##")
            Else
                sCNPJTomador = ""
            End If
        End If
        RdoAux2.Close
        
        Sql = "select codigomob,cnpj from mobiliario where codigomob=" & Val(!IdentificaPrestador)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sCNPJPrestador = Format(SubNull(RdoAux2!Cnpj), "0#\.###\.###/####-##")
        Else
            sCNPJPrestador = Format(!IdentificaPrestador, "0#\.###\.###/####-##")
        End If
        RdoAux2.Close
        
        

        If optRec(0).value = True Then
            If !TipoNota = 2 Then GoTo proximo
        Else
            If !TipoNota = 1 Then GoTo proximo
        End If


        If !TipoNota = 1 Then
            sTipoNota = "Emi"
        Else
            sTipoNota = "Rec"
        End If
        
        If !StatusNota = 0 Then
            sStatusNota = "Rec"
        ElseIf !StatusNota = 1 Then
            sStatusNota = "Nor"
        Else
            sStatusNota = "Can"
        End If
        
        If !Natureza = 1 Then
            sNatureza = "Srv"
        Else
            sNatureza = "Mst"
        End If
        
        If !TipoNota = 1 And !Recolhimento = 0 Then
            sRecolhe = "Isento"
        ElseIf !TipoNota = 1 And !Recolhimento = 1 Then
            sRecolhe = "Retido"
        ElseIf !TipoNota = 1 And !Recolhimento = 2 Then
            sRecolhe = "A Recolher"
        ElseIf !TipoNota = 1 And !Recolhimento = 3 Then
            sRecolhe = "Simples"
        ElseIf !TipoNota = 2 And !Recolhimento = 1 Then
            sRecolhe = "Disp.Ret"
        ElseIf !TipoNota = 2 And !Recolhimento = 2 Then
            sRecolhe = "Ret.Sub.Trib"
        ElseIf !TipoNota = 2 And !Recolhimento = 3 Then
            sRecolhe = "Ret.Res.Trib"
        End If
        
        If !TipoTomador = 0 Then
            sTipoTomador = "CNPJ"
        ElseIf !TipoTomador = 1 Then
            sTipoTomador = "CPF"
        ElseIf !TipoTomador = 2 Then
            sTipoTomador = "IM"
        End If
        sIdTomador = Format(!IdentificaTomador, "00000000000000")
        sIdPrestador = Format(!IdentificaPrestador, "00000000000000")
        sAtiv = SubNull(!Atividade)
        sAliq = Format(!Aliquota, "#0.00") & "%"
        
        sGuia = SubNull(!NumDoc)
        sPago = "N"
        If Val(sGuia) > 0 Then
           ' If Not IsNull(!valorpagoreal) Then
                Sql = "select valorpagoreal from debitopago where numdocumento=" & nNumdoc
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                nValorPago = RdoAux2!valorpagoreal
                RdoAux2.Close
                If nValorPago > 0 Then
                    sPago = "S"
                End If
            'End If
        End If
        
'        For x = 1 To UBound(aNota)
 '           If aNota(x) = !NumeroNota Then
 '               GoTo Proximo
 '           End If
'        Next
        If !NumeroNota > 10000 Then GoTo proximo
        If isInLongArray(aNotaNum, !NumeroNota) Then
           For x = nIndexFind To UBound(aNota)
               If aNota(x).NumeroNota = !NumeroNota And aNota(x).Serie = !Serie Then
                   GoTo proximo
               End If
            Next
        End If
        
        ReDim Preserve aNotaNum(UBound(aNotaNum) + 1)
        aNotaNum(UBound(aNotaNum)) = !NumeroNota
        
        ReDim Preserve aNota(UBound(aNota) + 1)
        aNota(UBound(aNota)).NumeroNota = !NumeroNota
        aNota(UBound(aNota)).Serie = !Serie
        
        Sql = "insert extratonf(usuario,idprestador,tipoprestador,anoref,mesref,tiponota,nota,serie,dataemissao,stiponota,"
        Sql = Sql & "statusnota,datacancel,natureza,valortotal,valorservico,valorimposto,recolhimento,atividade,aliq,razaoprestador,"
        Sql = Sql & "cidadeprestador,ufprestador,idtomador,tipotomador,razaotomador,cidadetomador,uftomador,guia,pago,cnpjtomador,cnpjprestador) values('"
        Sql = Sql & NomeDeLogin & "','" & sIdPrestador & "'," & !TipoPrestador & "," & !AnoRef & "," & !MesRef & "," & !TipoNota & ",'"
        Sql = Sql & !NumeroNota & "','" & !Serie & "','" & Format(!DataEmissao, "mm/dd/yyyy") & "','" & sTipoNota & "','" & sStatusNota & "',"
        Sql = Sql & IIf(Not IsNull(!DataCancel), "'" & Format(!DataCancel, "mm/dd/yyyy") & "'", "Null") & ",'" & sNatureza & "',"
        Sql = Sql & Virg2Ponto(Format(!ValorTotal, "#0.00")) & "," & Virg2Ponto(Format(!ValorServico, "#0.00")) & "," & Virg2Ponto(Format(!ValorImposto, "#0.00")) & ",'"
        Sql = Sql & sRecolhe & "','" & Left(sAtiv, 5) & "','" & sAliq & "','" & Mask(!RazaoPrestador) & "','" & UCase(Mask(!CidadePrestador)) & "','" & !UFPrestador & "','"
        Sql = Sql & sIdTomador & "','" & sTipoTomador & "','" & Mask(!RazaoTomador) & "','" & Mask(!CidadeTomador) & "','" & !UFTomador & "','" & sGuia & "','" & sPago & "','"
        Sql = Sql & sCNPJTomador & "','" & sCNPJPrestador & "')"
        cn.Execute Sql, rdExecDirect
proximo:
        nPosRow = nPosRow + 1
       .MoveNext
       DoEvents
    Loop
   .Close
End With

PBar.value = 0
PBar.Color = vbWhite

If Not bHasRows Then
    Sql = "insert extratonf(usuario,idprestador,tipoprestador,anoref,mesref,tiponota,nota,serie,razaoprestador) values('" & NomeDeLogin & "','" & RetornaNumero(mskCNPJ.Text) & "',"
    Sql = Sql & "0," & cmbAnoISS.Text & ",0,0,'','','" & Mask(txtRazao.Text) & "')"
    cn.Execute Sql, rdExecDirect
End If

Sql = "insert extratonfresumo(usuario,e01,e02,e03,e04,e05,e06,e07,e08,e09,e10,e11,e12,etot,r01,r02,r03,r04,r05,r06,r07,r08,r09,r10,r11,r12,rtot) values('"
Sql = Sql & NomeDeLogin & "','" & lblIss(1).Caption & "','" & lblIss(2).Caption & "','" & lblIss(3).Caption & "','" & lblIss(4).Caption & "','" & lblIss(5).Caption
Sql = Sql & "','" & lblIss(6).Caption & "','" & lblIss(7).Caption & "','" & lblIss(8).Caption & "','" & lblIss(9).Caption & "','" & lblIss(10).Caption
Sql = Sql & "','" & lblIss(11).Caption & "','" & lblIss(12).Caption & "','" & lblTotISSE.Caption & "','" & lblIss(13).Caption & "','" & lblIss(14).Caption
Sql = Sql & "','" & lblIss(15).Caption & "','" & lblIss(16).Caption & "','" & lblIss(17).Caption & "','" & lblIss(18).Caption & "','" & lblIss(19).Caption
Sql = Sql & "','" & lblIss(20).Caption & "','" & lblIss(21).Caption & "','" & lblIss(22).Caption & "','" & lblIss(23).Caption & "','" & lblIss(24).Caption & "','" & lblTotISSR.Caption & "')"
cn.Execute Sql, rdExecDirect

Liberado

frmReport.ShowReport2 "EXTRATONF", frmMdi.HWND, Me.HWND

Sql = "delete from extratonf where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "delete from extratonfresumo where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdRefreshBairro_Click()
cmbEEBairro.Clear
cmbEECidade_Click
End Sub

Private Sub cmdRefreshCity_Click()
cmbEEBairro.Clear
cmbEECidade.Clear
cmbEEUF_Click
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSairCNAE_Click()
frCNAE.Visible = False
cmdAddAtiv.Enabled = True
cmdAddISS.Enabled = True
cmdAddVS.Enabled = True
cmdDelISS.Enabled = True
cmdDelVS.Enabled = True
cmdGravar.Enabled = True
cmdCancel.Enabled = True
cmdOutrasAtiv.Enabled = True
cmdCancel.Enabled = True
cmdAddCnae1.Enabled = True
cmdAddCnae2.Enabled = True
cmdDelCnae.Enabled = True
cmdDelCnaeP.Enabled = True
cmdEdtCnae.Enabled = True
btAtivExtenso.Enabled = True
Tela(4).Enabled = True
'TabMob.ShowTabs = True

End Sub

Private Sub cmdSelect_Click()

If cmbCriterio.ListIndex = -1 Then
    MsgBox "Selecione o critério.", vbCritical, "Atenção"
    Exit Sub
End If

If CDbl(lblValorCnae.Caption) = 0 Then
    MsgBox "Critério sem valor.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtQtdeCnae.Text) = 0 Then
    MsgBox "Digite a quantidade.", vbCritical, "Atenção"
    Exit Sub
End If

grdVS.AddItem mskCodCNAE.Text & Chr(9) & cmbCriterio.ItemData(cmbCriterio.ListIndex) & Chr(9) & UCase(lblDescCNAE.Caption) & " - " & cmbCriterio.Text & Chr(9) & Val(txtQtdeCnae.Text) & Chr(9) & lblValorCnae.Caption

cmdSairCNAE_Click
End Sub

Private Sub cmdSil_Click()
If Val(txtCodEmpresa.Text) = 0 Then
   MsgBox "Selecione uma empresa cadastrada.", vbCritical, "Atenção"
   Exit Sub
Else
'    MsgBox "em manutenção"
    frmSIL.show vbModal
End If

End Sub

Private Sub cmdVSOld_Click()
If grdVS2.Visible = False Then
    grdVS2.Visible = True
    grdVS2.ZOrder 0
Else
    grdVS2.Visible = False
End If
End Sub


Private Sub Form_Activate()
If Val(CodEmpresa) > 0 Then
   Ocupado
   AtivaTela (0)
   txtCodEmpresa.Text = CodEmpresa
   Le
End If
CodEmpresa = 0
bExec = True
Liberado
End Sub

Private Sub Form_Load()
Dim x As Integer
ReDim aPic(0)
ReDim aTipoDoc(0)
ReDim aNota(0)
For x = 2008 To Year(Now)
    cmbAnoISS.AddItem x
Next
cmbAnoISS.ListIndex = 0
GridHeader
grdMain.ZOrder 0
sCH = ""
MontaMenuISS
AtivaTela (0)
'CodEmpresa = ""
Centraliza Me
sRet = RetEventUserForm(Me.Name)
Eventos "INICIAR"
CarregaCombo
frAtivo = 0
Tela(0).ZOrder 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmAtivISS
Unload frmVigSanitaria
Unload frmAtiv
CodEmpresa = 0
Set m_cMenuISS = Nothing

End Sub

Private Sub grdHist_Click()
If grdHist.Rows = 1 Then
   txtHist.Text = ""
   Exit Sub
End If
If grdHist.Row > 0 Then
    txtHist.Text = grdHist.TextMatrix(grdHist.Row, 2)
End If
End Sub

Private Sub FormHagana()
bNew = False
bEdit = False
bDel = False
evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

If NomeDeLogin <> "LUIZH" And NomeDeLogin <> "MARILIA" And NomeDeLogin <> "PAULA" And NomeDeLogin <> "NOELI" And NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "LEANDRO" And NomeDeLogin <> "PAULO" And NomeDeLogin <> "ROSANGELA" And NomeDeLogin <> "LEONARDODINIZ" And NomeDeLogin <> "EDUARDO" Then
    btMenu(9).Enabled = False
Else
    btMenu(9).Enabled = True
End If


End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdConsultar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdPrint.Visible = True
   cmdExtrato.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmCadMob
       If TypeOf Ct Is TextBox Or TypeOf Ct Is esMaskedEdit Or TypeOf Ct Is ComboBox Or TypeOf Ct Is CommandButton Then
         Ct.BackColor = Kde
         If TypeOf Ct Is CommandButton Then
            Ct.Enabled = False
         Else
            Ct.Locked = True
         End If
       End If
   Next
   cmbAnoISS.Enabled = True
   cmbAnoISS.Locked = False
   cmbAnoISS.BackColor = Branco
   cmbPagto.Enabled = True
   cmbPagto.Locked = False
   cmbPagto.BackColor = Branco
   chkEnd.Enabled = False
   cmdFoto.Enabled = True
   cmdAlterar3.Enabled = False
   cmdNovo3.Enabled = False
   cmdExcluir3.Enabled = False
   cmdGravar3.Visible = False
   cmdCancelar3.Visible = False
   chkVistoria.Enabled = False
   btAtivExtenso.Enabled = False
'   chkRE.Enabled = False
'   chkMei.Enabled = False
'   chkEmiteNF.Enabled = False
   chkIE.Enabled = False
   'chkAlvara.Enabled = False
   chkIsentoISS.Enabled = False
   chkIsentoTaxa.Enabled = False
   chkInscTemp.Enabled = False
   chkSubstitutoTributario.Enabled = False
   chkLiberadoVRE.Enabled = False
'   chkDanfe.Enabled = False
   chk24horas.Enabled = False
   chkBombon.Enabled = False
   chkEmpInd.Enabled = False
   cmdEditHist.Enabled = False
   cmdAddCnae1.Enabled = True
   cmdEdtCnae.Enabled = True
   cmbCnae.Enabled = True
    cmdAddAtiv.Enabled = False
    cmdAddISS.Enabled = False
    cmdAddVS.Enabled = False
    cmdDelISS.Enabled = False
    cmdDelVS.Enabled = False
    cmdGravar.Enabled = False
    cmdCancel.Enabled = False
    cmdCancel.Enabled = False
    cmdAddCnae1.Enabled = False
    cmdAddCnae2.Enabled = False
    cmdDelCnae.Enabled = False
    cmdDelCnaeP.Enabled = False
    cmdEdtCnae.Enabled = False
    txtValorNF.Locked = False
    btAddPlaca.Enabled = False
    btDelPlaca.Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdConsultar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdPrint.Visible = False
   cmdExtrato.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmCadMob
       If TypeOf Ct Is TextBox Or TypeOf Ct Is esMaskedEdit Or TypeOf Ct Is ComboBox Or TypeOf Ct Is CommandButton Then
          Ct.BackColor = vbWhite
         If TypeOf Ct Is CommandButton Then
            Ct.Enabled = True
         Else
            Ct.Locked = False
         End If
       End If
   Next
   
   If IsDate(mskDataEn.Text) Then
        txtNumProcE.Locked = True
        txtNumProcE.BackColor = Kde
        mskDataEn.Locked = True
        mskDataEn.BackColor = Kde
        mskDataPEn.Locked = True
        mskDataPEn.BackColor = Kde
   End If
   txtFoneCont.Locked = True
   txtFoneCont.BackColor = Kde
   txtEmailCont.Locked = True
   txtEmailCont.BackColor = Kde
   chkEnd.Enabled = True
   txtAtiv.BackColor = Kde
   txtAtiv.Locked = True
   txtValorAliq.BackColor = Kde
   txtValorAliq.Locked = True
   txtCodEsc.BackColor = Kde
   txtCodEsc.Locked = True
   txtNomeProf.BackColor = Kde
   txtNomeProf.Locked = True
   cmdAlterar3.Enabled = False
   cmdNovo3.Enabled = True
   cmdExcluir3.Enabled = True
   chkVistoria.Enabled = True
   chkIsentoISS.Enabled = True
   chkIsentoTaxa.Enabled = True
   chkLiberadoVRE.Enabled = True
   btAtivExtenso.Enabled = True
'   chkRE.Enabled = True
'   chkMei.Enabled = False
'   chkEmiteNF.Enabled = True
   chkIE.Enabled = True
   'chkAlvara.Enabled = True
   chkInscTemp.Enabled = True
   chkSubstitutoTributario.Enabled = True
'   chkDanfe.Enabled = True
   chk24horas.Enabled = True
   chkBombon.Enabled = True
   chkEmpInd.Enabled = True
   cmdEditHist.Enabled = True
cmdAddAtiv.Enabled = True
cmdAddISS.Enabled = True
cmdAddVS.Enabled = True
cmdDelISS.Enabled = True
cmdDelVS.Enabled = True
cmdGravar.Enabled = True
cmdCancel.Enabled = True
cmdCancel.Enabled = True
cmdAddCnae1.Enabled = True
cmdAddCnae2.Enabled = True
cmdDelCnae.Enabled = True
cmdDelCnaeP.Enabled = True
cmdEdtCnae.Enabled = True
btAddPlaca.Enabled = True
btDelPlaca.Enabled = True
End If

cmbModelo.Enabled = False
mskAb.Enabled = False
mskEn.Enabled = False
cmbModelo.BackColor = PnLivro.BackColor
mskAb.BackColor = PnLivro.BackColor
mskEn.BackColor = PnLivro.BackColor
txtCodEmpresa.Locked = True
txtCodEmpresa.BackColor = Kde
txtHist.Locked = True
txtHist.BackColor = Kde

mskDataPAb.BackColor = Kde
mskDataPEn.BackColor = Kde
mskDataPAb.Locked = True
mskDataPEn.Locked = True

txtNomeISS.BackColor = Tela(9).BackColor
txtCodIss.BackColor = Branco
txtCodIss.Locked = False
FormHagana

If NomeDeLogin <> "LUIZH" And NomeDeLogin <> "RENATA" And NomeDeLogin <> "MARILIA" And NomeDeLogin <> "EDUARDO" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "IMCAVAGNA" And NomeDeLogin <> "PAULA" And NomeDeLogin <> "NOELI" And NomeDeLogin <> "SCHWARTZ" Then
    txtSenha.Locked = True
End If
If NomeDeLogin <> "RITA" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "SCHWARTZ" Then
    chkLiberadoVRE.Enabled = False
End If


txtArq.Enabled = True
cmbTipoDoc.Locked = False
cmbAno.Locked = False
txtArq.BackColor = Branco
cmbTipoDoc.BackColor = Branco
cmbAno.BackColor = Branco

End Sub

Private Sub CarregaCombo()
Sql = "SELECT SIGLAUF,DESCUF FROM UF ORDER BY DESCUF; " & _
      "SELECT CODHORARIO,DESCHORARIO FROM HORARIOFUNC ORDER BY DESCHORARIO; " & _
      "SELECT CODIGOESC,NOMEESC FROM ESCRITORIOCONTABIL WHERE CODIGOESC>0 AND NOMEESC<>'' ORDER BY NOMEESC; " & _
      "SELECT CODIGO,NOME FROM TIPODOCUMENTO ORDER BY CODIGO;"

Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbUF.AddItem !DESCUF & " (" & !SiglaUF & ")"
       cmbEEUf.AddItem !DESCUF & " (" & !SiglaUF & ")"
      .MoveNext
    Loop
   .MoreResults
    cmbHorario.AddItem ""
    Do Until .EOF
       cmbHorario.AddItem !DESCHORARIO
       cmbHorario.ItemData(cmbHorario.NewIndex) = !CODHORARIO
      .MoveNext
    Loop
   .MoreResults
    cmbNomeEsc.AddItem ""
    Do Until .EOF
       cmbNomeEsc.AddItem !NOMEESC
       cmbNomeEsc.ItemData(cmbNomeEsc.NewIndex) = !codigoesc
      .MoveNext
    Loop
    .MoreResults
    Do Until .EOF
       ReDim Preserve aTipoDoc(UBound(aTipoDoc) + 1)
       aTipoDoc(UBound(aTipoDoc)).sNome = !Nome
       aTipoDoc(UBound(aTipoDoc)).nCod = !Codigo
       cmbTipoDoc.AddItem !Nome
       cmbTipoDoc.ItemData(cmbTipoDoc.NewIndex) = !Codigo
      .MoveNext
    Loop

   .Close

    For x = 1995 To Year(Now) + 1
        cmbAno.AddItem x
    Next
    cmbAno.Text = Year(Now)
    
End With
End Sub

Private Sub grdLivro_Click()

With grdLivro
    lblSeq2.Caption = .TextMatrix(.Row, 0)
    cmbModelo.Text = .TextMatrix(.Row, 1)
    mskAb.Text = .TextMatrix(.Row, 2)
    If .TextMatrix(.Row, 3) <> "" Then
        mskEn.Text = .TextMatrix(.Row, 3)
    Else
        LimpaMascara mskEn
    End If
End With

End Sub

Private Sub grdProc_DblClick()
Dim nAno As Integer, nNumproc As Integer, sNumProcesso As String
If grdProc.Row = 0 Then Exit Sub
sNumProcesso = grdProc.TextMatrix(grdProc.Row, 0)
nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
nNumproc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1))

AnoProcesso = nAno
CodProcesso = nNumproc
frmProcesso.show
frmProcesso.ZOrder 0

End Sub

Private Sub grdVS_DblClick()
 If grdVS.Rows > 1 Then
    MsgBox grdVS.TextMatrix(grdVS.Row, 2)
 End If

End Sub

Private Sub imgTmp_DblClick()
cmdAbrirPic_Click
End Sub

Private Sub lblIss_Click(Index As Integer)
Dim x As Integer
If frISS.Visible = True Then Exit Sub
If Index > 12 Then Index = Index - 12
For x = 1 To 12
    If x = Index Then
        lblIss(x).BackColor = &H80FFFF
        lblIss(x + 12).BackColor = &H80FFFF
    Else
        lblIss(x).BackColor = vbWhite
        lblIss(x + 12).BackColor = vbWhite
    End If
Next

frIE.Caption = "Notas do Mês de " & RemovePonto(lblMes(Index).Caption)
lblMesNF.Caption = Index
If chkHabilitar.value = vbChecked Then
    CarregaLista
End If
End Sub

Private Sub lblIss_DblClick(Index As Integer)
'If Val(txtCodEmpresa.Text) > 0 And (NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "JONATHAN") And frISS.Visible = False Then
'    If frISS.Visible = True Then Exit Sub
'    If Index > 12 Then Index = Index - 12
'
'    frISS.Caption = "Alteração da Nota de " & lblMes(Index).Caption
'    frISS.Visible = True: frISS.ZOrder 0
    lblMesNF.Caption = Index
'    If lblIss(Index).Caption = "SEM MOVIMENTO" Then
'        txtValorNF.Text = "0,00"
'        chkSemMov.Value = vbChecked
'    Else
'        txtValorNF.Text = lblIss(Index).Caption
'        chkSemMov.Value = vbUnchecked
'    End If
'    txtValorNF.SetFocus
'End If

End Sub

Private Sub lblMes_Click(Index As Integer)
lblIss_Click (Index)
End Sub

Private Sub lstEENomeLog_DblClick()
If lstEENomeLog.ListIndex > -1 Then
   txtEECodLogr.Text = lstEENomeLog.ItemData(lstEENomeLog.ListIndex)
   txtEECodLogr_LostFocus
   lstEENomeLog.Visible = False
   If txtEENumero.Enabled = True Then txtEENumero.SetFocus
End If

End Sub

Private Sub lstEENomeLog_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    If lstEENomeLog.ListIndex > -1 Then
       txtEECodLogr.Text = lstEENomeLog.ItemData(lstEENomeLog.ListIndex)
       txtEECodLogr_LostFocus
       lstEENomeLog.Visible = False
       txtEENumero.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstEENomeLog.Visible = False
   txtEENomeLogr.SetFocus
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   If txtNumero.Enabled = True Then txtNumero.SetFocus
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
       txtNumero.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogr.SetFocus
End If

End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub


Private Sub lvTmp_ItemClick(ByVal Item As MSComctlLib.ListItem)
If lvTmp.ListItems.Count > 0 Then
    LoadPic
End If
End Sub

Private Sub m_cMenuISS_Click(ItemNumber As Long)
Select Case m_cMenuISS.ItemKey(ItemNumber)
    Case "mnuISSF"
        lblTipoISS.Caption = "--> 11 - ISS FIXO"
    Case "mnuISSE"
        lblTipoISS.Caption = "--> 12 - ISS ESTIMADO"
    Case "mnuISSV"
        lblTipoISS.Caption = "--> 13 - ISS VARIÁVEL"
End Select

TipoISS = ""

Set frm = frmAtivISS
frm.nTipo = Val(Mid$(lblTipoISS.Caption, 5, 2))
frm.sForm = "frmCadMob"
frmAtivISS.show vbModeless
frmAtivISS.ZOrder 0

End Sub


Private Sub mnuDeca_Click()
If Val(txtCodEmpresa.Text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

frmReport.ShowReport3 "DECA", frmMdi.HWND, Me.HWND

End Sub

Private Sub mnuDecaV_Click()
If Val(txtCodEmpresa.Text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

frmReport.ShowReport3 "DECA2", frmMdi.HWND, Me.HWND

End Sub

Private Sub mnuFicha_Click()

Dim Sql As String

If Val(txtCodEmpresa.Text) = 0 Then
   MsgBox "Não existem Registros.", vbCritical, "Atenção"
   Exit Sub
End If

Sql = "DELETE FROM TBDADOSEMPRESA WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "INSERT TBDADOSEMPRESA(USUARIO,CODIGO) VALUES('" & NomeDeLogin & "'," & Val(txtCodEmpresa.Text) & ")"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport "CADMOBILIARIO", frmMdi.HWND, Me.HWND

Sql = "DELETE FROM TBDADOSEMPRESA WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mskAb_GotFocus()
mskAb.SetFocus
End Sub

Private Sub mskCnae_GotFocus()
On Error Resume Next
mskCnae.SelStart = 0
mskCnae.SelLength = Len(mskCnae.Text)
mskCnae.SetFocus
End Sub

Private Sub mskCNPJ_GotFocus()
mskCNPJ.SelStart = 0
mskCNPJ.SelLength = Len(mskCNPJ.Text)
End Sub

Private Sub mskCodCNAE_Change()
Dim sCnae As String, nDivisao As Integer, nGrupo As Integer, sClasse As String, nClasse As Integer, nSubClasse As Integer
lblDescCNAE.Caption = "": cmbCriterio.Clear
If Len(mskCodCNAE.ClipText) = 7 Then
    sCnae = RetornaNumero(mskCodCNAE.Text)
'    nDivisao = Val(Left(sCnae, 2))
'    nGrupo = Val(Mid(sCnae, 3, 1))
'    sClasse = Mid(sCnae, 4, 3)
'    sClasse = Left(sClasse, 1) & Right(sClasse, 1)
'    nClasse = Val(sClasse)
'    nSubClasse = Val(Right(sCnae, 2))
'
'    Sql = "SELECT * FROM CNAESUBCLASSE WHERE DIVISAO=" & nDivisao & " AND GRUPO=" & nGrupo & " AND CLASSE=" & nClasse & " AND SUBCLASSE=" & nSubClasse
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If .RowCount > 0 Then
'            lblDescCNAE.Caption = !Descricao
'        End If
'       .Close
'    End With
    
    Sql = "select * from cnae where cnae='" & sCnae & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            lblDescCNAE.Caption = !descricao
        End If
       .Close
    End With
    
    
    
    
    If lblDescCNAE.Caption <> "" Then
        Sql = "SELECT cnaecriterio.criterio,cnaecriteriodesc.descricao,cnaecriterio.valor From "
        Sql = Sql & "cnaecriteriodesc INNER JOIN cnaecriterio ON (cnaecriteriodesc.criterio = cnaecriterio.criterio) WHERE cnae='" & sCnae & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                cmbCriterio.AddItem !descricao
                cmbCriterio.ItemData(cmbCriterio.NewIndex) = !criterio
               .MoveNext
            Loop
           .Close
        End With
    End If
    
End If

End Sub

Private Sub mskCodCNAE_GotFocus()
mskCodCNAE.SelStart = 0
mskCodCNAE.SelLength = Len(mskCodCNAE.Text)
mskCodCNAE.SetFocus
End Sub

Private Sub mskCPF_GotFocus()
mskCPF.SelStart = 0
mskCPF.SelLength = Len(mskCPF.Text)
End Sub


Private Sub mskDataAP_GotFocus()
mskDataAP.SelStart = 0
mskDataAP.SelLength = Len(mskDataAP.Text)
End Sub

Private Sub mskEn_GotFocus()
mskEn.SetFocus
End Sub


Private Sub mskPlaca_GotFocus()
mskPlaca.SelStart = 0
mskPlaca.SelLength = Len(mskPlaca.Text)
mskPlaca.SetFocus
End Sub

Private Sub mskPlaca_LostFocus()
If mskPlaca.ClipText <> "" Then
    mskPlaca.Text = UCase(mskPlaca.Text)
    If Len(mskPlaca.ClipText) < 7 Then
        MsgBox "Nº de placa inválida.", vbCritical, "Erro"
        mskPlaca.SetFocus
    End If
End If
End Sub

Private Sub optRec_Click(Index As Integer)
If chkHabilitar.value = 1 Then
    CarregaISS
End If
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
If KeyAscii = 46 Then KeyAscii = 44

Tweak txtArea, KeyAscii, DecimalPositive
End Sub

Private Sub txtCapital_KeyPress(KeyAscii As Integer)
Tweak txtCapital, KeyAscii, DecimalPositive
End Sub

Private Sub txtCodEmpresa_GotFocus()
Dim s As String

If Val(txtCodEmpresa.Text) > 0 Then
    s = RetornaNumero(txtCodEmpresa.Text)
    txtCodEmpresa = Left$(s, Len(s) - 1) & Right$(s, 1)
End If

End Sub

Private Sub txtCodEmpresa_KeyPress(KeyAscii As Integer)
If Len(txtCodEmpresa.Text) = 6 Then
    KeyAscii = 0
    Exit Sub
End If
Tweak txtCodEmpresa, KeyAscii, IntegerPositive
End Sub

Private Sub Le()
Dim nCodigo As Long, RdoAux2 As rdoResultset, sCnae As String, sSecao As String, nDivisao As Integer, nGrupo As Integer, sClasse As String, nClasse As Integer, nSubClasse As Integer, nCriterio As Integer
Dim nValorAliq As Double, sTipoIss As String, sDesc As String, qd As New rdoQuery, sTmp As String, d As Long
nCodigo = Val(Left$(txtCodEmpresa.Text, 7))
Ocupado
grdVS2.Visible = False
Limpa
ReDim aLog(0)
Sql = "SELECT * FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtArea.Text = SubNull(!areatl)
        'dados da empresa
        txtCodEmpresa.Text = Format(!codigomob, "0000000")
        txtRazao.Text = !razaosocial
        txtImovel.Text = SubNull(!Imovel)
        txtFantasia.Text = SubNull(!NOMEFANTASIA)
        txtSIL.Text = SubNull(!Sil)
        txtPonto.Text = SubNull(!ponto_agencia)
        If Val(SubNull(!Cnpj)) > 0 Then
            'mskCNPJ.Text = Format(!Cnpj, "0#\.###\.###/####-##")
            mskCNPJ.Text = !Cnpj
        End If
        If Val(SubNull(!CPF)) > 0 Then
            mskCPF.Text = Format(RetornaNumero(!CPF), "00#\.###\.###-##")
        End If
        If Not IsNull(!rg) Then
           mskRG.Text = !rg
           txtOrgao.Text = SubNull(!ORGAO)
        End If
        txtInscEst.Text = SubNull(!inscestadual)
        If Not IsNull(!DataAbertura) Then
            mskDataAb.Text = Format(!DataAbertura, "dd/mm/yyyy")
        Else
            LimpaMascara mskDataAb
        End If
        txtNumProcA.Text = SubNull(!numprocesso)
        If Not IsNull(!DATAPROCESSO) Then
            mskDataPAb.Text = Format(!DATAPROCESSO, "dd/mm/yyyy")
        Else
            LimpaMascara mskDataPAb
        End If
        If Not IsNull(!dataencerramento) Then
            mskDataEn.Text = Format(!dataencerramento, "dd/mm/yyyy")
        Else
            LimpaMascara mskDataEn
        End If
        txtNumProcE.Text = SubNull(!NUMPROCENCERRAMENTO)
        If Not IsNull(!DATAPROCENCERRAMENTO) Then
            mskDataPEn.Text = Format(!DATAPROCENCERRAMENTO, "dd/mm/yyyy")
        Else
            LimpaMascara mskDataEn
        End If
        txtHorarioExt.Text = SubNull(!HORARIOEXT)
        
        'carrega log
        aLog(0).sDataAb = mskDataAb.Text
        aLog(0).sDataEn = mskDataEn.Text
        aLog(0).sNumProcAb = txtNumProcA.Text
        aLog(0).sNumProcEn = txtNumProcE.Text
        aLog(0).sDataProcAb = mskDataPAb.Text
        aLog(0).sDataProcEn = IIf(IsDate(mskDataPEn.Text), mskDataPEn.Text, "")
        
        If !Horario > 0 Then
           For d = 0 To cmbHorario.ListCount - 1
               If cmbHorario.ItemData(d) = !Horario Then
                  cmbHorario.ListIndex = d
                  Exit For
               End If
           Next
        Else
            cmbHorario.ListIndex = -1
        End If
        If Val(SubNull(!RESPCONTABIL)) > 0 Then
           For d = 0 To cmbNomeEsc.ListCount - 1
               If cmbNomeEsc.ItemData(d) = !RESPCONTABIL Then
                  cmbNomeEsc.ListIndex = d
                  Exit For
               End If
           Next
        Else
            cmbNomeEsc.ListIndex = -1
        End If
        If !VISTORIA = 0 Then
            chkVistoria.value = 0
        Else
            chkVistoria.value = 1
        End If
        frMei.Visible = IsMEI(nCodigo)
'        chkRE.value = Val(SubNull(!REGESPECIAL))
'        chkMei.value = IIf(IsMEI(nCodigo), 1, 0)
 '       If chkMei.value = vbChecked Then
  '          frMei.Visible = True
   '     Else
    '        frMei.Visible = False
     '   End If
'        chkEmiteNF.value = Val(SubNull(!EMITENF))
        chkIE.value = Val(SubNull(!ISSELETRO))
        'chkAlvara.value = Val(SubNull(!ALVARA))
        chkIsentoTaxa.value = Val(SubNull(!ISENTOTAXA))
        chkIsentoISS.value = Val(SubNull(!ISENTOISS))
        chkInscTemp.value = Val(SubNull(!INSCTEMP))
        If IsNull(!substituto_tributario_issqn) Then
            chkSubstitutoTributario.value = 0
        Else
            chkSubstitutoTributario.value = IIf(!substituto_tributario_issqn, 1, 0)
        End If
        If IsNull(!cadastro_vre) Then
            chkLiberadoVRE.value = 0
            chkLiberadoVRE.Visible = False
        Else
            If !cadastro_vre = 0 Then
                chkLiberadoVRE.value = 0
                chkLiberadoVRE.Visible = False
            Else
                If IsNull(!liberado_vre) Then
                    chkLiberadoVRE.value = 0
                Else
                    If !liberado_vre = 0 Then
                        chkLiberadoVRE.value = 0
                        chkLiberadoVRE.Visible = True
                    Else
                        chkLiberadoVRE.value = 1
                        chkLiberadoVRE.Visible = True
                    End If
                End If
            End If
        End If
        
        chkDanfe.value = Val(SubNull(!DANFE))
        If IsNull(!horas24) Then
            chk24horas.value = 0
        Else
            chk24horas.value = IIf(!horas24, 1, 0)
        End If
        If IsNull(!bombonieri) Then
            chkBombon.value = 0
        Else
            chkBombon.value = IIf(!bombonieri, 1, 0)
        End If
        If IsNull(!individual) Then
            chkEmpInd.value = 0
        Else
            chkEmpInd.value = IIf(!individual, 1, 0)
        End If
'        If chkAlvara.value = 1 Then
'            bAlvaraAutomatico = True
'        Else
'            bAlvaraAutomatico = False
'        End If
        If chkIsentoISS.value = 1 Then
            bIsentoIss = True
        Else
            bIsentoIss = False
        End If
        If chkIsentoTaxa.value = 1 Then
            bIsentoTaxa = True
        Else
            bIsentoTaxa = False
        End If
        
        'LOCALIZACAO
        d = 0
         bExec = False
        For d = 0 To cmbUF.ListCount - 1
            If Left$(Right$(cmbUF.List(d), 3), 2) = !SiglaUF Then
               cmbUF.ListIndex = d
               Exit For
            End If
        Next
        bExec = True
        cmbUF_Click
        bExec = False
        For d = 0 To cmbCidade.ListCount - 1
            If cmbCidade.ItemData(d) = !CodCidade Then
               cmbCidade.ListIndex = d
               Exit For
            End If
        Next
        bExec = True
        cmbCidade_Click
        If !CodBairro <> 999 Then
            bExec = False
            For d = 0 To cmbBairro.ListCount - 1
                If cmbBairro.ItemData(d) = !CodBairro Then
                   cmbBairro.ListIndex = d
                   Exit For
                End If
            Next
            bExec = True
        End If
        If !CodLogradouro > 0 Then
           txtCodLogr.Text = !CodLogradouro
           txtCodLogr_LostFocus
        Else
           txtCodLogr.Text = 0
           txtNomeLogr.Text = SubNull(!NomeLogr)
        End If
        txtNumero.Text = Val(SubNull(!Numero))
        txtCompl.Text = SubNull(!Complemento)
        If !CodCidade = 413 Then
            If !CodBairro = 96 Then
                If Not IsNull(!Cep) Then
                    mskCEP.Text = Format(!Cep, "00000-000")
                End If
            Else
                mskCEP.Text = RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!Numero)))
            End If
        Else
            mskCEP.Text = Format(!Cep, "00000-000")
        End If
       
        txtHP.Text = SubNull(!HOMEPAGE)
        'PROP/CONTATO
        Sql = "SELECT * FROM vwMOBILIARIOPROPRIETARIO WHERE CODMOBILIARIO=" & !codigomob
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                grdProp.AddItem Format(!CodCidadao, "000000") & Chr(9) & !nomecidadao & Chr(9) & SubNull(!CPF)
               .MoveNext
            Loop
           .Close
        End With
                
        txtFoneCont.Text = SubNull(!telefone)
        txtEmailCont.Text = SubNull(!Email)
        txtNomeContato.Text = SubNull(!NOMECONTATO)
        txtFone.Text = SubNull(!fonecontato)
        txtFax.Text = SubNull(!faxcontato)
        txtEmail.Text = SubNull(!emailcontato)
        txtCargo.Text = SubNull(!CARGOCONTATO)
        txtDDDNF.Text = SubNull(!ddd_nf)
        txtFoneNF.Text = SubNull(!telefone_nf)
        txtEmailNF.Text = SubNull(!email_nf)
        
        If Not IsNull(!CAPITALSOCIAL) Then
           txtCapital.Text = FormatNumber(!CAPITALSOCIAL, 2)
        Else
           txtCapital.Text = 0
        End If
        If Not IsNull(!QTDEEMPREGADO) Then
           txtNumFunc.Text = !QTDEEMPREGADO
        Else
           txtNumFunc.Text = 0
        End If
       
        If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "IMCAVAGNA" Then
            txtSenha.Text = SubNull(!SENHAISS)
        Else
            txtSenha.Text = "** Restrito **"
        End If
        
        'atividades extenso
        txtAtivExt.Text = SubNull(!ativextenso)
        txtQtde.Text = Val(SubNull(!QTDEPROF))
        If txtQtde.Text = 0 Then txtQtde.Text = 1
    
        'dispensa ie
        If Not IsNull(!DISPENSAIEDATA) Then
            bDispensadoIE = True
            chkIE.value = vbChecked
            mskDataIE.Text = Format(!DISPENSAIEDATA, "dd/mm/yyyy")
            txtNumProcIE.Text = SubNull(!DISPENSAIEPROC)
        Else
            bDispensadoIE = False
        End If
        mskDataAP.Text = Format(!DTALVARAPROVISORIO, "dd/mm/yyyy")
        
        If Val(SubNull(!codatividade)) > 0 Then
        Sql = "SELECT * FROM ATIVIDADE WHERE CODATIVIDADE=" & !codatividade
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        With RdoAux
            If Val(SubNull(!ALVARA)) = 0 Then
                chkAlvara.value = vbUnchecked
            Else
                chkAlvara.value = vbChecked
            End If
           .Close
        End With
        End If
        
        
    End If
   .Close
End With

Sql = "SELECT MOBILIARIO.CODATIVIDADE, DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Val(!codatividade) > 0 Then
            txtAtiv.Text = Format(!codatividade, "00000") & " - " & !descatividade
            
                        
            Select Case !CODIGOALIQ
                Case 1
                    txtValorAliq.Text = FormatNumber(!VALORALIQ1, 3)
                Case 2
                    txtValorAliq.Text = FormatNumber(!VALORALIQ2, 3)
                Case 3
                    txtValorAliq.Text = FormatNumber(!VALORALIQ3, 3)
            End Select
            If Not IsNull(!CODIGOALIQ) Then
               lblAliq.Caption = !CODIGOALIQ
            Else
               lblAliq.Caption = 0
            End If
            txtArea.Text = FormatNumber(!areatl, 2)
        Else
            txtAtiv.Text = 0
            txtValorAliq.Text = 0
            txtArea.Text = 0
        End If
    Else
        txtAtiv.Text = 0
        txtValorAliq.Text = 0
    '    txtArea.Text = 0
    End If
   .Close
End With

'grid iss
Sql = "SELECT MOBILIARIOATIVIDADEISS.CODMOBILIARIO,MOBILIARIOATIVIDADEISS.CODTRIBUTO,"
Sql = Sql & "MOBILIARIOATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE,MOBILIARIOATIVIDADEISS.SEQ,"
Sql = Sql & "MOBILIARIOATIVIDADEISS.QTDEISS,MOBILIARIOATIVIDADEISS.VALORISS,TRIBUTO.DESCTRIBUTO FROM MOBILIARIOATIVIDADEISS INNER JOIN "
Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE "
Sql = Sql & "Inner Join TRIBUTO ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TRIBUTO.CODTRIBUTO "
Sql = Sql & "Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & nCodigo
Sql = Sql & " ORDER BY SEQ"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !CodTributo = 11 Then
            sTipoIss = "F"
            nValorAliq = RetornaAliquotaISS(!codatividade, Format(Now, "dd/mm/yyyy"))
'            Sql = "SELECT ALIQUOTA FROM TABELAISS WHERE TIPOISS=11 AND CODIGOATIV=" & !CODATIVIDADE
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount > 0 Then
'                nValorAliq = RdoAux2!Aliquota
'            Else
'                nValorAliq = 0
'            End If
'            RdoAux2.Close
        ElseIf !CodTributo = 13 Then
            sTipoIss = "V"
            nValorAliq = RetornaAliquotaISS(!codatividade, Format(Now, "dd/mm/yyyy"))
'            Sql = "SELECT ALIQUOTA FROM TABELAISS WHERE TIPOISS=13 AND CODIGOATIV=" & !CODATIVIDADE
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            nValorAliq = RdoAux2!Aliquota
'            RdoAux2.Close
        ElseIf !CodTributo = 12 Then
            sTipoIss = "E"
            nValorAliq = !valoriss
        End If
        grdAtiv.AddItem sTipoIss & Chr(9) & Format(!codatividade, "0000") & " - " & !descatividade & Chr(9) & !QTDEISS & Chr(9) & FormatNumber(nValorAliq, 3)
       .MoveNext
    Loop
   .Close
End With

'grid VS
'Sql = "SELECT * FROM MOBILIARIOATIVIDADEVS2 WHERE CODMOBILIARIO=" & nCodigo
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        sCnae = Format(!divisao, "00") & !grupo & Left(Format(!classe, "00"), 1) & "-" & Right$(Format(!classe, "00"), 1) & "/" & Format(!subclasse, "00")
'        Sql = "SELECT * FROM CNAESUBCLASSE WHERE DIVISAO=" & !divisao & " AND GRUPO=" & !grupo & " AND CLASSE=" & !classe & " AND SUBCLASSE=" & !subclasse
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            sDesc = !descricao
'           .Close
'        End With
'        Sql = "SELECT cnaecriterio.criterio,cnaecriteriodesc.descricao,cnaecriterio.valor From cnaecriteriodesc INNER JOIN cnaecriterio ON "
'        Sql = Sql & "(cnaecriteriodesc.criterio = cnaecriterio.criterio) WHERE DIVISAO=" & !divisao & " AND GRUPO=" & !grupo & " AND CLASSE=" & !classe & " AND SUBCLASSE=" & !subclasse & " AND cnaecriterio.CRITERIO=" & !criterio
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            If .RowCount > 0 Then
'                sDesc = sDesc & " - " & !descricao
'            End If
'           .Close
'        End With
'        grdVS.AddItem sCnae & Chr(9) & !criterio & Chr(9) & sDesc & Chr(9) & !qtde & Chr(9) & Format(!Valor, "#0.0000")
'       .MoveNext
'    Loop
'   .Close
'End With

'grdVS2.Rows = 1
'Sql = "SELECT MOBILIARIOATIVIDADEVS.CODVIGSANIT,MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT,"
'Sql = Sql & "MOBILIARIOATIVIDADEVS.SEQ,MOBILIARIOATIVIDADEVS.QTDE,VIGSANITARIA.DESCVIGSANITARIA,VIGSANITARIA.VALORALIQ "
'Sql = Sql & "FROM MOBILIARIOATIVIDADEVS INNER JOIN VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT "
'Sql = Sql & "AND MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT "
'Sql = Sql & "Where MOBILIARIOATIVIDADEVS.CODMOBILIARIO = " & nCodigo
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        If !SUBCODVIGSANIT > 0 Then
'           Sql = "SELECT DESCVIGSANITARIA FROM VIGSANITARIA WHERE CODVIGSANIT=" & !CODVIGSANIT & " AND SUBCODVIGSANIT=0"
'           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'           grdVS2.AddItem Format(!CODVIGSANIT, "00") & Chr(9) & Format(!SUBCODVIGSANIT, "00") & Chr(9) & RdoAux2!DESCVIGSANITARIA & " - " & !DESCVIGSANITARIA & Chr(9) & Val(SubNull(!qtde)) & Chr(9) & FormatNumber(!valoraliq, 4)
'           RdoAux2.Close
'        Else
'           grdVS2.AddItem Format(!CODVIGSANIT, "00") & Chr(9) & Format(!SUBCODVIGSANIT, "00") & Chr(9) & !DESCVIGSANITARIA & Chr(9) & Val(SubNull(!qtde)) & Chr(9) & FormatNumber(!valoraliq, 4)
'        End If
'       .MoveNext
'    Loop
'   .Close
' End With

grdVS.Rows = 1
Sql = "SELECT mobiliariovs.codigo, mobiliariovs.cnae, mobiliariovs.criterio, mobiliariovs.qtde, cnae.descricao, cnaecriteriodesc.descricao AS desc2, cnaecriteriodesc.valor "
Sql = Sql & "FROM mobiliariovs INNER JOIN cnae ON mobiliariovs.cnae = cnae.cnae INNER JOIN cnaecriteriodesc ON mobiliariovs.criterio = cnaecriteriodesc.criterio "
Sql = Sql & "WHERE mobiliariovs.codigo = " & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdVS.AddItem Left(!Cnae, 4) & "-" & Mid(!Cnae, 5, 1) & "/" & Right(!Cnae, 2) & Chr(9) & !criterio & Chr(9) & UCase(!descricao) & Chr(9) & !QTDE & Chr(9) & Format(!Valor, "#0.0000")
       .MoveNext
    Loop
   .Close
 End With

'grid TL
Sql = "SELECT MOBILIARIOATIVIDADETL.CODIGOMOB,MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE,"
Sql = Sql & "MOBILIARIOATIVIDADETL.CODIGOALIQ,MOBILIARIOATIVIDADETL.AREA,MOBILIARIOATIVIDADETL.QTDE,VALORALIQ1,"
Sql = Sql & "VALORALIQ2,VALORALIQ3 FROM MOBILIARIOATIVIDADETL INNER JOIN ATIVIDADE ON "
Sql = Sql & "MOBILIARIOATIVIDADETL.CODATIVIDADE = ATIVIDADE.CODATIVIDADE "
Sql = Sql & "Where MOBILIARIOATIVIDADETL.CODIGOMOB = " & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Select Case !CODIGOALIQ
            Case 1
                nValorAliq = FormatNumber(!VALORALIQ1, 3)
            Case 2
                nValorAliq = FormatNumber(!VALORALIQ2, 3)
            Case 3
                nValorAliq = FormatNumber(!VALORALIQ3, 3)
        End Select
        grdTemp.AddItem Format(!codatividade, "00") & Chr(9) & Format(!descatividade, "00") & Chr(9) & nValorAliq & Chr(9) & !CODIGOALIQ & Chr(9) & FormatNumber(!Area, 2) & Chr(9) & !QTDE
       .MoveNext
    Loop
   .Close
End With

'grid historico
Sql = "SELECT mobiliariohist.codmobiliario, mobiliariohist.seq, mobiliariohist.datahist, mobiliariohist.obs, mobiliariohist.userid, usuario.nomelogin "
Sql = Sql & "FROM mobiliariohist LEFT OUTER JOIN usuario ON mobiliariohist.userid = usuario.Id "
Sql = Sql & "Where CODMOBILIARIO = " & nCodigo & " ORDER BY SEQ "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdHist.AddItem Format(!DATAHIST, "dd/mm/yyyy") & Chr(9) & Format(!Seq, "00") & Chr(9) & !obs & Chr(9) & SubNull(!NomeLogin)
       .MoveNext
    Loop
   .Close
End With

'LISTA PLACA
Sql = "SELECT PLACA FROM MOBILIARIOPLACA "
Sql = Sql & "Where CODIGO = " & nCodigo & " ORDER BY PLACA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstPlaca.AddItem !PLACA
       .MoveNext
    Loop
   .Close
End With

'grid NF
Sql = "SELECT SEQ,SERIE,NUMINI,NUMFIM,NUMAUT,DATAAUT,CANCEL,USUARIO FROM MOBILIARIONF  "
Sql = Sql & "Where CODIGOMOB = " & nCodigo & " ORDER BY SEQ "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdNF.AddItem Format(!Seq, "0000") & Chr(9) & SubNull(!Serie) & Chr(9) & Format(!numini, "00000000") & _
        Chr(9) & Format(!numfim, "00000000") & Chr(9) & IIf(IsNull(!NUMAUT), "XXXXX", IIf(!Cancel = False, !NUMAUT, !NUMAUT & " (Cancelado)")) & Chr(9) & IIf(IsNull(!DATAAUT), "", Format(!DATAAUT, "dd/mm/yyyy")) & Chr(9) & SubNull(!USUARIO)
       .MoveNext
    Loop
   .Close
End With

'grid LIVRO
Sql = "SELECT SEQ,MODELO,DATAAB,DATAEN FROM MOBILIARIOLIVRO  "
Sql = Sql & "Where CODIGOMOB = " & nCodigo & " ORDER BY SEQ "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdLivro.AddItem Format(!Seq, "0000") & Chr(9) & !MODELO & Chr(9) & Format(!DATAAB, "dd/mm/yyyy") & Chr(9) & IIf(IsNull(!DATAEN), "", Format(!DATAEN, "dd/mm/yyyy"))
       .MoveNext
    Loop
   .Close
End With
If grdLivro.Rows > 1 Then
   grdLivro.Row = 1: grdLivro_Click
End If

'PROFISSIONAL RESPONSAVEL
Sql = "SELECT CODPROFRESP,NOMEORGAO,NUMREGISTRORESP,NOMECIDADAO FROM vwCONSULTACONSELHORESP "
Sql = Sql & "WHERE CODIGOMOB=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Not IsNull(!CODPROFRESP) Then
           txtNomeProf.Text = !CODPROFRESP & " - " & SubNull(!nomecidadao)
        End If
        txtTipoConselho.Text = SubNull(!NOMEORGAO)
        txtNumRegistro.Text = SubNull(!NUMREGISTRORESP)
    End If
   .Close
End With

'endereço entrega
Sql = "SELECT TIPO,CODLOGRADOURO,NOMELOGRADOURO,NUMIMOVEL,COMPLEMENTO,UF,CODCIDADE,CODBAIRRO,"
Sql = Sql & "CEP,DESCBAIRRO,DESCCIDADE FROM MOBILIARIOENDENTREGA WHERE CODMOBILIARIO=" & nCodigo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If (!Tipo) = 1 Then
            OptEE(0).value = True
        Else
            OptEE(1).value = True
        End If
        If Val(!CodLogradouro) > 0 Then
            txtEECodLogr.Text = !CodLogradouro
            txtEENumero.Text = !NUMIMOVEL
            txtEECodLogr_LostFocus
        Else
            txtEECodLogr.Text = 0
            txtEENomeLogr.Text = !NomeLogradouro
            txtEENumero.Text = !NUMIMOVEL
            If Not IsNull(!Cep) Then
               If Len(Trim$(!Cep)) = 9 Then
                  mskEECep.Text = !Cep
               End If
            End If
        End If
        txtEECompl.Text = !Complemento
        bExec = False
        For d = 0 To cmbEEUf.ListCount - 1
            If Left$(Right$(cmbEEUf.List(d), 3), 2) = !UF Then
               cmbEEUf.ListIndex = d
               Exit For
            End If
        Next
        bExec = True
        cmbEEUF_Click
        bExec = False
        If !CodCidade > 0 Then
            For d = 0 To cmbEECidade.ListCount - 1
                If cmbEECidade.ItemData(d) = !CodCidade Then
                   cmbEECidade.ListIndex = d
                   Exit For
                End If
            Next
        Else
            cmbEECidade.ListIndex = -1
        End If
        bExec = True
        cmbEECidade_Click
        If !CodBairro <> 999 And !CodBairro <> 0 Then
            bExec = False
            For d = 0 To cmbEEBairro.ListCount - 1
                If cmbEEBairro.ItemData(d) = !CodBairro Then
                   cmbEEBairro.ListIndex = d
                   Exit For
                End If
            Next
            bExec = True
        End If
        bExec = False
        chkEnd.value = vbUnchecked
        bExec = True
    Else
        chkEnd.value = vbChecked
    End If
   .Close
End With

'dados para as fotos
Sql = "SELECT CODREDUZIDO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,"
Sql = Sql & "CODLOGR,LI_NUM FROM vwCnsImovel WHERE CODLOGR=" & Val(txtCodLogr.Text)
Sql = Sql & " AND LI_NUM=" & Val(txtNumero.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtImovel.Text = !CODREDUZIDO
       lblDist.Caption = !Distrito
       lblSetor.Caption = Format(!Setor, "00")
       lblQuadra.Caption = Format(!Quadra, "0000")
       lblLote.Caption = Format(!Lote, "00000")
       lblSeq.Caption = Format(!Seq, "000")
       lblUnidade.Caption = Format(!Unidade, "00")
       lblSubUnid.Caption = Format(!SubUnidade, "000")
    End If
   .Close
End With

'suspenção
Sql = "SELECT CODTIPOEVENTO,DATAEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodigo
Sql = Sql & " ORDER BY DATAEVENTO DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        lblSusp.Visible = False
    Else
        If !CODTIPOEVENTO = 2 Then
            lblSusp.Visible = True
            lblSusp.Caption = "(SUSPENSA EM " & Format(!DATAEVENTO, "dd/mm/yyyy") & ")"
        Else
            lblSusp.Visible = False
        End If
    End If
   .Close
End With

'cnae
Sql = "SELECT * FROM MOBILIARIOCNAE WHERE CODMOBILIARIO=" & nCodigo & " AND PRINCIPAL=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        sClasse = Format(!classe, "00")
        mskCnae.Text = Format(!divisao, "00") & !grupo & Left$(sClasse, 1) & "-" & Right$(sClasse, 1) & "/" & Format(!subclasse, "00")
    End If
   .Close
End With

Sql = "SELECT * FROM MOBILIARIOCNAE WHERE CODMOBILIARIO=" & nCodigo & " AND PRINCIPAL=0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sClasse = Format(!classe, "00")
        cmbCnae.AddItem Format(!divisao, "00") & !grupo & Left$(sClasse, 1) & "-" & Right$(sClasse, 1) & "/" & Format(!subclasse, "00")
        'cmbCnae.AddItem !secao & Format(!divisao, "00") & !grupo & Left$(sClasse, 1) & "-" & Right$(sClasse, 1) & "/" & Format(!subclasse, "00")
       .MoveNext
    Loop
    If cmbCnae.ListCount > 0 Then cmbCnae.ListIndex = 0
   .Close
End With

Sql = "SELECT ANO, NUMERO, COMPLEMENTO, DATAENTRADA, DATAARQUIVA From processogti Where INSC = " & nCodigo & " ORDER BY ANO, NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdProc.AddItem CStr(!Numero) & "-" & RetornaDVProcesso(!Numero) & "/" & !Ano & Chr(9) & Format(!DATAENTRADA, "dd/mm/yyyy") & Chr(9) & !Complemento & _
        Chr(9) & IIf(IsNull(!DATAARQUIVA), "", Format(!DATAARQUIVA, "dd/mm/yyyy"))
        
       .MoveNext
    Loop
   .Close
End With
txtCodIss.Text = nCodigo
txtNomeISS.Text = txtRazao.Text
cmbAnoISS.Text = Year(Now)

'DOCUMENTOPIC
'Sql = "select documento from documentopic where codigo=" & nCodigo
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        sTmp = !Documento
'        ReDim Preserve aPic(UBound(aPic) + 1)
'        aPic(UBound(aPic)).sTipo = Left(sTmp, 2)
'        aPic(UBound(aPic)).sAno = Mid(sTmp, 3, 4)
'        aPic(UBound(aPic)).sArq = sTmp
'        aPic(UBound(aPic)).sExt = Right(sTmp, 3)
'        For x = 1 To UBound(aTipoDoc)
'            If aTipoDoc(x).nCod = Val(Left(sTmp, 2)) Then
'                aPic(UBound(aPic)).sTipoExt = aTipoDoc(x).sNome
'                Exit For
'            End If
'        Next
'       .MoveNext
'    Loop
'   .Close
'End With
'
'If UBound(aPic) > 0 Then
'    lblPagDoc.Caption = "Documento 1 de " & UBound(aPic)
'    cmdD2.Enabled = True
'    nPointer = 1
'    LoadPic
'End If

IECheck
SNCheck
Liberado
End Sub

Private Sub Limpa()
    LimpaISS
    ReDim aPic(0)
    lblPagDoc.Caption = "Documento 0 de 0"
    frMei.Visible = False
    cmdD1.Enabled = False: cmdD2.Enabled = False
    cmbTipoDoc.ListIndex = -1
    txtArq.Text = ""
    txtImovel.Text = ""
    imgTmp.Picture = LoadPicture("")
    frISS.Visible = False
    txtCodIss.Text = ""
    txtNomeISS.Text = ""
    LimpaMascara mskDataIE
    txtNumProcIE.Text = ""
    chkHabilitar.value = vbUnchecked
    cmbPagto.ListIndex = 0
   'dados da empresa
    chkLiberadoVRE.value = vbUnchecked
    txtSIL.Text = ""
    txtCodEmpresa.Text = ""
    txtRazao.Text = ""
    txtFantasia.Text = ""
    txtInscEst.Text = ""
    txtPonto.Text = ""
    LimpaMascara mskCPF
    LimpaMascara mskCNPJ
    LimpaMascara mskDataAb
    LimpaMascara mskDataEn
    LimpaMascara mskDataPAb
    LimpaMascara mskDataAP
    LimpaMascara mskDataPEn
    LimpaMascara mskRG
    txtOrgao.Text = ""
    txtNumProcA.Text = ""
    txtNumProcE.Text = ""
    cmbHorario.ListIndex = -1
    lblSN.Caption = "NÃO"
    lblISSEletro.Caption = "NÃO"
'    chkRE.value = 0
 '   chkMei.value = 0
  '  chkEmiteNF.value = 0
    chkIE.value = 0
    'chkAlvara.value = 0
    chkIsentoISS.value = 0
    chkIsentoTaxa.value = 0
    txtHorarioExt.Text = ""
    grdProc.Rows = 1
    'localização
    cmbUF.ListIndex = -1
    cmbCidade.Clear
    cmbBairro.Clear
    txtCodLogr.Text = ""
    txtNomeLogr.Text = ""
    txtNumero.Text = ""
    LimpaMascara mskCEP
    txtHP.Text = ""
    
    'proprietário
    grdProp.Rows = 1
    txtNomeContato.Text = ""
    txtCargo.Text = ""
    txtFone.Text = ""
    txtEmail.Text = ""
    txtEmailNF.Text = ""
    txtFax.Text = ""
    txtDDDNF.Text = ""
    txtFoneNF.Text = ""
    
    'endereço entrega
    cmbEEUf.ListIndex = -1
    cmbEEBairro.Clear
    cmbCidade.Clear
    txtEECodLogr.Text = ""
    txtEENomeLogr.Text = ""
    txtEENumero.Text = ""
    txtEECompl.Text = ""
    LimpaMascara mskEECep
        
    'Atividades
    txtAtiv.Text = ""
    txtValorAliq.Text = 0
    txtArea.Text = 0
    txtAtivExt.Text = ""
    lblTipoISS.Caption = "-->"
    grdAtiv.Rows = 1
    grdVS.Rows = 1
    LimpaMascara mskCnae
    cmbCnae.Clear
    
    'Outros
    txtSenha.Text = ""
    txtNomeProf.Text = ""
    txtTipoConselho.Text = ""
    txtNumRegistro.Text = ""
    txtCodEsc.Text = ""
    cmbNomeEsc.ListIndex = -1
    txtNumFunc.Text = 0
    txtCapital.Text = 0
    txtFoneCont.Text = ""
    txtEmailCont = ""
    LimpaMascara mskPlaca
    lstPlaca.Clear
    
    'historico
    grdHist.Rows = 1
    txtHist.Text = ""
    
    'informativo
    grdNF.Rows = 1
    grdLivro.Rows = 1
    'inscrição
    lblDist.Caption = "0"
    lblSetor.Caption = "00"
    lblQuadra.Caption = "0000"
    lblLote.Caption = "00000"
    lblSeq.Caption = "000"
    lblUnidade.Caption = "00"
    lblSubUnid.Caption = "000"
    
    grdTemp.Rows = 1
    
End Sub

Private Sub txtCodEmpresa_LostFocus()
If Val(txtCodEmpresa.Text) > 0 Then
    txtCodEmpresa = Left$(txtCodEmpresa.Text, Len(txtCodEmpresa.Text) - 1) & "-" & Right$(txtCodEmpresa.Text, 1)
End If

End Sub

Private Sub txtCodEsc_KeyPress(KeyAscii As Integer)
Tweak txtCodEsc, KeyAscii, IntegerPositive

End Sub

Private Sub txtCodIss_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    txtCodIss_LostFocus
Else
    Tweak txtCodIss, KeyAscii, IntegerPositive
End If
End Sub

Private Sub txtCodIss_LostFocus()
Dim Sql As String, RdoAux As rdoResultset

txtNomeISS.Text = ""
If Val(txtCodIss.Text) = 0 Then Exit Sub

If Val(txtCodIss.Text) >= 100000 And Val(txtCodIss.Text) < 500000 Then
    Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & Val(txtCodIss.Text)
ElseIf Val(txtCodIss.Text) >= 500000 And Val(txtCodIss.Text) < 800000 Then
    Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCodIss.Text)
Else
    MsgBox "Código inválido!", vbCritical, "Atenção"
    LimpaISS
    Exit Sub
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtNomeISS.Text = !Nome
        If chkHabilitar.value = vbChecked Then
            CarregaISS
        End If
    Else
        MsgBox "Código inválido!", vbCritical, "Atenção"
        LimpaISS
        .Close
        Exit Sub
    End If
   .Close
End With


End Sub
    
Private Sub txtCodLogr_Change()
If Val(txtCodLogr.Text) = 0 And txtCompl.BackColor = Branco Then
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
Tweak txtCodLogr, KeyAscii, IntegerPositive

End Sub

Private Sub txtCodLogr_LostFocus()
If Not bExec Then Exit Sub
If Val(txtCodLogr.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogr.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(SubNull(!AbrevTitLog))) & " " & !NomeLogradouro
       Else
          MsgBox "Código do Logradouro Inválido.", vbExclamation, "Atenção"
          txtCodLogr.Text = ""
          txtNomeLogr.Text = ""
       End If
   End With
End If

If Val(txtNumero.Text) > 10000 Then
    MsgBox "Nº inválido.", vbExclamation, "Atenção"
    txtNumero.SetFocus
    Exit Sub
End If
If cmbBairro.Text <> "ZONA RURAL" Then
    LimpaMascara mskCEP
    If Val(txtCodLogr.Text) > 0 Then
         mskCEP.Text = RetornaCEP(Val(txtCodLogr.Text), Val(txtNumero.Text))
    Else
        LimpaMascara mskCEP
    End If
End If

End Sub

Private Sub txtDDDNF_KeyPress(KeyAscii As Integer)
Tweak txtDDDNF, KeyAscii, IntegerPositive
End Sub

Private Sub txtEECodLogr_Change()
If Val(txtEECodLogr.Text) = 0 And txtEECompl.BackColor = Branco Then
   txtEENomeLogr.Enabled = True
   txtEENomeLogr.BackColor = Branco
   txtEENomeLogr.Text = ""
Else
   txtEENomeLogr.Text = ""
   txtEENomeLogr.Enabled = False
   txtEENomeLogr.BackColor = Kde
End If

End Sub

Private Sub txtEECodLogr_KeyPress(KeyAscii As Integer)
Tweak txtEECodLogr, KeyAscii, IntegerPositive
End Sub

Private Sub txtEECodLogr_LostFocus()
If Val(txtEECodLogr.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtEECodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtEENomeLogr.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          MsgBox "Codigo logradouro inválido.", vbExclamation, "Atenção"
          txtEECodLogr.Text = ""
          txtEENomeLogr.Text = ""
       End If
   End With
End If

LimpaMascara mskEECep
If Val(txtEENumero.Text) > 10000 Then
    MsgBox "Nº inválido.", vbExclamation, "Atenção"
    txtEENumero.SetFocus
    Exit Sub
End If
If Val(txtEECodLogr.Text) > 0 Then
     mskEECep.Text = RetornaCEP(Val(txtEECodLogr.Text), Val(txtEENumero.Text))
Else
    If txtEENomeLogr.Text = "" Then
        LimpaMascara mskEECep
    End If
End If

End Sub

Private Sub txtEENomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstEENomeLog.Clear
   If txtEENomeLogr.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtEENomeLogr) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstEENomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstEENomeLog.ItemData(lstEENomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstEENomeLog.Visible = True
             lstEENomeLog.ListIndex = 0
             lstEENomeLog.ZOrder 0
             lstEENomeLog.SetFocus
          Else
             MsgBox "Digite o nome do logradouro a ser pesquisado, sem especificar o tipo e o título.", vbInformation, "Atenção"
             lstEENomeLog.Visible = False
             txtEENomeLogr.SetFocus
          End If
      End With
   End If
ElseIf KeyAscii = vbKeyEscape Then
        lstEENomeLog.Visible = False
        txtEENomeLogr.SetFocus
Else
   txtEECodLogr.Text = 0
End If

End Sub

Private Sub txtEENumero_KeyPress(KeyAscii As Integer)
Tweak txtEENumero, KeyAscii, IntegerPositive
End Sub

Private Sub txtFoneNF_KeyPress(KeyAscii As Integer)
Tweak txtFoneNF, KeyAscii, IntegerPositive
End Sub

Private Sub txtImovel_KeyPress(KeyAscii As Integer)
Tweak txtImovel, KeyAscii, IntegerPositive
End Sub

Private Sub txtInscEst_KeyPress(KeyAscii As Integer)
Tweak txtInscEst, KeyAscii, IntegerPositive
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
             lstNomeLog.ZOrder 0
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

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
Tweak txtNumero, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumero_LostFocus()

If Val(txtNumero.Text) > 10000 Then
    MsgBox "Nº inválido.", vbExclamation, "Atenção"
    txtNumero.SetFocus
    Exit Sub
End If
If cmbBairro.ListIndex = -1 Then Exit Sub
If cmbBairro.ItemData(cmbBairro.ListIndex) = 96 Then Exit Sub
LimpaMascara mskCEP

If Val(txtCodLogr.Text) > 0 Then
     mskCEP.Text = RetornaCEP(Val(txtCodLogr.Text), Val(txtNumero.Text))
    'dados para as fotos
    Sql = "SELECT codreduzido,DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,"
    Sql = Sql & "CODLOGR,LI_NUM FROM vwCnsImovel WHERE CODLOGR=" & Val(txtCodLogr.Text)
    Sql = Sql & " AND LI_NUM=" & Val(txtNumero.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
           txtImovel.Text = !CODREDUZIDO
           lblDist.Caption = !Distrito
           lblSetor.Caption = Format(!Setor, "00")
           lblQuadra.Caption = Format(!Quadra, "0000")
           lblLote.Caption = Format(!Lote, "00000")
           lblSeq.Caption = Format(!Seq, "000")
           lblUnidade.Caption = Format(!Unidade, "00")
           lblSubUnid.Caption = Format(!SubUnidade, "000")
        Else
            txtImovel.Text = "0"
           lblDist.Caption = "0"
           lblSetor.Caption = "00"
           lblQuadra.Caption = "0000"
           lblLote.Caption = "00000"
           lblSeq.Caption = "000"
           lblUnidade.Caption = "00"
           lblSubUnid.Caption = "000"
        End If
       .Close
    End With
Else
    LimpaMascara mskCEP
End If

End Sub

Private Sub txtNumFunc_KeyPress(KeyAscii As Integer)
Tweak txtNumFunc, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumProcA_LostFocus()

Dim sValidaProc As String
If Trim(txtNumProcA.Text) = "" Then Exit Sub
sValidaProc = ValidaProcesso(txtNumProcA.Text)
If sValidaProc <> "OK" Then
    MsgBox sValidaProc, vbCritical, "Atenção"
    LimpaMascara mskDataPAb
    'txtNumProcA.SetFocus
Else
    mskDataPAb.Text = Format(RetornaDataProcesso(Val(Left$(txtNumProcA.Text, Len(txtNumProcA.Text) - 5)), Val(Right$(txtNumProcA.Text, 4))), "dd/mm/yyyy")
End If

End Sub

Private Sub txtNumProcE_LostFocus()
Dim sValidaProc As String
If Trim(txtNumProcE.Text) = "" Then Exit Sub
If Len(txtNumProcE.Text) < 5 Then Exit Sub
sValidaProc = ValidaProcesso(txtNumProcE.Text)
If sValidaProc <> "OK" Then
    MsgBox sValidaProc, vbCritical, "Atenção"
End If
mskDataPEn.Text = Format(RetornaDataProcesso(Val(Left$(txtNumProcE.Text, Len(txtNumProcE.Text) - 5)), Val(Right$(txtNumProcE.Text, 4))), "dd/mm/yyyy")

End Sub

Private Sub txtNumProcIE_LostFocus()
Dim sValidaProc As String
If Trim(txtNumProcIE.Text) = "" Then Exit Sub
sValidaProc = ValidaProcesso2(txtNumProcIE.Text)
If sValidaProc <> "OK" And Left(sValidaProc, 4) <> "Este" Then
    MsgBox sValidaProc, vbCritical, "Atenção"
    txtNumProcIE.Text = ""
End If
End Sub

Private Sub txtQtdeCnae_KeyPress(KeyAscii As Integer)
Tweak txtQtdeCnae, KeyAscii, IntegerPositive
End Sub

Private Sub txtValorAliq_KeyPress(KeyAscii As Integer)
Tweak txtValorAliq, KeyAscii, DecimalPositive
End Sub

Private Sub SNCheck()
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNASN(" & Format(Val(txtCodEmpresa.Text), "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
'ConectaEicon
'Sql = "select * from  tb_inter_empr_snacional_giss Where NUM_CADASTRO=" & Val(txtCodEmpresa.Text) & " order by timestamp desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
     If RdoAux!RETORNO = 1 Then
'    If RdoAux.RowCount > 0 Then
'        If IsNull(!data_fim) Then
            lblSN.Caption = "SIM"
            lblSN.ForeColor = "&H00008000"
'        Else
'            lblSN.Caption = "NÃO"
'            lblSN.ForeColor = "&H000000C0"
 '       End If
     Else
        lblSN.Caption = "NÃO"
        lblSN.ForeColor = "&H000000C0"
     End If
    .Close
End With
'cnEicon.Close
End Sub

Private Sub IECheck()
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNAIE(" & Format(Val(txtCodEmpresa.Text), "000000") & ",'" & Format(Now, sDataFormat) & "') AS RETORNO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
     If RdoAux!RETORNO = 1 Then
        lblISSEletro.Caption = "SIM"
        lblISSEletro.ForeColor = "&H00008000"
     Else
        lblISSEletro.Caption = "NÃO"
        lblISSEletro.ForeColor = "&H000000C0"
     End If
    .Close
End With

End Sub


Private Sub MontaMenuISS()

   Set m_cMenuISS = New cPopupMenu
   With m_cMenuISS
      .hwndOwner = Me.HWND
      .HeaderStyle = ecnmHeaderSeparator
      .GradientHighlight = True
      
      i = .AddItem("ISS Fixo", "", 1, , , , , "mnuISSF")
      .OwnerDraw(i) = True
      i = .AddItem("ISS Estimado", "", 1, , , , , "mnuISSE")
      .OwnerDraw(i) = True
      i = .AddItem("ISS Variável", "", 1, , , , , "mnuISSV")
      .OwnerDraw(i) = True
      
   End With
   
End Sub

Private Sub AtivaTela(nTela As Integer)
Dim t As Integer
frIE.Visible = False
frIE.Width = 4965
frIE.Left = 6345

For t = 0 To 11
    If t = 10 Then GoTo proximo
    If t <> nTela Then
        Tela(t).Visible = False
        btMenu(t).BackColor = &H400000
    Else
        btMenu(t).BackColor = &HFF0000
        Tela(t).Visible = True
        Tela(t).Left = 2040
        Tela(t).Top = 60
    End If
proximo:
Next
If nTela = 9 Then
    lblIss_Click (1)
    frIE.Visible = True
    frIE.ZOrder 0
End If

End Sub

Private Sub LimpaISS()
Dim j As Integer

For j = 1 To 24
    lblIss(j).Caption = "0,00"
    lblIss(j).ForeColor = vbBlack
Next
lblTotISSE.Caption = "0,00"
lblTotISSR.Caption = "0,00"
grdMain.Clear

End Sub

Private Sub CarregaISS()
Dim Sql As String, RdoAux As rdoResultset, nPos As Integer, aISS() As ISSELETRO, RdoAux2 As rdoResultset, nNumdoc As Long, nValorPago As Double
Dim x As Integer, bAchou As Boolean, nTotalE As Double, nTotalR As Double, sCNPJTomador As String
Dim nTotalRows As Long, nPosRows As Long, sCNPJPrestador As String, aNotaNum() As Long

If txtNomeISS.Text = "" Then Exit Sub


PBar.value = 0
Ocupado
On Error Resume Next
ReDim aNota(0): nPos = 1: ReDim aISS(0): nTotalE = 0: nTotalR = 0: ReDim aNotaNum(0)

Sql = "SELECT DISTINCT nfisseletro2.identificaprestador, nfisseletro2.tipoprestador, nfisseletro2.tiponota, nfisseletro2.numeronota, nfisseletro2.serie, nfisseletro2.dataemissao,"
Sql = Sql & "nfisseletro2.mesref, nfisseletro2.anoref, nfisseletro2.statusnota, nfisseletro2.datacancel, nfisseletro2.natureza, nfisseletro2.valortotal, nfisseletro2.valorservico,"
Sql = Sql & "nfisseletro2.valorimposto, nfisseletro2.recolhimento, nfisseletro2.atividade, nfisseletro2.aliquota, nfisseletro2.razaoprestador, nfisseletro2.cidadeprestador,"
Sql = Sql & "nfisseletro2.ufprestador, nfisseletro2.localprestador, nfisseletro2.identificatomador, nfisseletro2.tipotomador, nfisseletro2.razaotomador, nfisseletro2.cidadetomador,"
Sql = Sql & "nfisseletro2.UFTomador , nfisseletro2.LocalTomador, nfisseletro2.NumDoc, nfisseletro2.simplesnac from nfisseletro2 "
Sql = Sql & "WHERE ((IDENTIFICAPRESTADOR='" & txtCodIss.Text & "' OR IDENTIFICAPRESTADOR='" & Format(txtCodIss.Text, "00000000000000") & "') or (IDENTIFICATOMADOR='" & txtCodIss.Text & "' OR IDENTIFICATOMADOR='" & Format(txtCodIss.Text, "00000000000000") & "'))"
Sql = Sql & "and AnoRef = " & Val(cmbAnoISS.Text)

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotalRows = RdoAux.RowCount
    nPos = 1
    Do Until .EOF
        If nPosRows Mod 20 = 0 Then
            CallPb nPosRows, nTotalRows
        End If
'        If !NumDoc = 3031726 Then MsgBox "teste"
        'If !Recolhimento <> 2 And !TipoNota <> 1 Then GoTo PROXIMO
        If Val(!IdentificaTomador) = Val(txtCodIss.Text) Then
            If !Recolhimento <> 2 And !Recolhimento <> 4 Then
            'If !Recolhimento <> 4 Then
                GoTo proximo
            Else
                If !TipoNota = 1 Then
                    GoTo proximo
                End If
            End If
        Else
            If !TipoNota = 2 Then
                GoTo proximo
            End If
'            GoTo PROXIMO
        End If
        
'        If !NumDocumento = 3033354 Then MsgBox "teste"
        If Val(!IdentificaPrestador) <> Val(txtCodEmpresa.Text) Then
            If !TipoNota = 1 Then
                GoTo proximo
            End If
        End If
        
        Sql = "select codigomob,cnpj from mobiliario where codigomob=" & Val(!IdentificaTomador)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sCNPJTomador = Format(SubNull(RdoAux2!Cnpj), "0#\.###\.###/####-##")
        Else
            sCNPJTomador = ""
        End If
        RdoAux2.Close

        Sql = "select codigomob,cnpj from mobiliario where codigomob=" & Val(!IdentificaPrestador)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            sCNPJPrestador = Format(SubNull(RdoAux2!Cnpj), "0#\.###\.###/####-##")
        Else
            sCNPJPrestador = ""
        End If
        RdoAux2.Close
        
        nNumdoc = Val(SubNull(!NumDoc))
        If nNumdoc > 0 Then
            Sql = "select valorpagoreal from debitopago where numdocumento=" & nNumdoc
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nValorPago = RdoAux2!valorpagoreal
            RdoAux2.Close
        End If
        
        
        If optRec(0).value = True Then
            If !TipoNota = 2 Then GoTo proximo
        Else
            If !TipoNota = 1 Then GoTo proximo
        End If
       ' If !NumeroNota = 10000000000015# Then MsgBox "teste"
        
        If isInLongArray(aNotaNum, !NumeroNota) Then
           For x = nIndexFind To UBound(aNota)
               If aNota(x).NumeroNota = !NumeroNota And aNota(x).Serie = !Serie Then
                   GoTo proximo
               End If
            Next
        End If
        
        
        
        
        
        
'        bAchou = False
'        For x = 1 To UBound(aNota)
'            If aNota(x).NumeroNota = !NumeroNota And aNota(x).Serie = !Serie Then
'                bAchou = True
'                Exit For
'            End If
 '       Next
 '       If bAchou Then GoTo PROXIMO
        
        ReDim Preserve aNotaNum(nPos)
        aNotaNum(nPos) = !NumeroNota
        
        ReDim Preserve aNota(nPos)
        aNota(nPos).NumeroNota = !NumeroNota
        aNota(nPos).Serie = !Serie
        aNota(nPos).TipoNota = !TipoNota
        aNota(nPos).AnoRef = !AnoRef
        aNota(nPos).MesRef = !MesRef
        aNota(nPos).DataEmissao = Format(!DataEmissao, "dd/mm/yyyy")
        aNota(nPos).StatusNota = !StatusNota
        aNota(nPos).Natureza = !Natureza
        aNota(nPos).DataCancel = Format(!DataCancel, "dd/mm/yyyy")
        aNota(nPos).ValorTotal = !ValorTotal
        aNota(nPos).ValorServico = !ValorServico
        aNota(nPos).ValorImposto = !ValorImposto
        aNota(nPos).Recolhimento = !Recolhimento
        aNota(nPos).Atividade = !Atividade
        aNota(nPos).Aliquota = !Aliquota
        aNota(nPos).RazaoPrestador = !RazaoPrestador
        aNota(nPos).CNPJPrestador = sCNPJPrestador
        aNota(nPos).IdentificaTomador = !IdentificaTomador
        aNota(nPos).TipoTomador = !TipoTomador
        aNota(nPos).RazaoTomador = !RazaoTomador
        aNota(nPos).CidadeTomador = !CidadeTomador
        aNota(nPos).UFTomador = !UFTomador
        aNota(nPos).CNPJTomador = sCNPJTomador
        aNota(nPos).NumGuia = SubNull(!NumDoc)
        aNota(nPos).Pago = IIf(aNota(nPos).NumGuia > 0 And nValorPago > 0, "Sim", "Não")
        'aNota(nPos).Pago = IIf(aNota(nPos).NumGuia > 0 And !valorpagoreal > 0, "Sim", "Não")
        bAchou = False
        For x = 1 To UBound(aISS)
            If aISS(x).nAno = aNota(nPos).AnoRef And aISS(x).nMes = aNota(nPos).MesRef Then
                bAchou = True
                Exit For
            End If
        Next
        If bAchou Then
            If aNota(nPos).TipoNota = 1 Then 'emitida
                aISS(x).nValorEmitida = aISS(x).nValorEmitida + aNota(nPos).ValorTotal
            Else
                aISS(x).nValorRecebida = aISS(x).nValorRecebida + aNota(nPos).ValorTotal
            End If
        Else
            ReDim Preserve aISS(UBound(aISS) + 1)
            aISS(UBound(aISS)).nAno = aNota(nPos).AnoRef
            aISS(UBound(aISS)).nMes = aNota(nPos).MesRef
            If aNota(nPos).TipoNota = 1 Then 'emitida
                aISS(UBound(aISS)).nValorEmitida = aNota(nPos).ValorTotal
            Else
                aISS(UBound(aISS)).nValorRecebida = aNota(nPos).ValorTotal
            End If
        End If
        nPos = nPos + 1
        nPosRows = nPosRows + 1
        DoEvents
proximo:
       .MoveNext
    Loop
   .Close
End With


For nPos = 1 To UBound(aISS)
    aISS(nPos).bSemMovimento = False
Next

'Sql = "SELECT * FROM NFISSELETROSMOV WHERE CODIGO=" & Val(Left(txtCodEmpresa.Text, 7)) & " AND "
Sql = "SELECT * FROM NFISSELETROSMOV WHERE CODIGO=" & Val(txtCodIss.Text) & " AND "
Sql = Sql & "ANO=" & Val(cmbAnoISS.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bAchou = False
        For nPos = 1 To UBound(aISS)
            If aISS(nPos).nAno = !Ano And aISS(nPos).nMes = !Mes Then
                aISS(nPos).bSemMovimento = True
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aISS(UBound(aISS) + 1)
            aISS(UBound(aISS)).nAno = !Ano
            aISS(UBound(aISS)).nMes = !Mes
            aISS(UBound(aISS)).bSemMovimento = True
        End If
       .MoveNext
    Loop
   .Close
End With
Liberado

For x = 1 To 12
    lblIss(x).Caption = FormatNumber(0, 2)
    lblTotISSE.Caption = FormatNumber(0, 2)
    lblTotISSR.Caption = FormatNumber(0, 2)
Next

If UBound(aISS) = 0 Then
    GoTo fim
End If

For x = 1 To 12
    For nPos = 1 To UBound(aISS)
        If aISS(nPos).nAno = cmbAnoISS.Text And aISS(nPos).nMes = x Then
            lblIss(x).Caption = FormatNumber(aISS(nPos).nValorEmitida, 2)
            lblIss(x + 12).Caption = FormatNumber(aISS(nPos).nValorRecebida, 2)
            lblIss(x).ForeColor = vbBlack
            If aISS(nPos).nValorEmitida = 0 Then
                If aISS(nPos).bSemMovimento = True Then
                    lblIss(x).Caption = "SEM MOV."
                End If
            End If
            nTotalE = nTotalE + aISS(nPos).nValorEmitida
            nTotalR = nTotalR + aISS(nPos).nValorRecebida
        End If
    Next
Next

For x = 1 To 12
    If lblIss(x).Caption = "0,00" Then
        lblIss(x).Caption = "SEM DECL."
        lblIss(x).ForeColor = vbRed
    End If
Next

lblTotISSE.Caption = FormatNumber(nTotalE, 2)
lblTotISSR.Caption = FormatNumber(nTotalR, 2)

fim:
PBar.value = 0
CallPb 0, 10
CarregaLista
End Sub

Private Sub txtValorNF_GotFocus()
txtValorNF.SelStart = 0
txtValorNF.SelLength = Len(txtValorNF.Text)
End Sub

Private Sub txtValorNF_KeyPress(KeyAscii As Integer)
Tweak txtValorNF, KeyAscii, DecimalPositive
End Sub

Private Sub GridHeader()

With grdMain
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "NumNota", "N° Nota", ecgHdrTextALignRight, , 50
    .AddColumn "Serie", "Série", ecgHdrTextALignCentre, , 40
    .AddColumn "TipoNota", "Tipo", ecgHdrTextALignLeft, , 50
    .AddColumn "DtEmissao", "Dt.Emissão", ecgHdrTextALignCentre, , 70
    .AddColumn "Situação", "Situação", ecgHdrTextALignLeft, , 60
    .AddColumn "DtCancel", "Dt.Cancel.", ecgHdrTextALignCentre, , 70
    .AddColumn "Natureza", "Natureza", ecgHdrTextALignLeft, , 60
    .AddColumn "VlTotal", "Vl.Total", ecgHdrTextALignRight, , 70
    .AddColumn "VlServico", "Vl.Serviço", ecgHdrTextALignRight, , 70
    .AddColumn "VlImposto", "Vl.Imposto", ecgHdrTextALignRight, , 70
    .AddColumn "Recolh", "Recolhim.", ecgHdrTextALignLeft, , 70
    .AddColumn "Atividade", "Atividade", ecgHdrTextALignLeft, , 60
    .AddColumn "Aliq", "Aliq", ecgHdrTextALignRight, , 40
    .AddColumn "RazaoPrestador", "Razão Prestador", ecgHdrTextALignLeft, , 130
    .AddColumn "CNPJPrestador", "CNPJ Prestador", ecgHdrTextALignLeft, , 110
    .AddColumn "IdTomador", "Id.Tomador", ecgHdrTextALignLeft, , 110
    .AddColumn "TipoTomador", "Tipo", ecgHdrTextALignLeft, , 40
    .AddColumn "RazaoTomador", "Razão Tomador", ecgHdrTextALignLeft, , 130
    .AddColumn "CNPJTomador", "CNPJ Tomador", ecgHdrTextALignLeft, , 110
    .AddColumn "CidadeTomador", "Cidade/UF", ecgHdrTextALignLeft, , 130
    .AddColumn "NumeroGuia", "Nº Guia", ecgHdrTextALignCentre, , 90
    .AddColumn "Pagto", "Pagto", ecgHdrTextALignCentre, , 40
    
End With

End Sub

Private Sub CarregaLista()
Dim x As Long, Sql As String, RdoAux As rdoResultset

Ocupado
grdMain.Redraw = False
grdMain.Clear
grdMain.Redraw = True
grdMain.Redraw = False

For x = 1 To UBound(aNota)
    With aNota(x)
        'If aNota(x).NumeroNota = 10000000000015# Then MsgBox "teste"
        If .AnoRef = Val(cmbAnoISS.Text) And .MesRef = Val(lblMesNF.Caption) Then
            If cmbPagto.ListIndex = 1 And .Pago = "Não" Then GoTo proximo
            If cmbPagto.ListIndex = 2 And .Pago = "Sim" Then GoTo proximo
            
            grdMain.AddRow
            grdMain.CellDetails grdMain.Rows, 1, .NumeroNota, DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 2, .Serie, DT_CENTER
            If .TipoNota = 1 Then
                grdMain.CellDetails grdMain.Rows, 3, "Emitida", DT_LEFT
            Else
                grdMain.CellDetails grdMain.Rows, 3, "Recebida", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 4, .DataEmissao, DT_CENTER
            If .StatusNota = 0 Then
                grdMain.CellDetails grdMain.Rows, 5, "Recebida", DT_LEFT
            ElseIf .StatusNota = 1 Then
                grdMain.CellDetails grdMain.Rows, 5, "Normal", DT_LEFT
            ElseIf .StatusNota = 2 Then
                grdMain.CellDetails grdMain.Rows, 5, "Cancelada", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 6, .DataCancel, DT_CENTER
            If .Natureza = 1 Then
                grdMain.CellDetails grdMain.Rows, 7, "Serviço", DT_CENTER
            Else
                grdMain.CellDetails grdMain.Rows, 7, "Mista", DT_CENTER
            End If
            
            grdMain.CellDetails grdMain.Rows, 8, FormatNumber(.ValorTotal, 2), DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 9, FormatNumber(.ValorServico, 2), DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 10, FormatNumber(.ValorImposto, 2), DT_RIGHT
            If .TipoNota = 1 And .Recolhimento = 0 Then
                grdMain.CellDetails grdMain.Rows, 11, "Isento", DT_LEFT
            ElseIf .TipoNota = 1 And .Recolhimento = 1 Then
                grdMain.CellDetails grdMain.Rows, 11, "Retido", DT_LEFT
            ElseIf .TipoNota = 1 And .Recolhimento = 2 Then
                grdMain.CellDetails grdMain.Rows, 11, "A Recolher", DT_LEFT
            ElseIf .TipoNota = 1 And .Recolhimento = 3 Then
                grdMain.CellDetails grdMain.Rows, 11, "Simples", DT_LEFT
            ElseIf .TipoNota = 2 And .Recolhimento = 1 Then
                grdMain.CellDetails grdMain.Rows, 11, "Disp.Ret.", DT_LEFT
            ElseIf .TipoNota = 2 And .Recolhimento = 2 Then
                grdMain.CellDetails grdMain.Rows, 11, "Ret.Sub.Trib.", DT_LEFT
            ElseIf .TipoNota = 2 And .Recolhimento = 3 Then
                grdMain.CellDetails grdMain.Rows, 11, "Ret.Res.Trib.", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 12, .Atividade, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 13, FormatNumber(.Aliquota, 2) & "%", DT_RIGHT
            grdMain.CellDetails grdMain.Rows, 14, .RazaoPrestador, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 15, .CNPJPrestador, DT_LEFT
            If .TipoTomador = 0 Then
                grdMain.CellDetails grdMain.Rows, 16, Format(.IdentificaTomador, "0#\.###\.###/####-##"), DT_LEFT
                grdMain.CellDetails grdMain.Rows, 17, "CNPJ", DT_LEFT
            ElseIf .TipoTomador = 1 Then
                grdMain.CellDetails grdMain.Rows, 16, Format(.IdentificaTomador, "00#\.###\.###-##"), DT_LEFT
                grdMain.CellDetails grdMain.Rows, 17, "CPF", DT_LEFT
            ElseIf .TipoTomador = 2 Then
                grdMain.CellDetails grdMain.Rows, 16, .IdentificaTomador, DT_LEFT
                grdMain.CellDetails grdMain.Rows, 17, "IM", DT_LEFT
            End If
            grdMain.CellDetails grdMain.Rows, 18, .RazaoTomador, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 19, .CNPJTomador, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 20, .CidadeTomador & "/" & .UFTomador, DT_LEFT
            grdMain.CellDetails grdMain.Rows, 21, .NumGuia, DT_CENTER
            grdMain.CellDetails grdMain.Rows, 22, .Pago, DT_CENTER
            
        End If
    End With
proximo:
Next
Liberado
grdMain.Redraw = True

End Sub

Private Sub VALORADICIONAL()
Dim z As Variant, sNomeArq As String, x As Integer, sMes As String, aValor() As VALORADICIONAL
Dim nCFOP As Integer, ax As String, nValorEntrada As Double, nValorSaida As Double

If Val(txtCodEmpresa.Text) = 0 Then
    MsgBox "Selecione uma empresa", vbExclamation, "Atenção"
    Exit Sub
End If
z = InputBox("Digite o ano de referência.", "Informação requerida", Year(Now) - 1)
If Val(z) = 0 Then
    Exit Sub
ElseIf Val(z) < 2000 Or Val(z) > Year(Now) + 5 Then
    MsgBox "Ano inválido", vbCritical, "Atenção"
    Exit Sub
End If

ReDim aValor(12)
Sql = "SELECT DISTINCT codreduzido, ref, cfop, basecalculo, isentasntrib, outras"
Sql = Sql & " FROM GIADETALHE WHERE CODREDUZIDO=" & Val(txtCodEmpresa.Text) & " AND YEAR(REF)=" & Val(z)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCFOP = RetornaNumero(!CFOP)
        If nCFOP < 5000 Then 'ENTRADA
            aValor(Month(!REF)).nBaseCalculoE = aValor(Month(!REF)).nBaseCalculoE + !BASECALCULO
            aValor(Month(!REF)).nValorNTribE = aValor(Month(!REF)).nValorNTribE + !ISENTASNTRIB
            aValor(Month(!REF)).nOutrasE = aValor(Month(!REF)).nOutrasE + !OUTRAS
        Else 'SAIDA
            aValor(Month(!REF)).nBaseCalculoS = aValor(Month(!REF)).nBaseCalculoS + !BASECALCULO
            aValor(Month(!REF)).nValorNTribS = aValor(Month(!REF)).nValorNTribS + !ISENTASNTRIB
            aValor(Month(!REF)).nOutrasS = aValor(Month(!REF)).nOutrasS + !OUTRAS
        End If
       .MoveNext
    Loop
   .Close
End With

sNomeArq = sPathBin & "\VALORADIC.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

Print #FF1, "******************************************************************************"
Print #FF1, "VALOR ADICIONAL - ANO DE REFERÊNCIA: " & z
Print #FF1, "RAZÃO SOCIAL: " & txtRazao.Text
Print #FF1, "INSCRIÇÃO: " & txtCodEmpresa.Text & " CNPJ: " & mskCNPJ.Text & " IE: " & txtInscEst.Text
Print #FF1, "******************************************************************************"
Print #FF1, ""
ax = FillSpace("MÊS", 15) & FillSpace("   VALOR SAÍDA", 15) & FillLeft("VALOR ENTRADA", 15) & FillLeft(" VL.ADICIONAL", 15)
Print #FF1, ax
Print #FF1, "****************************************************************************"

For x = 1 To 12
    Select Case x
        Case 1
            sMes = "Janeiro"
        Case 2
            sMes = "Fevereiro"
        Case 3
            sMes = "Março"
        Case 4
            sMes = "Abril"
        Case 5
            sMes = "Maio"
        Case 6
            sMes = "Junho"
        Case 7
            sMes = "Julho"
        Case 8
            sMes = "Agosto"
        Case 9
            sMes = "Setembro"
        Case 10
            sMes = "Outubro"
        Case 11
            sMes = "Novembro"
        Case 12
            sMes = "Dezembro"
    End Select
    nValorEntrada = aValor(x).nBaseCalculoE + aValor(x).nValorNTribE + aValor(x).nOutrasE
    nValorSaida = aValor(x).nBaseCalculoS + aValor(x).nValorNTribS + aValor(x).nOutrasS
    ax = FillSpace(sMes & "/" & CStr(Val(z)), 15) & FillLeft(FormatNumber(nValorSaida, 2), 15)
    ax = ax & FillLeft(FormatNumber(nValorEntrada, 2), 15) & FillLeft(FormatNumber(nValorSaida - nValorEntrada, 2), 15)
    Print #FF1, ax
    
    
Next

Print #FF1, ""
Print #FF1, ""
ax = FillSpace("CFOP", 7) & FillSpace("   VALOR", 15)
Print #FF1, ax
Print #FF1, "********************************************************"
Sql = "SELECT CFOP,SUM(BASECALCULO) as soma1,SUM(ISENTASNTRIB) as soma2,SUM(OUTRAS) as soma3 FROM GIAdetalhe WHERE "
Sql = Sql & "CODREDUZIDO=" & Val(txtCodEmpresa.Text) & " AND YEAR(REF)=" & Val(z) & " GROUP BY CFOP ORDER BY CFOP"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ax = FillSpace(!CFOP, 7) & FillLeft(FormatNumber(!SOMA1 + !SOMA2 + !SOMA3, 2), 15)
        Print #FF1, ax
        .MoveNext
    Loop
   .Close
End With
Print #FF1, ""
Print #FF1, "PMJ - VALORADIC.TXT - GERADO PELO SISTEMA GTI EM " & Format(Now, "dd/mm/yyyy")

Close #FF1
ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

Liberado

End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub LoadPic()
Dim sAno As String, sExt As String, sArq As String, sPathOrigem As String, sTipo As String
If Val(txtCodEmpresa.Text) = 0 Then Exit Sub
lblPagDoc.Caption = "Documento " & nPointer & " de " & UBound(aPic)
lblTipoDoc.Caption = aPic(nPointer).sTipoExt & "/" & aPic(nPointer).sAno
If nPointer = 1 Then
    cmdD1.Enabled = False
Else
    cmdD1.Enabled = True
End If

If nPointer = UBound(aPic) Then
    cmdD2.Enabled = False
Else
    cmdD2.Enabled = True
End If

'If NomeDoComputador = "MATHWORLD" Or NomeDoComputador = "GTI" Then Exit Sub
If NomeDoComputador = "SKYNET" Then
    sPathOrigem = "\\200.232.123.115\atualizagti\Documentos\"
Else
    sPathOrigem = "\\192.168.200.130\atualizagti\Documentos\"
End If
sAno = aPic(nPointer).sAno
sArq = aPic(nPointer).sArq
sExt = aPic(nPointer).sExt
sTipo = aPic(nPointer).sTipo
On Error Resume Next
If UCase(sExt) = "JPG" Then
    imgTmp.Picture = LoadPicture(sPathOrigem & sAno & "\" & sArq)
Else
    imgTmp.Picture = imgTmp2.Picture
End If

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Function IsMEI(nCodigo As Long) As Boolean
Dim nRet As Boolean, Sql As String, RdoAux As rdoResultset
nRet = False

'ConectaEicon
'Sql = "select * from GTI_EICON..tb_inter_empr_mei_giss where NUM_CADASTRO=" & nCodigo & " order by TIMESTAMP desc"
'Set RdoAux = cnEicon.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
Sql = "select * from PERIODOMEI where CODIGO=" & nCodigo & " order by id desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!Datafim) Then
        nRet = True
    End If
   .Close
End With
'cnEicon.Close
IsMEI = nRet

End Function
