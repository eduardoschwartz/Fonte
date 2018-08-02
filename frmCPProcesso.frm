VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCPProcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle dos Processos (Setor de Divida Ativa)"
   ClientHeight    =   4830
   ClientLeft      =   5130
   ClientTop       =   2505
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   8775
   Begin VB.Frame pnlPrint 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Impressão de Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1890
      Left            =   1695
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   4485
      Begin VB.CheckBox chkP4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Comprovante de entrega de Documento"
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
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   1050
         Width           =   3765
      End
      Begin VB.CheckBox chkP3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Comunicado de entrega de Documento"
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
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   705
         Width           =   3765
      End
      Begin VB.CheckBox chkP1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Protocolo de Entrada do Processo"
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
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   360
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin prjChameleon.chameleonButton cmdOKPrint 
         Height          =   345
         Left            =   3540
         TabIndex        =   4
         ToolTipText     =   "Imprimir os documentos selecionados"
         Top             =   1395
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":0000
         PICN            =   "frmCPProcesso.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancelPrint 
         Height          =   345
         Left            =   3960
         TabIndex        =   5
         ToolTipText     =   "Cancelar operação"
         Top             =   1395
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":0176
         PICN            =   "frmCPProcesso.frx":0192
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
   Begin VB.Frame frReq1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Setor/Depto./Secretaria"
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
      Height          =   1110
      Left            =   0
      TabIndex        =   107
      Top             =   3705
      Width           =   7350
      Begin VB.TextBox txtCidadao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6030
         MaxLength       =   6
         TabIndex        =   109
         Top             =   225
         Width           =   945
      End
      Begin VB.ComboBox cmbReq 
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   210
         Width           =   4650
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome..:"
         Height          =   225
         Index           =   16
         Left            =   135
         TabIndex        =   112
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblNomeCidadao 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         ForeColor       =   &H00000040&
         Height          =   225
         Left            =   720
         TabIndex        =   111
         Top             =   675
         Width           =   6330
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidadão..:"
         Height          =   225
         Index           =   15
         Left            =   5175
         TabIndex        =   110
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame frReq2 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Requerente do Processo"
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
      Height          =   615
      Left            =   15
      TabIndex        =   81
      Top             =   3720
      Width           =   7305
      Begin prjChameleon.chameleonButton cmdEditCid 
         Height          =   315
         Left            =   5715
         TabIndex        =   82
         ToolTipText     =   "Editar requerente do processo"
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
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
         MICON           =   "frmCPProcesso.frx":02EC
         PICN            =   "frmCPProcesso.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGuia 
         Height          =   315
         Left            =   6840
         TabIndex        =   83
         ToolTipText     =   "Gerar guia"
         Top             =   180
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "$"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":0462
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCidadaoCO 
         Height          =   315
         Left            =   6525
         TabIndex        =   84
         ToolTipText     =   "Exibir cidadão gravado no processo original"
         Top             =   180
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "O"
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
         MICON           =   "frmCPProcesso.frx":047E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblNomeCid 
         BackStyle       =   0  'Transparent
         Caption         =   "MARCELA DE SOUZA BRITO CARVALHO"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   780
         TabIndex        =   86
         Top             =   300
         Width           =   4890
      End
      Begin VB.Label lblCodCid 
         BackStyle       =   0  'Transparent
         Caption         =   "523888"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   85
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Datas das Ocorrências"
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
      Height          =   825
      Left            =   15
      TabIndex        =   64
      Top             =   1680
      Width           =   7305
      Begin prjChameleon.chameleonButton cmdOC 
         Height          =   225
         Left            =   4590
         TabIndex        =   65
         ToolTipText     =   "Observação de Cancelamento"
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   14
         TX              =   "!"
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
         MICON           =   "frmCPProcesso.frx":049A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOS 
         Height          =   225
         Left            =   2160
         TabIndex        =   66
         ToolTipText     =   "Observação de Suspensão"
         Top             =   520
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   14
         TX              =   "!"
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
         MICON           =   "frmCPProcesso.frx":04B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOR 
         Height          =   225
         Left            =   4590
         TabIndex        =   67
         ToolTipText     =   "Observação de Reativação"
         Top             =   255
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   14
         TX              =   "!"
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
         MICON           =   "frmCPProcesso.frx":04D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOA 
         Height          =   225
         Left            =   6975
         TabIndex        =   68
         ToolTipText     =   "Observação de Arquivamento"
         Top             =   525
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   397
         BTYPE           =   14
         TX              =   "!"
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
         MICON           =   "frmCPProcesso.frx":04EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblDtCancelamento 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3645
         TabIndex        =   80
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDtReativacao 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3645
         TabIndex        =   79
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblDtArquivamento 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   6045
         TabIndex        =   78
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDtSuspencao 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1245
         TabIndex        =   77
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDtEntrada 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1245
         TabIndex        =   76
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Suspenção.....:"
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   75
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivamento..:"
         Height          =   225
         Index           =   10
         Left            =   4905
         TabIndex        =   74
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelamento.:"
         Height          =   225
         Index           =   9
         Left            =   2460
         TabIndex        =   73
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reativação.....:"
         Height          =   225
         Index           =   8
         Left            =   2490
         TabIndex        =   72
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada...........:"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   71
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexos............:"
         Height          =   225
         Index           =   13
         Left            =   4905
         TabIndex        =   70
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   6060
         TabIndex        =   69
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Endereços de Ocorrência"
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
      Height          =   1065
      Left            =   -45
      TabIndex        =   61
      Top             =   5535
      Visible         =   0   'False
      Width           =   7350
      Begin MSFlexGridLib.MSFlexGrid grdEnd 
         Height          =   765
         Left            =   90
         TabIndex        =   62
         Top             =   270
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   1349
         _Version        =   393216
         Rows            =   3
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "^Código   |<Nome do Logradouro                                                      |>Número   "
      End
      Begin prjChameleon.chameleonButton cmdEditEnd 
         Height          =   315
         Left            =   6030
         TabIndex        =   63
         ToolTipText     =   "Editar endereço de ocorrência"
         Top             =   600
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "frmCPProcesso.frx":050A
         PICN            =   "frmCPProcesso.frx":0526
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Informações Gerais do Processo"
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
      Height          =   1665
      Left            =   15
      TabIndex        =   44
      Top             =   45
      Width           =   7305
      Begin VB.CheckBox chkInterno 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Processo Interno..:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   49
         Top             =   600
         Width           =   1635
      End
      Begin VB.CheckBox chkFisico 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Processo Físico..:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.ComboBox cmbOrigem 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmCPProcesso.frx":0680
         Left            =   4590
         List            =   "frmCPProcesso.frx":0682
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   540
         Width           =   2565
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1530
         MaxLength       =   150
         TabIndex        =   46
         Top             =   1260
         Width           =   5595
      End
      Begin VB.ComboBox cmbAssunto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1530
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   900
         Width           =   5625
      End
      Begin VB.Label lblAtendente 
         BackStyle       =   0  'Transparent
         Caption         =   "NOME"
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
         Height          =   225
         Left            =   4410
         TabIndex        =   60
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblAno 
         BackStyle       =   0  'Transparent
         Caption         =   "2005"
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
         Height          =   255
         Left            =   2940
         TabIndex        =   59
         Top             =   300
         Width           =   585
      End
      Begin VB.Label lblNumProc 
         BackStyle       =   0  'Transparent
         Caption         =   "35"
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
         Height          =   255
         Left            =   1530
         TabIndex        =   58
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento......:"
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   57
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto...............:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   56
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente..:"
         Height          =   225
         Index           =   5
         Left            =   3510
         TabIndex        =   55
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Origem....:"
         Height          =   225
         Index           =   3
         Left            =   3780
         TabIndex        =   54
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano..:"
         Height          =   225
         Index           =   1
         Left            =   2460
         TabIndex        =   53
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo...:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   52
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         Height          =   225
         Index           =   14
         Left            =   6150
         TabIndex        =   51
         Top             =   300
         Width           =   435
      End
      Begin VB.Label lblHora 
         BackStyle       =   0  'Transparent
         Caption         =   "12:35"
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
         Height          =   255
         Left            =   6630
         TabIndex        =   50
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4665
      Left            =   7365
      TabIndex        =   30
      Top             =   75
      Width           =   1365
      Begin prjChameleon.chameleonButton cmdExcluir 
         Height          =   315
         Left            =   90
         TabIndex        =   31
         ToolTipText     =   "Excluir Registro"
         Top             =   780
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmCPProcesso.frx":0684
         PICN            =   "frmCPProcesso.frx":06A0
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
         Left            =   135
         TabIndex        =   32
         ToolTipText     =   "Sair da Tela"
         Top             =   4230
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmCPProcesso.frx":0742
         PICN            =   "frmCPProcesso.frx":075E
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
         Left            =   135
         TabIndex        =   33
         ToolTipText     =   "Consulta Processos Cadastrados"
         Top             =   3870
         Width           =   1185
         _ExtentX        =   2090
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":07CC
         PICN            =   "frmCPProcesso.frx":07E8
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
         TabIndex        =   34
         ToolTipText     =   "Novo Registro"
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmCPProcesso.frx":0942
         PICN            =   "frmCPProcesso.frx":095E
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
         Left            =   90
         TabIndex        =   35
         ToolTipText     =   "Editar Registro"
         Top             =   420
         Width           =   1185
         _ExtentX        =   2090
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
         MICON           =   "frmCPProcesso.frx":0AB8
         PICN            =   "frmCPProcesso.frx":0AD4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdTramite 
         Height          =   315
         Left            =   90
         TabIndex        =   36
         ToolTipText     =   "Tramitar um Processo"
         Top             =   1140
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Tramitar"
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
         MICON           =   "frmCPProcesso.frx":0C2E
         PICN            =   "frmCPProcesso.frx":0C4A
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
         Left            =   90
         TabIndex        =   37
         ToolTipText     =   "Cancelar um Processo"
         Top             =   2340
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
         MICON           =   "frmCPProcesso.frx":0DA4
         PICN            =   "frmCPProcesso.frx":0DC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdArquivar 
         Height          =   315
         Left            =   90
         TabIndex        =   38
         ToolTipText     =   "Arquivar um Processo"
         Top             =   1980
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Arquivar  "
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
         MICON           =   "frmCPProcesso.frx":0F1A
         PICN            =   "frmCPProcesso.frx":0F36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSuspender 
         Height          =   315
         Left            =   90
         TabIndex        =   39
         ToolTipText     =   "Suspender um Processo"
         Top             =   2700
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "S&uspender"
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
         MICON           =   "frmCPProcesso.frx":0FD6
         PICN            =   "frmCPProcesso.frx":0FF2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdReativar 
         Height          =   315
         Left            =   90
         TabIndex        =   40
         ToolTipText     =   "Reativar um Processo"
         Top             =   3060
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Reativar  "
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
         MICON           =   "frmCPProcesso.frx":1091
         PICN            =   "frmCPProcesso.frx":10AD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdImprimir 
         Height          =   315
         Left            =   90
         TabIndex        =   41
         ToolTipText     =   "Impressão do Protocolo de Entrada e Requerimento"
         Top             =   1500
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Imprimir"
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
         MICON           =   "frmCPProcesso.frx":1121
         PICN            =   "frmCPProcesso.frx":113D
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
         Left            =   90
         TabIndex        =   42
         ToolTipText     =   "Gravar o Registro"
         Top             =   60
         Width           =   1185
         _ExtentX        =   2090
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":1297
         PICN            =   "frmCPProcesso.frx":12B3
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
         Height          =   315
         Left            =   90
         TabIndex        =   43
         ToolTipText     =   "Cancelar Edição"
         Top             =   420
         Width           =   1185
         _ExtentX        =   2090
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":1658
         PICN            =   "frmCPProcesso.frx":1674
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
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Observações"
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
      Height          =   1215
      Left            =   30
      TabIndex        =   28
      Top             =   2490
      Width           =   7305
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   90
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   270
         Width           =   7110
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Documentos"
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
      Height          =   585
      Left            =   0
      TabIndex        =   22
      Top             =   5760
      Visible         =   0   'False
      Width           =   7350
      Begin prjChameleon.chameleonButton cmdEditDoc 
         Height          =   315
         Left            =   6030
         TabIndex        =   23
         ToolTipText     =   "Editar documentos"
         Top             =   180
         Width           =   1065
         _ExtentX        =   1879
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
         MICON           =   "frmCPProcesso.frx":17CE
         PICN            =   "frmCPProcesso.frx":17EA
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
         Caption         =   "Documentos Entregues..:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Pendentes..:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2670
         TabIndex        =   26
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label lblDoc1 
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2010
         TabIndex        =   25
         Top             =   270
         Width           =   285
      End
      Begin VB.Label lblDoc2 
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4620
         TabIndex        =   24
         Top             =   270
         Width           =   285
      End
   End
   Begin VB.Frame frCidadao 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cidadão original gravado no processo"
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
      Height          =   2445
      Left            =   525
      TabIndex        =   6
      Top             =   2370
      Visible         =   0   'False
      Width           =   6495
      Begin prjChameleon.chameleonButton cmdFecharCO 
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         ToolTipText     =   "Fechar esta Tela"
         Top             =   2025
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Fechar"
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
         MICON           =   "frmCPProcesso.frx":1944
         PICN            =   "frmCPProcesso.frx":1960
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
         Caption         =   "Nome.........:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   21
         Top             =   405
         Width           =   915
      End
      Begin VB.Label lblCONome 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   20
         Top             =   405
         Width           =   5280
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ.:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   19
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   18
         Top             =   1035
         Width           =   915
      End
      Begin VB.Label lblCODoc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "R.G..:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2790
         TabIndex        =   16
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblCORG 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3330
         TabIndex        =   15
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label lblCOEnd 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   14
         Top             =   1035
         Width           =   5280
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Complem....:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label lblCOCompl 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   12
         Top             =   1350
         Width           =   5280
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro.........:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   1665
         Width           =   915
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade..:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3285
         TabIndex        =   10
         Top             =   1665
         Width           =   735
      End
      Begin VB.Label lblCOBairro 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   9
         Top             =   1665
         Width           =   2130
      End
      Begin VB.Label lblCOCidade 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4005
         TabIndex        =   8
         Top             =   1665
         Width           =   2355
      End
   End
   Begin VB.Frame pnlDoc 
      BackColor       =   &H00E4FEFC&
      Caption         =   "Documentos Necessários"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2445
      Left            =   1035
      TabIndex        =   92
      Top             =   2010
      Width           =   6165
      Begin MSComctlLib.ListView lvDoc 
         Height          =   2055
         Left            =   90
         TabIndex        =   93
         Top             =   300
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   3625
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrição do Documento"
            Object.Width           =   7409
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Data Entrega"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame pnlObs 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Observação da Ocorrência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2775
      Left            =   645
      TabIndex        =   87
      Top             =   1860
      Visible         =   0   'False
      Width           =   6525
      Begin VB.TextBox txtObsData 
         Appearance      =   0  'Flat
         Height          =   1905
         Left            =   90
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   88
         Top             =   360
         Width           =   6315
      End
      Begin prjChameleon.chameleonButton cmdGravarObs 
         Height          =   345
         Left            =   5580
         TabIndex        =   89
         ToolTipText     =   "Gravar Observação"
         Top             =   2340
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":19CE
         PICN            =   "frmCPProcesso.frx":19EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancelarObs 
         Height          =   345
         Left            =   5985
         TabIndex        =   90
         ToolTipText     =   "Cancelar operação"
         Top             =   2340
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   609
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCPProcesso.frx":1B44
         PICN            =   "frmCPProcesso.frx":1B60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblOcor 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   120
         TabIndex        =   91
         Top             =   2400
         Width           =   4365
      End
   End
   Begin VB.Frame pnlEnd 
      BackColor       =   &H00E4FEFC&
      Caption         =   "Endereços de Ocorrência"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2325
      Left            =   0
      TabIndex        =   94
      Top             =   5580
      Width           =   6225
      Begin VB.ListBox lstNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         ItemData        =   "frmCPProcesso.frx":1CBA
         Left            =   2040
         List            =   "frmCPProcesso.frx":1CBC
         TabIndex        =   98
         Top             =   585
         Visible         =   0   'False
         Width           =   3990
      End
      Begin VB.TextBox txtNumeroLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         MaxLength       =   15
         TabIndex        =   97
         Top             =   1740
         Width           =   1305
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   96
         Top             =   1425
         Width           =   3990
      End
      Begin VB.TextBox txtNumLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         TabIndex        =   95
         Top             =   1425
         Width           =   855
      End
      Begin prjChameleon.chameleonButton cmdAlterar2 
         Height          =   315
         Left            =   4260
         TabIndex        =   99
         ToolTipText     =   "Editar Registro"
         Top             =   1830
         Width           =   885
         _ExtentX        =   1561
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
         MICON           =   "frmCPProcesso.frx":1CBE
         PICN            =   "frmCPProcesso.frx":1CDA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid grdEnd2 
         Height          =   975
         Left            =   120
         TabIndex        =   100
         Top             =   330
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1720
         _Version        =   393216
         Rows            =   4
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   15007484
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "^Código   |<Nome do Logradouro                                                      |>Número   "
      End
      Begin prjChameleon.chameleonButton cmdCancel2 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   5160
         TabIndex        =   101
         ToolTipText     =   "Cancelar Edição"
         Top             =   1830
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Canc."
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
         MICON           =   "frmCPProcesso.frx":1E34
         PICN            =   "frmCPProcesso.frx":1E50
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNovo2 
         Height          =   315
         Left            =   3360
         TabIndex        =   102
         ToolTipText     =   "Novo Registro"
         Top             =   1830
         Width           =   885
         _ExtentX        =   1561
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
         MICON           =   "frmCPProcesso.frx":1FAA
         PICN            =   "frmCPProcesso.frx":1FC6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExcluir2 
         Height          =   315
         Left            =   5160
         TabIndex        =   103
         ToolTipText     =   "Excluir Registro"
         Top             =   1830
         Width           =   885
         _ExtentX        =   1561
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
         MICON           =   "frmCPProcesso.frx":2120
         PICN            =   "frmCPProcesso.frx":213C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGravar2 
         Height          =   315
         Left            =   4260
         TabIndex        =   104
         ToolTipText     =   "Gravar os Dados"
         Top             =   1830
         Width           =   885
         _ExtentX        =   1561
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
         MICON           =   "frmCPProcesso.frx":21DE
         PICN            =   "frmCPProcesso.frx":21FA
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
         Caption         =   "Número.......:"
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   106
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   105
         Top             =   1455
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmCPProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, sql As String, RdoAux2 As rdoResultset
Dim Evento As String, Evento2 As String, bExec As Boolean

Private Sub cmbAssunto_Click()
If cmbAssunto.ListIndex = -1 Then Exit Sub
txtCompl.Text = cmbAssunto.Text
End Sub

Private Sub cmbOrigem_Click()
If cmbOrigem.ListIndex = 0 Then
    frReq1.Visible = True:    frReq2.Visible = False
Else
    frReq1.Visible = False:    frReq2.Visible = True
End If
End Sub

Private Sub cmdAlterar_Click()
If lblAno.Caption = "" Or lblNumProc.Caption = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(lblDtCancelamento.Caption) Then
   MsgBox "Processo Cancelado.", vbExclamation, "Atenção"
   Exit Sub
End If

If IsDate(lblDtSuspencao.Caption) Then
   MsgBox "Processo Suspenso.", vbExclamation, "Atenção"
   Exit Sub
End If
Evento = "Alterar"
Eventos "INCLUIR"
If IsDate(lblDtArquivamento.Caption) Then
    cmdEditCid.Enabled = False
    cmdEditDoc.Enabled = False
    cmdEditEnd.Enabled = False
    txtCompl.Locked = False
    txtCompl.BackColor = Branco
    cmdRepair.Enabled = False
    chkFisico.Enabled = False
    
End If
End Sub

Private Sub cmdAlterar2_Click()

If Val(txtNumLog.Text) = 0 Then
    MsgBox "Selecione um endereço.", vbExclamation, "Atenção"
    Exit Sub
End If

Evento2 = "Alterar"
Eventos2 "INCLUIR"
txtNumLog.SetFocus

End Sub


Private Sub cmdArquivar_Click()
Dim z As Variant, t As Variant
If lblAno.Caption = "" Or lblNumProc = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(lblDtArquivamento.Caption) Then
   MsgBox "Processo já Arquivado.", vbExclamation, "Atenção"
   Exit Sub
Else
    If IsDate(lblDtCancelamento.Caption) Then
        MsgBox "Não é possivel arquivar, processo cancelado.", vbExclamation, "Atenção"
    Else
        If Not VerificaTramite Then
            MsgBox "Não é possível arquivar, tramite não concluido", vbExclamation, "Atenção"
        Else
            If MsgBox("Arquivar este processo ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
Inicio:
                t = InputBox("Digite o motivo.", "OBSERVAÇÃO DE ARQUIVAMENTO")
                If t = "" Then
                   MsgBox "Observação obrigatória", vbExclamation, "Atenção"
                   GoTo Inicio
                Else
                    z = Format(Now, "dd/mm/yyyy")
                    sql = "UPDATE CPPROCESSOGTI SET DATAARQUIVA='" & Format(z, "mm/dd/yyyy") & "',OBSA='" & Mask(CStr(t)) & "' WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                    cn.Execute sql, rdExecDirect
                    lblDtArquivamento = z
                 End If
            End If
         End If
     End If
End If

End Sub

Private Sub cmdCancel_Click()

Eventos "INICIAR"
Evento = ""
HabilitaPainelPrincipal
If lblNumProc.Caption <> "" Then
   CodProcessoCP = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
   Limpa
   Le
End If

End Sub

Private Sub cmdCancel2_Click()
If grdEnd2.Rows > 1 Then grdEnd2_Click
Eventos2 "INICIAR"
cmdEditEnd.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
Dim t As Variant, z As Variant
If lblAno.Caption = "" Or lblNumProc = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(lblDtCancelamento.Caption) Then
   MsgBox "Processo já Cancelado.", vbExclamation, "Atenção"
   Exit Sub
Else
    If IsDate(lblDtArquivamento.Caption) Then
       MsgBox "Processo Arquivado.", vbExclamation, "Atenção"
       Exit Sub
    Else
        If MsgBox("Cancelar este processo ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
Inicio:
            t = InputBox("Digite o motivo.", "OBSERVAÇÃO DE CANCELAMENTO")
            If t = "" Then
               MsgBox "Observação obrigatória", vbExclamation, "Atenção"
               GoTo Inicio
            Else
                z = Format(Now, "dd/mm/yyyy")
                sql = "UPDATE CPPROCESSOGTI SET DATACANCEL='" & Format(z, "mm/dd/yyyy") & "', OBSC='" & Mask(CStr(t)) & "' WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                cn.Execute sql, rdExecDirect
                lblDtCancelamento = z
            End If
        End If
    End If
End If

End Sub

Private Sub cmdCancelarObs_Click()
pnlObs.Visible = False
End Sub

Private Sub cmdCancelPrint_Click()
HabilitaPainelPrincipal
pnlPrint.Visible = False
End Sub

Private Sub cmdCloseObs_Click()
pnlObs.Visible = False
End Sub

Private Sub cmdCidadaoCO_Click()
Dim sql As String, RdoAux As rdoResultset, nNumProc As Integer, sDoc As String

If Val(lblAno.Caption) = 0 Then Exit Sub
If cmbReq.ListIndex > -1 Then Exit Sub
If frmCPProcesso.lblNumProc.Caption = "" Then Exit Sub
nNumProc = Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
sql = "SELECT cpprocessocidadao.anoproc, cpprocessocidadao.numproc, cpprocessocidadao.codcidadao, cpprocessocidadao.nomecidadao, cpprocessocidadao.doc,cpprocessocidadao.RG,cpprocessocidadao.ORGAO, "
sql = sql & "cpprocessocidadao.numimovel, cpprocessocidadao.complemento, cpprocessocidadao.codbairro, cpprocessocidadao.codcidade, cpprocessocidadao.siglauf,"
sql = sql & "cpprocessocidadao.cep, vwLOGRADOURO.abrevtipolog, vwLOGRADOURO.abrevtitlog, vwLOGRADOURO.nomelogradouro, cpprocessocidadao.codlogradouro,"
sql = sql & "cidade.desccidade , bairro.DescBairro FROM cpprocessocidadao INNER JOIN vwLOGRADOURO ON cpprocessocidadao.codlogradouro = vwLOGRADOURO.codlogradouro LEFT OUTER JOIN "
sql = sql & "bairro ON cpprocessocidadao.siglauf = bairro.siglauf AND cpprocessocidadao.codcidade = bairro.codcidade AND "
sql = sql & "cpprocessocidadao.codbairro = bairro.codbairro LEFT OUTER JOIN cidade ON cpprocessocidadao.siglauf = cidade.siglauf AND cpprocessocidadao.codcidade = cidade.codcidade WHERE ANOPROC=" & Val(lblAno.Caption) & " AND NUMPROC=" & nNumProc
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblCONome.Caption = !nomecidadao
        sDoc = Trim(!Doc)
        If Len(sDoc) = 11 Then
            sDoc = Format(Trim(sDoc), "000\.000\.000-00")
        ElseIf Len(sDoc) = 14 Then
            sDoc = Format(Trim(sDoc), "00\.000\.000/0000-00")
        Else
            sDoc = ""
        End If
        lblCODoc.Caption = sDoc
        lblCOEnd.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & ", " & !NUMIMOVEL
        lblCOCompl.Caption = !Complemento
        lblCOBairro.Caption = SubNull(!DescBairro)
        lblCOCidade.Caption = SubNull(!descCidade) & "/" & SubNull(!SiglaUF)
        lblCORG.Caption = !rg & " - " & !ORGAO
    Else
        MsgBox "Não existem informações gravadas sobre o cidadão de origem, apenas processos a partir de 24/01/2011 possuem esta informação.", vbInformation, "Atenção"
       .Close
       Exit Sub
    End If
   .Close
End With

frCidadao.Visible = True
frCidadao.ZOrder (0)

End Sub

Private Sub cmdConsultar_Click()
frmCPCnsProcesso.show
frmCPCnsProcesso.ZOrder 0
End Sub

Private Sub cmdEditCid_Click()

If Val(lblAno.Caption) = 0 Then Exit Sub

DesabilitaPainelPrincipal
cmdEditEnd.Enabled = False
cmdEditCid.Enabled = False
cmdEditDoc.Enabled = False

CodCidadao = Val(lblCodCid.Caption)
Set frm2 = frmCidadao
frm2.sForm = Me.Name
If Evento = "Novo" Then
   frm2.sTipoCidadao = "N"
Else
   frm2.sTipoCidadao = "A"
End If
frmCidadao.show
frmCidadao.ZOrder 0

End Sub

Private Sub cmdEditDoc_Click()
Dim nSim As Integer, nNao As Integer, x As Integer
If Val(lblAno.Caption) = 0 Then Exit Sub

nSim = 0: nNao = 0

If Not pnlDoc.Visible Then
    If Evento = "Novo" Then
        If cmbAssunto.ListIndex > -1 Then
            z = SendMessage(lvDoc.hwnd, LVM_DELETEALLITEMS, 0, 0)
            sql = "SELECT CPASSUNTODOC.CODDOC, CPDOCUMENTO.NOME FROM CPDOCUMENTO INNER JOIN "
            sql = sql & "CPASSUNTODOC ON CPDOCUMENTO.CODIGO = CPASSUNTODOC.CODDOC INNER JOIN CPASSUNTO ON "
            sql = sql & "CPASSUNTODOC.CODASSUNTO = CPASSUNTO.CODIGO Where CPASSUNTO.Codigo = " & cmbAssunto.ItemData(cmbAssunto.ListIndex)
            Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                    Set itmX = lvDoc.ListItems.Add(, "C" & Format(!CODDOC, "000"), !nome)
                   .MoveNext
                Loop
               .Close
            End With
        Else
            MsgBox "Selecione o Assunto.", vbExclamation, "Atenção"
            Exit Sub
        End If
    End If
    DesabilitaPainelPrincipal
    cmdEditEnd.Enabled = False
    cmdEditCid.Enabled = False
    pnlDoc.Visible = True
    pnlDoc.ZOrder 0
Else
    HabilitaPainelPrincipal
    If Evento <> "Novo" Then
        cmbOrigem.BackColor = Kde
        cmbAssunto.BackColor = Kde
        chkFisico.BackColor = Kde
        chkInterno.BackColor = Kde
        txtCompl.BackColor = Kde
        cmbOrigem.Enabled = False
        cmbAssunto.Enabled = False
        chkFisico.Enabled = False
        chkInterno.Enabled = False
        txtCompl.Locked = True
    End If
    cmdEditEnd.Enabled = True
    cmdEditCid.Enabled = True
    For x = 1 To lvDoc.ListItems.Count
        If lvDoc.ListItems(x).Checked = True Then
            nSim = nSim + 1
        Else
            lvDoc.ListItems(x).SubItems(1) = ""
            nNao = nNao + 1
        End If
    Next
    lblDoc1.Caption = nSim
    lblDoc2.Caption = nNao
    pnlDoc.Visible = False
End If

End Sub

Private Sub cmdEditEnd_Click()
Dim x As Integer

If Val(lblAno.Caption) = 0 Then Exit Sub

If Not pnlEnd.Visible Then
    DesabilitaPainelPrincipal
    
    grdEnd2.Rows = 1
    With grdEnd
        For x = 1 To .Rows - 1
            grdEnd2.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2)
        Next
    End With
    If grdEnd2.Rows > 1 Then grdEnd2_Click
    Eventos2 "INICIAR"
    txtNumLog.Text = "": txtNomeLog.Text = "": txtNumeroLog.Text = ""
    pnlEnd.Visible = True
    pnlEnd.ZOrder 0
Else
    
    HabilitaPainelPrincipal
    
    grdEnd.Rows = 1
    With grdEnd2
        For x = 1 To .Rows - 1
            grdEnd.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2)
        Next
    End With
    pnlEnd.Visible = False
End If

End Sub

Private Sub cmdExcluir_Click()

If lblAno.Caption = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

sql = "SELECT ANO,NUMERO,SEQ FROM CPTRAMITACAO WHERE ANO=" & Val(lblAno.Caption)
sql = sql & " AND NUMERO=" & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        .Close
        MsgBox "Processo possue tramite e não pode ser excluido.", vbExclamation, "Atenção"
        Exit Sub
    End If
   .Close
End With

If MsgBox("Excluir este  Processo ???", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE EXCLUSÃO") = vbNo Then Exit Sub

sql = "DELETE FROM CPPROCESSOEND WHERE ANO=" & lblAno.Caption & " AND NUMPROCESSO=" & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
cn.Execute sql, rdExecDirect
sql = "DELETE FROM CPPROCESSODOC WHERE ANO=" & lblAno.Caption & " AND NUMERO=" & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
cn.Execute sql, rdExecDirect
sql = "DELETE FROM CPPROCESSOGTI WHERE ANO=" & lblAno.Caption & " AND NUMERO=" & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
cn.Execute sql, rdExecDirect
Log Form, Me.Caption, Exclusão, "Excluído processo '" & frmCPProcesso.lblNumProc.Caption & "'"

Limpa

End Sub

Private Sub cmdExcluir2_Click()
If Val(txtNumLog.Text) = 0 Then
    MsgBox "Selecione um endereço.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Remover este endereço ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    If grdEnd2.Rows > 2 Then
       grdEnd2.RemoveItem (grdEnd2.Row)
    Else
       grdEnd2.Rows = 1
    End If
    grdEnd2_Click
End If

End Sub

Private Sub cmdFecharCO_Click()
frCidadao.Visible = False
End Sub

Private Sub cmdGravar_Click()
Dim sql As String, RdoAux As rdoResultset
If cmbOrigem.ListIndex = -1 Then
    MsgBox "Selecione a origem do processo.", vbExclamation, "Atenção"
    cmbOrigem.SetFocus
    Exit Sub
End If

If cmbAssunto.ListIndex = -1 Then
    MsgBox "Selecione o Assunto do processo.", vbExclamation, "Atenção"
    cmbAssunto.SetFocus
    Exit Sub
End If

If frReq2.Visible = True Then
    If lblCodCid.Caption = "" Then
        MsgBox "Selecione o requerente do processo.", vbExclamation, "Atenção"
        Exit Sub
    End If
Else
    If cmbReq.ListIndex = -1 Then
        MsgBox "Selecione o requerente do processo.", vbExclamation, "Atenção"
        Exit Sub
    Else
        sql = "SELECT * FROM CPCENTROCUSTO WHERE CODIGO=" & cmbReq.ItemData(cmbReq.ListIndex)
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux!Ativo = 0 Then
            MsgBox "Este centro de custos esta desativado e não pode ser utilizado.", vbCritical, "Atenção"
            Exit Sub
        End If
    End If
End If

'If Val(txtCidadao.Text) > 0 Then
'    If Val(txtCidadao.Text) < 500000 Or Val(txtCidadao.Text) >= 6000000 Then
'        MsgBox "Código de cidadão inválido.", vbCritical, "Atenção"
'        Exit Sub
'    End If
'End If

Grava

Eventos "INICIAR"
Evento = ""
End Sub

Private Sub cmdGravar2_Click()
Dim x As Integer, bAchou As Boolean

If Val(txtNumLog.Text) = 0 Then
    MsgBox "Selecione o logradouro.", vbExclamation, "Atenção"
    Exit Sub
End If
If txtNumeroLog.Text = "" Then
    MsgBox "Digite o número do logradouro.", vbExclamation, "Atenção"
    Exit Sub
End If
'If Val(txtNumeroLog.text) > 3000 Then
'    MsgBox "Número do logradouro inválido.", vbExclamation, "Atenção"
'    Exit Sub
'End If

If Evento2 = "Novo" Then
    bAchou = False
    With grdEnd2
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) = Val(txtNumLog.Text) And UCase$(.TextMatrix(x, 2)) = UCase$(txtNumeroLog.Text) Then
                bAchou = True
            End If
        Next
    End With
    If Not bAchou Then
        grdEnd2.AddItem txtNumLog.Text & Chr(9) & txtNomeLog.Text & Chr(9) & txtNumeroLog.Text
    Else
        MsgBox "Endereço ja cadastrado.", vbExclamation, "Atenção"
        Exit Sub
    End If
Else
    With grdEnd2
        .TextMatrix(.Row, 0) = txtNumLog.Text
        .TextMatrix(.Row, 1) = txtNomeLog.Text
        .TextMatrix(.Row, 2) = txtNumeroLog.Text
    End With
End If

Eventos2 "INICIAR"
cmdEditEnd.Enabled = True
End Sub

Private Sub cmdGravarObs_Click()
sql = "UPDATE CPPROCESSOGTI SET "
Select Case Left(lblOcor.Caption, 1)
    Case "A"
            sql = sql & "OBSA='"
    Case "C"
            sql = sql & "OBSC='"
    Case "S"
            sql = sql & "OBSS='"
    Case "R"
            sql = sql & "OBSR='"
End Select
sql = sql & Mask(txtObsData.Text) & "' WHERE  ANO=" & Val(lblAno.Caption)
sql = sql & " AND NUMERO=" & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
cn.Execute sql, rdExecDirect

pnlObs.Visible = False
End Sub

Private Sub cmdGuia_Click()
frm2ViaLaser.txtCod.Text = lblCodCid.Caption
frm2ViaLaser.txtCod_LostFocus
End Sub

Private Sub cmdImprimir_Click()
If lblAno.Caption = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

DesabilitaPainelPrincipal
pnlPrint.Visible = True: pnlPrint.ZOrder (0)
End Sub

Private Sub cmdNovo_Click()
Limpa
CarregaCombo True
lblDtEntrada.Caption = Format(Now, "dd/mm/yyyy")
lblAno.Caption = Year(Now)
lblHora.Caption = Format(Now, "hh:mm")
'lblAtendente.Caption = Left$(Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8), 30)
lblAtendente.Caption = NomeDeLogin

Evento = "Novo"
Eventos "INCLUIR"
'chkFisico.SetFocus
cmbAssunto.SetFocus
If cmbOrigem.ListCount > 0 Then cmbOrigem.ListIndex = 0
End Sub

Private Sub cmdNovo2_Click()
txtNumLog.Text = ""
txtNomeLog.Text = ""
txtNumeroLog.Text = ""
Evento2 = "Novo"
Eventos2 "INCLUIR"
txtNumLog.SetFocus
End Sub

Private Sub cmdOA_Click()
If pnlObs.Visible = True Then
    pnlObs.Visible = False
    Exit Sub
End If
If IsDate(lblDtArquivamento.Caption) Then
    lblOcor.Caption = "ARQUIVAMENTO"
    pnlObs.Visible = True: pnlObs.ZOrder 0
    sql = "SELECT ANO,NUMERO,OBSA FROM CPPROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    sql = sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtObsData.Text = SubNull(!OBSA)
        Else
            txtObsData.Text = ""
        End If
       .Close
    End With
End If
End Sub

Private Sub cmdOC_Click()
If pnlObs.Visible = True Then
    pnlObs.Visible = False
    Exit Sub
End If
If IsDate(lblDtCancelamento.Caption) Then
    lblOcor.Caption = "CANCELAMENTO"
    pnlObs.Visible = True: pnlObs.ZOrder 0
    sql = "SELECT ANO,NUMERO,OBSC FROM CPPROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    sql = sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtObsData.Text = SubNull(!OBSC)
        Else
            txtObsData.Text = ""
        End If
       .Close
    End With
End If
End Sub

Private Sub cmdOKPrint_Click()
HabilitaPainelPrincipal
pnlPrint.Visible = False

If chkP1.Value = 0 And chkP3.Value = 0 And chkP4.Value = 0 Then
    MsgBox "Nenhum relatório selecionado.", vbExclamation, "Atenção"
    Exit Sub
End If

If chkP1.Value = 1 Then
    frmReport.ShowReport "CPPROTOCOLOENTRADA", frmMdi.hwnd, Me.hwnd
End If
If chkP3.Value = 1 Then
    sql = "SELECT * From CPPROCESSODOC  Where ANO = " & Val(frmCPProcesso.lblAno.Caption) & " And Numero = " & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2)) & " And Data Is Null"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        frmReport.ShowReport "CPCOMUNICADODOC", frmMdi.hwnd, Me.hwnd
       .Close
    End With
End If
If chkP4.Value = 1 Then
    sql = "SELECT * From CPPROCESSODOC  Where ANO = " & Val(frmCPProcesso.lblAno.Caption) & " And Numero = " & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2)) & " And Data Is not Null"
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "Este processo não possue documentos entregues.", vbExclamation, "Atenção"
        Else
            frmReport.ShowReport "CPCOMPROVANTEDOC", frmMdi.hwnd, Me.hwnd
        End If
       .Close
    End With
End If

End Sub

Private Sub cmdOR_Click()
If pnlObs.Visible = True Then
    pnlObs.Visible = False
    Exit Sub
End If
If IsDate(lblDtReativacao.Caption) Then
    lblOcor.Caption = "REATIVAÇÃO"
    pnlObs.Visible = True: pnlObs.ZOrder 0
    sql = "SELECT ANO,NUMERO,OBSR FROM CPPROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    sql = sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtObsData.Text = SubNull(!OBSR)
        Else
            txtObsData.Text = ""
        End If
       .Close
    End With
End If
End Sub

Private Sub cmdOS_Click()
If pnlObs.Visible = True Then
    pnlObs.Visible = False
    Exit Sub
End If
If IsDate(lblDtSuspencao.Caption) Then
    lblOcor.Caption = "SUSPENSÃO"
    pnlObs.Visible = True: pnlObs.ZOrder 0
    sql = "SELECT ANO,NUMERO,OBSS FROM CPPROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    sql = sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtObsData.Text = SubNull(!OBSS)
        Else
            txtObsData.Text = ""
        End If
       .Close
    End With
End If
End Sub

Private Sub cmdReativar_Click()
Dim t As Variant, z As Variant
If lblAno.Caption = "" Or lblNumProc = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If
If Not IsDate(lblDtArquivamento.Caption) And Not IsDate(lblDtCancelamento.Caption) And Not IsDate(lblDtSuspencao.Caption) Then
   MsgBox "Processo encontra-se ativo.", vbExclamation, "Atenção"
   Exit Sub
Else
    If IsDate(lblDtCancelamento.Caption) Then
        MsgBox "Não é possivel reativar, processo cancelado.", vbExclamation, "Atenção"
    Else
        If MsgBox("Reativar este processo ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
Inicio:
                t = InputBox("Digite o motivo.", "OBSERVAÇÃO DE REATIVAÇÃO")
                If t = "" Then
                   MsgBox "Observação obrigatória", vbExclamation, "Atenção"
                   GoTo Inicio
                Else
                    z = Format(Now, "dd/mm/yyyy")
                    sql = "UPDATE CPPROCESSOGTI SET DATAREATIVA='" & Format(z, "mm/dd/yyyy") & "',DATAARQUIVA=NULL,DATASUSPENSO=NULL,DATACANCEL=NULL,OBSR='" & Mask(CStr(t)) & "' "
                    sql = sql & "WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                    cn.Execute sql, rdExecDirect
                    lblDtReativacao.Caption = z
                    lblDtArquivamento.Caption = "  /  /    "
                    lblDtCancelamento.Caption = "  /  /    "
                    lblDtSuspencao.Caption = "  /  /    "
                End If
        End If
     End If
End If

End Sub

Private Sub cmdSair_Click()
Unload frmCnsCidadao
Unload Me
End Sub

Private Sub cmdSuspender_Click()
Dim z As Variant, t As Variant
If lblAno.Caption = "" Or lblNumProc = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(lblDtSuspencao.Caption) Then
   MsgBox "Processo já Suspenso.", vbExclamation, "Atenção"
   Exit Sub
Else
    If IsDate(lblDtArquivamento.Caption) Then
       MsgBox "Processo Arquivado.", vbExclamation, "Atenção"
       Exit Sub
    Else
        If IsDate(lblDtCancelamento.Caption) Then
           MsgBox "Processo Cancelado.", vbExclamation, "Atenção"
           Exit Sub
        Else
            If MsgBox("Suspender este processo ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
Inicio:
                t = InputBox("Digite o motivo.", "OBSERVAÇÃO DE SUSPENÇÃO")
                If t = "" Then
                   MsgBox "Observação obrigatória", vbExclamation, "Atenção"
                   GoTo Inicio
                Else
                    z = Format(Now, "dd/mm/yyyy")
                    sql = "UPDATE CPPROCESSOGTI SET DATASUSPENSO='" & Format(z, "mm/dd/yyyy") & "',OBSS='" & Mask(CStr(t)) & "' WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                    cn.Execute sql, rdExecDirect
                    lblDtSuspencao.Caption = z
                End If
            End If
        End If
    End If
End If


End Sub

Private Sub cmdTramite_Click()

If lblAno.Caption = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Evento = "Novo" Then
    If cmbAssunto.ListIndex = -1 Then
        MsgBox "Selecione o Assunto.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If Val(lblCodCid.Caption) = 0 Then
        MsgBox "Selecione o Requerente.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

frmCPTramite.show
frmCPTramite.ZOrder 0
End Sub

Private Sub Form_Activate()
On Error Resume Next
If cmdConsultar.Visible = True And CodProcessoCP > 0 Then
    Limpa
    Le
End If

End Sub

Private Sub Form_Load()
frReq2.Visible = False
CodCidadao = 0
Limpa
Centraliza Me
sRet = RetEventUserForm(Me.Name)
CarregaCombo False
Eventos "INICIAR"
Eventos2 "INICIAR"
Evento = ""

End Sub

Private Sub Le()
On Error Resume Next
Dim k As Integer, sNomeLog As String
Dim nSim As Integer, nNao As Integer, RdoAux2 As rdoResultset
bExec = False: nSim = 0: nNao = 0
sql = "SELECT * FROM CPPROCESSOGTI WHERE NUMERO=" & CodProcessoCP & " AND ANO=" & AnoProcessoCP
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblNumProc.Caption = Format(CodProcessoCP & RetornaDVProcesso(CodProcessoCP), "000000-0")
        lblAno.Caption = AnoProcessoCP
        chkFisico.Value = IIf(!FISICO, 1, 0)
        chkInterno.Value = IIf(!INTERNO, 1, 0)
        If chkInterno.Value = 1 Then
            cmbOrigem.ListIndex = 0
        Else
            cmbOrigem.ListIndex = 0
        End If
        
        For k = 0 To cmbAssunto.ListCount - 1
            If cmbAssunto.ItemData(k) = !CODASSUNTO Then
               cmbAssunto.ListIndex = k
               Exit For
            End If
        Next
        If chkInterno.Value = 1 Then
            For k = 0 To cmbReq.ListCount - 1
                cmbReq.ListIndex = k
                If cmbReq.ItemData(cmbReq.ListIndex) = !CPCENTROCUSTO Then
                    Exit For
                End If
            Next
        End If
        txtCompl.Text = SubNull(!Complemento)
        txtObs.Text = SubNull(!OBSERVACAO)
        txtInsc.Text = SubNull(!INSC)
        lblAtendente.Caption = SubNull(!RESPONSAVEL)
        txtCidadao.Text = SubNull(!CodCidadao)
'        lblCodCid.Caption = !CodCidadao
        sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurReadOnly)
        lblNomeCidadao.Caption = SubNull(RdoAux2!nomecidadao)
        RdoAux2.Close
        If Not IsNull(!HORA) Then lblHora.Caption = !HORA
        If Not IsNull(!DATAENTRADA) Then lblDtEntrada.Caption = Format(!DATAENTRADA, "dd/mm/yyyy")
        If Not IsNull(!DATAREATIVA) Then lblDtReativacao.Caption = Format(!DATAREATIVA, "dd/mm/yyyy")
        If Not IsNull(!DataCancel) Then lblDtCancelamento.Caption = Format(!DataCancel, "dd/mm/yyyy")
        If Not IsNull(!DATAARQUIVA) Then lblDtArquivamento.Caption = Format(!DATAARQUIVA, "dd/mm/yyyy")
        If Not IsNull(!DATASUSPENSO) Then lblDtSuspencao.Caption = Format(!DATASUSPENSO, "dd/mm/yyyy")
        
       'CARREGA ENDEREÇO
        sql = "SELECT CPPROCESSOEND.CODLOGR, vwLOGRADOURO.NOMETIPOLOG, vwLOGRADOURO.NOMETITLOG, "
        sql = sql & "vwLOGRADOURO.NomeLogradouro , CPPROCESSOEND.Numero FROM CPPROCESSOEND INNER JOIN "
        sql = sql & "vwLOGRADOURO ON CPPROCESSOEND.CODLOGR = vwLOGRADOURO.CODLOGRADOURO "
        sql = sql & "Where CPPROCESSOEND.ANO = " & AnoProcessoCP & " And CPPROCESSOEND.NUMPROCESSO = " & CodProcessoCP
        Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sNomeLog = Trim$(SubNull(!NomeTipoLog)) & " " & Trim$(SubNull(!NomeTitLog)) & " " & !NomeLogradouro
                grdEnd.AddItem !CodLogr & Chr(9) & sNomeLog & Chr(9) & !Numero
               .MoveNext
            Loop
           .Close
        End With
       
       'CARREGA DOC
       sql = "SELECT CPPROCESSODOC.CODDOC, CPDOCUMENTO.NOME, CPPROCESSODOC.DATA FROM CPPROCESSODOC INNER JOIN "
       sql = sql & "CPDOCUMENTO ON CPPROCESSODOC.CODDOC = CPDOCUMENTO.CODIGO Where CPPROCESSODOC.ANO = " & AnoProcessoCP & " And CPPROCESSODOC.Numero = " & CodProcessoCP
       Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            If .RowCount = 0 Then GoTo DOC2
            Do Until .EOF
                Set itmX = lvDoc.ListItems.Add(, "C" & Format(!CODDOC, "000"), !nome)
                If Not IsNull(!Data) Then
                   itmX.SubItems(1) = !Data
                   lvDoc.ListItems(lvDoc.ListItems.Count).Checked = True
                   nSim = nSim + 1
                Else
                   nNao = nNao + 1
                End If
               .MoveNext
            Loop
       End With
       lblDoc1.Caption = nSim: lblDoc2.Caption = nNao
    Else
        MsgBox "Processo não cadastrado.", vbExclamation, "Atenção"
    End If
   .Close
   GoTo fim
   
DOC2:
    sql = "SELECT CPASSUNTODOC.CODDOC, CPDOCUMENTO.NOME FROM CPDOCUMENTO INNER JOIN "
    sql = sql & "CPASSUNTODOC ON CPDOCUMENTO.CODIGO = CPASSUNTODOC.CODDOC INNER JOIN CPASSUNTO ON "
    sql = sql & "CPASSUNTODOC.CODASSUNTO = CPASSUNTO.CODIGO Where CPASSUNTO.Codigo = " & cmbAssunto.ItemData(cmbAssunto.ListIndex)
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Set itmX = lvDoc.ListItems.Add(, "C" & Format(!CODDOC, "000"), !nome)
             nNao = nNao + 1
           .MoveNext
        Loop
       .Close
    End With
    lblDoc1.Caption = nSim: lblDoc2.Caption = nNao
   
End With

fim:
CodProcessoCP = 0
bExec = True
End Sub

Private Sub CarregaCombo(bAtivo As Boolean)

cmbAssunto.Clear
sql = "SELECT CODIGO,NOME FROM CPASSUNTO"
If bAtivo Then
    sql = sql & " WHERE ATIVO=1"
End If
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbAssunto.AddItem !nome
       cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With

cmbOrigem.Clear
sql = "SELECT CODIGO,NOME FROM CPORIGEM"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbOrigem.AddItem !nome
       cmbOrigem.ItemData(cmbOrigem.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With
cmbReq.Clear
sql = "SELECT CODIGO,DESCRICAO FROM CPCENTROCUSTO"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbReq.AddItem !Descricao
       cmbReq.ItemData(cmbReq.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With
cmbReq.Text = "SETOR DE DíVIDA ATIVA"

End Sub

Private Sub Limpa()
On Error Resume Next
lblNumProc.Caption = ""
lblAno.Caption = ""
cmbOrigem.ListIndex = -1
cmbAssunto.ListIndex = -1
chkFisico.Value = 0
chkInterno.Value = 1
txtInsc.Text = ""
txtCompl.Text = ""
txtCidadao.Text = ""
lblNomeCidadao.Caption = ""
lblAtendente.Caption = ""
lblDtArquivamento.Caption = "  /  /    "
lblDtCancelamento.Caption = "  /  /    "
lblDtEntrada.Caption = "  /  /    "
lblDtReativacao.Caption = "  /  /    "
lblDtSuspencao.Caption = "  /  /    "
txtObs.Text = ""
grdEnd.Rows = 1
lblCodCid.Caption = ""
lblNomeCid.Caption = ""
pnlDoc.Visible = False
pnlEnd.Visible = False
lblDoc1.Caption = 0
lblDoc2.Caption = 0
z = SendMessage(lvDoc.hwnd, LVM_DELETEALLITEMS, 0, 0)
lblHora.Caption = ""
End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdExcluir.Visible = True
    cmdSair.Visible = True
    cmdConsultar.Visible = True
    cmdTramite.Visible = True
    cmdImprimir.Visible = True
    cmdGravar.Visible = False
    cmdCancel.Visible = False
    If bArq Then
        cmdArquivar.Enabled = True
    End If
    If bCan Then
        cmdCancelar.Enabled = True
    End If
    If bSus Then
        cmdSuspender.Enabled = True
    End If
    If bRea Then
        cmdReativar.Enabled = True
    End If
    If bAne Then
        cmdAnexar.Enabled = True
    End If
    cmbReq.BackColor = Kde
    cmbOrigem.BackColor = Kde
    cmbAssunto.BackColor = Kde
    chkFisico.BackColor = Kde
    chkInterno.BackColor = Kde
    txtCompl.BackColor = Kde
    txtCidadao.BackColor = Kde
    txtObs.BackColor = Kde
    cmbReq.Enabled = False
    cmbOrigem.Enabled = False
    cmbAssunto.Enabled = False
    chkFisico.Enabled = False
    chkInterno.Enabled = False
    txtCompl.Locked = True
    txtObs.Locked = True
    txtCidadao.Locked = True
    cmdNovo2.Enabled = False
    cmdAlterar2.Enabled = False
    cmdExcluir2.Enabled = False
    cmdEditCid.Enabled = True
    cmdEditEnd.Enabled = True
    'CarregaCombo False
    cmdOA.Enabled = True
    cmdOC.Enabled = True
    cmdOR.Enabled = True
    cmdOS.Enabled = True
ElseIf Tipo = "INCLUIR" Then
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdExcluir.Visible = False
    cmdSair.Visible = False
    cmdConsultar.Visible = False
    cmdTramite.Visible = False
    cmdImprimir.Visible = False
    cmdGravar.Visible = True
    cmdCancel.Visible = True
   
    If bArq Then
        cmdArquivar.Enabled = False
    End If
    If bCan Then
        cmdCancelar.Enabled = False
    End If
    If bSus Then
        cmdSuspender.Enabled = False
    End If
    If bRea Then
        cmdReativar.Enabled = False
    End If
    If bAne Then
        cmdAnexar.Enabled = False
    End If
    cmbReq.BackColor = Branco
    
'    cmbReq.Enabled = True
    If Evento = "Novo" Then
        cmbAssunto.BackColor = Branco
        chkFisico.BackColor = Kde
        chkInterno.BackColor = Kde
        cmbAssunto.Enabled = True
        '
        chkInterno.Enabled = True
        txtCompl.BackColor = Branco
        txtCidadao.BackColor = Branco
        txtCompl.Locked = False
        txtObs.BackColor = Branco
        txtObs.Locked = False
        txtCidadao.Locked = False
        'cmdEditCid.Enabled = True
        cmdEditEnd.Enabled = True
    Else
        txtCompl.Locked = True
        txtCompl.BackColor = Kde
        If chkInterno.Value = False Then
            txtObs.Locked = True
            txtObs.BackColor = Kde
        Else
            txtObs.Locked = False
            txtObs.BackColor = Branco
        End If
        'cmdEditCid.Enabled = False
        cmdEditEnd.Enabled = False
    End If
    cmdEditCid.Enabled = True
    cmdEditDoc.Enabled = True
    
    cmdEditEnd.Enabled = True
    chkFisico.Enabled = True
    cmdNovo2.Enabled = True
    cmdAlterar2.Enabled = True
    cmdExcluir2.Enabled = True
    cmdOA.Enabled = False
    cmdOC.Enabled = False
    cmdOR.Enabled = False
    cmdOS.Enabled = False

End If
If cmdAlterar.Enabled = False Then
    If NomeDeLogin = "LUIZH" Or NomeDeLogin = "ROSANGELA" Or NomeDeLogin = "MARTA" Or NomeDeLogin = "RITA" Or NomeDeLogin = "DANIELAR" Or NomeDeLogin = "ANAP" Or NomeDeLogin = "MORITA" Or NomeDeLogin = "PATRICIAG" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "EDUARDO" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "PAULA" Or NomeDeLogin = "NOELI" Or NomeDeLogin = "LORAINE" Or NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "MARILIA" Then
        cmdEditInsc.Enabled = True
    End If
End If

End Sub

Private Sub Eventos2(Tipo As String)

If Tipo = "INICIAR" Then
    cmdNovo2.Visible = True
    cmdAlterar2.Visible = True
    cmdExcluir2.Visible = True
    cmdGravar2.Visible = False
    cmdCancel2.Visible = False
    txtNumeroLog.BackColor = Tzahov
    txtNumLog.BackColor = Tzahov
    txtNomeLog.BackColor = Tzahov
    txtNumeroLog.Locked = True
    txtNumLog.Locked = True
    txtNomeLog.Locked = True
    grdEnd2.Enabled = True
    If cmdNovo.Enabled = True Then
        cmdEditEnd.Enabled = True
    End If
ElseIf Tipo = "INCLUIR" Then
   cmdNovo2.Visible = False
   cmdAlterar2.Visible = False
   cmdExcluir2.Visible = False
   cmdGravar2.Visible = True
   cmdCancel2.Visible = True
   txtNumeroLog.BackColor = Branco
   txtNumLog.BackColor = Branco
   txtNomeLog.BackColor = Branco
   txtNumeroLog.Locked = False
   txtNumLog.Locked = False
   txtNomeLog.Locked = False
   grdEnd2.Enabled = False
   cmdEditEnd.Enabled = False
End If

End Sub

Private Sub Grava()
Dim MaxCod As Long, p As Integer, nCodCidadao As Long, nCodReq As Integer, MinCod As Long, sDoc As String

If Evento = "Novo" Then
    sql = "SELECT MAX(NUMERO) AS MAXIMO FROM CPPROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = !MAXIMO + 1
        End If
       .Close
    End With
Else
    MaxCod = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
End If

'nCodCidadao = Val(lblCodCid.Caption)
nCodCidadao = Val(txtCidadao.Text)
If cmbReq.ListIndex > -1 Then
    nCodReq = cmbReq.ItemData(cmbReq.ListIndex)
'    nCodCidadao = 0
End If

If Evento = "Novo" Then
    sql = "INSERT CPPROCESSOGTI(ANO,NUMERO,FISICO,ORIGEM,INTERNO,CODASSUNTO,COMPLEMENTO,"
    sql = sql & "OBSERVACAO,DATAENTRADA,DATAREATIVA,DATACANCEL,DATAARQUIVA,DATASUSPENSO,"
    sql = sql & "ETIQUETA,CODCIDADAO,RESPONSAVEL,MOTIVOCANCEL,CENTROCUSTO,HORA,INSC) VALUES(" & Val(lblAno.Caption) & ","
    sql = sql & MaxCod & "," & chkFisico.Value & "," & cmbOrigem.ItemData(cmbOrigem.ListIndex) & ","
    sql = sql & chkInterno.Value & "," & cmbAssunto.ItemData(cmbAssunto.ListIndex) & ",'" & Mask(txtCompl.Text) & "','"
    sql = sql & Mask(txtObs.Text) & "','" & Format(lblDtEntrada.Caption, "mm/dd/yyyy") & "'," & "Null" & ","
    sql = sql & "Null" & "," & "Null" & "," & "Null" & "," & "0" & "," & nCodCidadao & ",'" & lblAtendente.Caption & "',"
    sql = sql & "Null" & "," & nCodReq & ",'" & lblHora.Caption & "'," & 0 & ")"
    lblNumProc.Caption = CStr(MaxCod) & RetornaDVProcesso(MaxCod)
    lblNumProc.Caption = Format(lblNumProc.Caption, "000000-0")
Else
    sql = "UPDATE CPPROCESSOGTI SET FISICO=" & chkFisico.Value & ",ORIGEM=" & cmbOrigem.ItemData(cmbOrigem.ListIndex) & ","
    sql = sql & "INTERNO=" & chkInterno.Value & ",CODASSUNTO=" & cmbAssunto.ItemData(cmbAssunto.ListIndex) & ","
    sql = sql & "COMPLEMENTO='" & Mask(txtCompl.Text) & "',OBSERVACAO='" & Mask(txtObs.Text) & "',INSC=" & 0 & ","
    sql = sql & "CODCIDADAO=" & nCodCidadao & ",CENTROCUSTO=" & nCodReq
    sql = sql & " WHERE NUMERO=" & MaxCod & " AND ANO=" & Val(lblAno.Caption)
End If
cn.Execute sql, rdExecDirect

sql = "DELETE FROM CPPROCESSOEND WHERE ANO=" & Val(lblAno.Caption) & " AND NUMPROCESSO=" & MaxCod
cn.Execute sql, rdExecDirect

With grdEnd
    For p = 1 To .Rows - 1
        sql = "INSERT CPPROCESSOEND (ANO,NUMPROCESSO,CODLOGR,NUMERO) VALUES("
        sql = sql & Val(lblAno.Caption) & "," & MaxCod & "," & .TextMatrix(p, 0) & ",'" & .TextMatrix(p, 2) & "')"
        cn.Execute sql, rdExecDirect
    Next
End With

sql = "DELETE FROM CPPROCESSODOC WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & MaxCod
cn.Execute sql, rdExecDirect

With lvDoc
    For p = 1 To .ListItems.Count
        sql = "INSERT CPPROCESSODOC (ANO,NUMERO,CODDOC,DATA) VALUES("
        sql = sql & Val(lblAno.Caption) & "," & MaxCod & "," & Val(Right$(.ListItems(p).Key, 3)) & ","
        sql = sql & IIf(.ListItems(p).SubItems(1) = "", "Null", "'" & Format(.ListItems(p).SubItems(1), "mm/dd/yyyy") & "'") & ")"
        cn.Execute sql, rdExecDirect
    Next
End With

'GRAVA CIDADÃO NO PROCESSO A PARTIR DE 01/2011
'If nCodCidadao > 0 Then
'
'    If Evento <> "Novo" Then
'        Sql = "DELETE FROM cpprocessocidadao WHERE ANOPROC=" & Val(lblAno.Caption) & " AND NUMPROC=" & MaxCod
'        cn.Execute Sql, rdExecDirect
'    End If
'    Sql = "SELECT * FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If Not IsNull(!CPF) And SubNull(!CPF) <> "" Then
'
'            sDoc = !CPF
'        Else
'            If Not IsNull(!Cnpj) Then
'                sDoc = !Cnpj
'            Else
'                sDoc = ""
'            End If
'        End If
'        Sql = "INSERT cpprocessocidadao (ANOPROC,NUMPROC,CODCIDADAO,NOMECIDADAO,DOC,CODLOGRADOURO,NUMIMOVEL,COMPLEMENTO,CODBAIRRO,CODCIDADE,SIGLAUF,CEP,RG,ORGAO) VALUES("
'        Sql = Sql & Val(lblAno.Caption) & "," & MaxCod & "," & nCodCidadao & ",'" & Mask(!NOMECIDADAO) & "','" & sDoc & "'," & Val(SubNull(!CODLOGRADOURO)) & ","
'        Sql = Sql & Val(SubNull(!NUMIMOVEL)) & ",'" & !COMPLEMENTO & "'," & Val(SubNull(!codbairro)) & "," & !CODCIDADE & ",'" & SubNull(!siglauf) & "'," & Val(SubNull(!CEP)) & ",'"
'        Sql = Sql & SubNull(!rg) & "','" & SubNull(!ORGAO) & "')"
'        cn.Execute Sql, rdExecDirect
'    End With
'
'End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmCnsProcesso2
End Sub

Private Sub grdEnd2_Click()
txtNumLog.Text = ""
txtNomeLog.Text = ""
txtNumeroLog.Text = ""
If grdEnd2.Rows = 1 Then Exit Sub
If grdEnd2.Row > 0 Then
    txtNumLog.Text = grdEnd2.TextMatrix(grdEnd2.Row, 0)
    txtNumLog_LostFocus
    txtNumeroLog.Text = grdEnd2.TextMatrix(grdEnd2.Row, 2)
End If

End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub lvDoc_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Evento = "" Then
    Item.Checked = Not Item.Checked
    Exit Sub
End If

If Item.Checked = True Then
   Item.SubItems(1) = Format(Now, "dd/mm/yyyy")
Else
   Item.SubItems(1) = ""
End If
End Sub

Private Sub txtCidadao_Change()
lblNomeCidadao.Caption = ""
End Sub

Private Sub txtCidadao_KeyPress(KeyAscii As Integer)
Tweak txtCidadao, KeyAscii, IntegerPositive
End Sub

Private Sub txtCidadao_LostFocus()

If Val(txtCidadao.Text) > 0 Then
    lblNomeCidadao.Caption = RetornaNome(Val(txtCidadao.Text))
End If
End Sub

Private Sub txtNumLog_KeyPress(KeyAscii As Integer)
Tweak txtNumLog, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumLog_LostFocus()
If Val(txtNumLog.Text) > 0 Then
   sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   sql = sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   sql = sql & "ABREVTITLOG,NOMELOGRADOURO "
   sql = sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtNumLog.Text)
   Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
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
Else
    txtNomeLog.Text = ""
End If

End Sub

Private Sub txtNomeLog_Change()
If Trim$(txtNomeLog) = "" Then
   txtNumLog.Text = 0
End If
End Sub

Private Sub txtNomeLog_GotFocus()
txtNomeLog.SelStart = 0
txtNomeLog.SelLength = Len(txtNomeLog)
End Sub

Private Sub txtNomeLog_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLog.Text <> "" Then
      sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      sql = sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      sql = sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      sql = sql & "NUMOFIC FROM vwLOGRADOURO "
      sql = sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLog) & "%' "
      sql = sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
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

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
   txtNumLog.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtNumLog_LostFocus
   lstNomeLog.Visible = False
   txtNum.SetFocus
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtNumLog.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtNumLog_LostFocus
       lstNomeLog.Visible = False
       txtNumeroLog.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
End If

End Sub

Private Sub DesabilitaPainelPrincipal()

cmdNovo.Enabled = False
cmdAlterar.Enabled = False
cmdExcluir.Enabled = False
cmdGravar.Enabled = False
cmdCancel.Enabled = False
cmdConsultar.Enabled = False
cmdSair.Enabled = False
cmdTramite.Enabled = False
cmdImprimir.Enabled = False
cmdArquivar.Enabled = False
cmdCancelar.Enabled = False
cmdSuspender.Enabled = False
cmdReativar.Enabled = False

cmdOA.Enabled = False
cmdOC.Enabled = False
cmdOR.Enabled = False
cmdOS.Enabled = False
chkInterno.Enabled = False: chkFisico.Enabled = False: cmbOrigem.Enabled = False: txtCompl.Locked = True: txtObs.Locked = True: cmbAssunto.Enabled = False
chkInterno.BackColor = Kde: chkFisico.BackColor = Kde: cmbOrigem.BackColor = Kde: txtCompl.BackColor = Kde: txtObs.BackColor = Kde: cmbAssunto.BackColor = Kde
End Sub

Public Sub HabilitaPainelPrincipal()
On Error Resume Next

cmdGravar.Enabled = True
cmdCancel.Enabled = True
cmdConsultar.Enabled = True
cmdSair.Enabled = True
cmdTramite.Enabled = True

If bArq Then
    cmdArquivar.Enabled = True
End If
If bCan Then
    cmdCancelar.Enabled = True
End If
If bSus Then
    cmdSuspender.Enabled = True
End If
If bRea Then
    cmdReativar.Enabled = True
End If
If bAne Then
    cmdAnexar.Enabled = True
End If
cmdOA.Enabled = True
cmdOC.Enabled = True
cmdOR.Enabled = True
cmdOS.Enabled = True
If Evento = "" Then
    cmdImprimir.Enabled = True
    cmdNovo.Enabled = True
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdArquivar.Enabled = True
    cmdCancelar.Enabled = True
    cmdReativar.Enabled = True
    cmdSuspender.Enabled = True
    cmdAnexar.Enabled = True
    cmdRepair.Enabled = True
    cmdArquivos.Enabled = True
End If
If Evento <> "" Then
    chkInterno.Enabled = True: chkFisico.Enabled = True: cmbOrigem.Enabled = True: txtCompl.Locked = False: txtObs.Locked = False: cmbAssunto.Enabled = True
    cmbOrigem.BackColor = Branco: txtCompl.BackColor = Branco: txtObs.BackColor = Branco: cmbAssunto.BackColor = Branco
End If

End Sub

Private Function VerificaTramite() As Boolean

Dim aSeq() As Integer

'CARREGA TODOS OS TRAMITES
ReDim aSeq(0)
sql = "SELECT * FROM CPTRAMITACAOcc Where ano = " & lblAno.Caption & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        sql = "SELECT CPASSUNTOCC.SEQ,CPCENTROCUSTO.CODIGO, CPCENTROCUSTO.DESCRICAO FROM CPASSUNTOCC INNER JOIN "
        sql = sql & "CPCENTROCUSTO ON CPASSUNTOCC.CODCC = CPCENTROCUSTO.CODIGO "
        sql = sql & "WHERE CPASSUNTOCC.CODASSUNTO =" & frmCPProcesso.cmbAssunto.ItemData(frmCPProcesso.cmbAssunto.ListIndex)
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aSeq(UBound(aSeq) + 1)
                aSeq(UBound(aSeq)) = !Seq
               .MoveNext
            Loop
           .Close
        End With
    Else
        sql = "SELECT CPTRAMITACAOcc.seq, CPTRAMITACAOcc.ccusto, CPCENTROCUSTO.DESCRICAO "
        sql = sql & "FROM CPTRAMITACAOcc INNER JOIN CPCENTROCUSTO ON CPTRAMITACAOcc.ccusto = CPCENTROCUSTO.CODIGO "
        sql = sql & "Where CPTRAMITACAOcc.ano = " & lblAno.Caption & " And CPTRAMITACAOcc.Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        sql = sql & " order by CPTRAMITACAOCC.SEQ"
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aSeq(UBound(aSeq) + 1)
                aSeq(UBound(aSeq)) = !Seq
               .MoveNext
            Loop
           .Close
        End With
    End If
   .Close
End With

VerificaTramite = True
'VERIFICA OS TRAMITES CONCLUIDOS
If Val(frmCPProcesso.lblNumProc.Caption) > 0 Then
    For x = 1 To UBound(aSeq)
        sql = "SELECT CCUSTO,DESCRICAO,DATAHORA,NOMECOMPLETO,DESCDESPACHO FROM VWCPTRAMITACAO WHERE ANO=" & lblAno.Caption
        sql = sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        sql = sql & " AND SEQ=" & aSeq(x)
        Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If IsNull(!DATAHORA) Then
                    VerificaTramite = False
                    Exit Function
                End If
            Else
                    VerificaTramite = False
                    Exit Function
            End If
           .Close
        End With
    Next
End If

End Function

