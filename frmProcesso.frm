VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProcesso 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle Geral dos Processos"
   ClientHeight    =   6075
   ClientLeft      =   10830
   ClientTop       =   4200
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   8805
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
      Height          =   2385
      Left            =   1710
      TabIndex        =   68
      Top             =   1230
      Visible         =   0   'False
      Width           =   4485
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
         TabIndex        =   73
         Top             =   360
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkP2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Requerimento de abertura de processo"
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
         TabIndex        =   72
         Top             =   690
         Width           =   3735
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
         TabIndex        =   71
         Top             =   1020
         Width           =   3765
      End
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
         TabIndex        =   70
         Top             =   1365
         Width           =   3765
      End
      Begin VB.CheckBox chkP5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Requerimento de Cancelamento de processo"
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
         TabIndex        =   69
         Top             =   1695
         Width           =   4140
      End
      Begin prjChameleon.chameleonButton cmdCancelPrint 
         Height          =   345
         Left            =   3960
         TabIndex        =   75
         ToolTipText     =   "Cancelar operação"
         Top             =   1980
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
         MICON           =   "frmProcesso.frx":0000
         PICN            =   "frmProcesso.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOrgão 
         Height          =   255
         Left            =   3825
         TabIndex        =   117
         ToolTipText     =   "Trocar órgão por CPF"
         Top             =   690
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   450
         BTYPE           =   14
         TX              =   "<>"
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
         MCOL            =   8421504
         MPTR            =   1
         MICON           =   "frmProcesso.frx":0176
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOKPrint 
         Height          =   345
         Left            =   3540
         TabIndex        =   74
         ToolTipText     =   "Imprimir os documentos selecionados"
         Top             =   1980
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
         MICON           =   "frmProcesso.frx":0192
         PICN            =   "frmProcesso.frx":01AE
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
      Left            =   30
      TabIndex        =   36
      Top             =   4410
      Width           =   7350
      Begin prjChameleon.chameleonButton cmdEditDoc 
         Height          =   315
         Left            =   6030
         TabIndex        =   37
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
         MICON           =   "frmProcesso.frx":0308
         PICN            =   "frmProcesso.frx":0324
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         TabIndex        =   41
         Top             =   270
         Width           =   285
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
         TabIndex        =   40
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Pendentes..:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2670
         TabIndex        =   39
         Top             =   270
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Entregues..:"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   38
         Top             =   270
         Width           =   1845
      End
   End
   Begin VB.Frame frObs 
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
      Left            =   45
      TabIndex        =   15
      Top             =   2460
      Width           =   7305
      Begin VB.TextBox txtInsc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5940
         MaxLength       =   6
         TabIndex        =   5
         Text            =   "112333"
         Top             =   135
         Width           =   780
      End
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   90
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   450
         Width           =   7110
      End
      Begin prjChameleon.chameleonButton cmdEditInsc 
         Height          =   270
         Left            =   6795
         TabIndex        =   98
         ToolTipText     =   "Alterar nº de inscrição"
         Top             =   135
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   476
         BTYPE           =   14
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
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcesso.frx":047E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdExpandirObs 
         Height          =   225
         Left            =   1755
         TabIndex        =   118
         ToolTipText     =   "Expandir ou recolher o campo de observação"
         Top             =   135
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   397
         BTYPE           =   7
         TX              =   "Aumentar/Diminuir"
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
         FCOL            =   16384
         FCOLO           =   16384
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcesso.frx":049A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Inscrição..:"
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
         Index           =   1
         Left            =   4185
         TabIndex        =   97
         Top             =   180
         Width           =   1665
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5970
      Left            =   7380
      TabIndex        =   52
      Top             =   90
      Width           =   1365
      Begin prjChameleon.chameleonButton cmdArquivos 
         Height          =   315
         Left            =   90
         TabIndex        =   96
         ToolTipText     =   "Conteúdo digital do Processo"
         Top             =   4140
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Arqui&vos"
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
         MICON           =   "frmProcesso.frx":04B6
         PICN            =   "frmProcesso.frx":04D2
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
         Left            =   90
         TabIndex        =   53
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
         MCOL            =   32768
         MPTR            =   1
         MICON           =   "frmProcesso.frx":0599
         PICN            =   "frmProcesso.frx":05B5
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
         TabIndex        =   54
         ToolTipText     =   "Sair da Tela"
         Top             =   5445
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
         MICON           =   "frmProcesso.frx":0657
         PICN            =   "frmProcesso.frx":0673
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
         Left            =   90
         TabIndex        =   55
         ToolTipText     =   "Consulta Processos Cadastrados"
         Top             =   5085
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
         MICON           =   "frmProcesso.frx":06E1
         PICN            =   "frmProcesso.frx":06FD
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
         TabIndex        =   56
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
         MICON           =   "frmProcesso.frx":0857
         PICN            =   "frmProcesso.frx":0873
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
         TabIndex        =   57
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
         MICON           =   "frmProcesso.frx":09CD
         PICN            =   "frmProcesso.frx":09E9
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
         TabIndex        =   58
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
         MICON           =   "frmProcesso.frx":0B43
         PICN            =   "frmProcesso.frx":0B5F
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
         TabIndex        =   59
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
         MICON           =   "frmProcesso.frx":0CB9
         PICN            =   "frmProcesso.frx":0CD5
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
         TabIndex        =   60
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcesso.frx":0E2F
         PICN            =   "frmProcesso.frx":0E4B
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
         TabIndex        =   61
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmProcesso.frx":0EEB
         PICN            =   "frmProcesso.frx":0F07
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
         TabIndex        =   62
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmProcesso.frx":0FA6
         PICN            =   "frmProcesso.frx":0FC2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAnexar 
         Height          =   315
         Left            =   90
         TabIndex        =   63
         ToolTipText     =   "Anexar um Processo"
         Top             =   3420
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "A&nexos"
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
         MICON           =   "frmProcesso.frx":1036
         PICN            =   "frmProcesso.frx":1052
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
         TabIndex        =   64
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
         MICON           =   "frmProcesso.frx":11AC
         PICN            =   "frmProcesso.frx":11C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRepair 
         Height          =   315
         Left            =   90
         TabIndex        =   65
         ToolTipText     =   "Anexar um Processo"
         Top             =   3780
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Corrigir"
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
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmProcesso.frx":1322
         PICN            =   "frmProcesso.frx":133E
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
         TabIndex        =   66
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
         MICON           =   "frmProcesso.frx":13C1
         PICN            =   "frmProcesso.frx":13DD
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
         TabIndex        =   67
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
         MICON           =   "frmProcesso.frx":1782
         PICN            =   "frmProcesso.frx":179E
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
      Left            =   30
      TabIndex        =   26
      Top             =   15
      Width           =   7305
      Begin VB.ComboBox cmbAssunto 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1530
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   900
         Width           =   5625
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1530
         MaxLength       =   150
         TabIndex        =   4
         Top             =   1260
         Width           =   5595
      End
      Begin VB.ComboBox cmbOrigem 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4590
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   540
         Width           =   2565
      End
      Begin VB.CheckBox chkFisico 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Processo Físico..:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.CheckBox chkInterno 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Processo Interno..:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1980
         TabIndex        =   1
         Top             =   600
         Width           =   1635
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
         TabIndex        =   51
         Top             =   300
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora:"
         Height          =   225
         Index           =   14
         Left            =   6150
         TabIndex        =   50
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo...:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   35
         Top             =   300
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano..:"
         Height          =   225
         Index           =   1
         Left            =   2460
         TabIndex        =   34
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Origem....:"
         Height          =   225
         Index           =   3
         Left            =   3780
         TabIndex        =   33
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Atendente..:"
         Height          =   225
         Index           =   5
         Left            =   3510
         TabIndex        =   32
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto...............:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento......:"
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
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
         TabIndex        =   29
         Top             =   300
         Width           =   855
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
         TabIndex        =   28
         Top             =   300
         Width           =   585
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
         TabIndex        =   27
         Top             =   300
         Width           =   1695
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
      Left            =   30
      TabIndex        =   13
      Top             =   4995
      Width           =   7350
      Begin MSFlexGridLib.MSFlexGrid grdEnd 
         Height          =   765
         Left            =   90
         TabIndex        =   8
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
         TabIndex        =   9
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
         MICON           =   "frmProcesso.frx":18F8
         PICN            =   "frmProcesso.frx":1914
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
      Left            =   30
      TabIndex        =   14
      Top             =   1650
      Width           =   7305
      Begin prjChameleon.chameleonButton cmdOC 
         Height          =   225
         Left            =   4590
         TabIndex        =   44
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
         MICON           =   "frmProcesso.frx":1A6E
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
         TabIndex        =   45
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
         MICON           =   "frmProcesso.frx":1A8A
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
         TabIndex        =   46
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
         MICON           =   "frmProcesso.frx":1AA6
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
         TabIndex        =   47
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
         MICON           =   "frmProcesso.frx":1AC2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAnexo 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   6060
         TabIndex        =   49
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexos............:"
         Height          =   225
         Index           =   13
         Left            =   4905
         TabIndex        =   48
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada...........:"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reativação.....:"
         Height          =   225
         Index           =   8
         Left            =   2490
         TabIndex        =   24
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelamento.:"
         Height          =   225
         Index           =   9
         Left            =   2460
         TabIndex        =   23
         Top             =   540
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivamento..:"
         Height          =   225
         Index           =   10
         Left            =   4905
         TabIndex        =   22
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Suspenção.....:"
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblDtEntrada 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1245
         TabIndex        =   20
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblDtSuspencao 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1245
         TabIndex        =   19
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDtArquivamento 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   6045
         TabIndex        =   18
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label lblDtReativacao 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3645
         TabIndex        =   17
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblDtCancelamento 
         BackStyle       =   0  'Transparent
         Caption         =   "12/01/2005"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3645
         TabIndex        =   16
         Top             =   540
         Width           =   1005
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
      Height          =   750
      Left            =   30
      TabIndex        =   10
      Top             =   3645
      Width           =   7305
      Begin VB.OptionButton OptEnd 
         Caption         =   "Com."
         Height          =   195
         Index           =   1
         Left            =   6480
         TabIndex        =   120
         Top             =   495
         Width           =   690
      End
      Begin VB.OptionButton OptEnd 
         Caption         =   "Res."
         Height          =   195
         Index           =   0
         Left            =   5805
         TabIndex        =   119
         Top             =   495
         Value           =   -1  'True
         Width           =   645
      End
      Begin prjChameleon.chameleonButton cmdEditCid 
         Height          =   270
         Left            =   5760
         TabIndex        =   7
         ToolTipText     =   "Editar requerente do processo"
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   476
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
         MICON           =   "frmProcesso.frx":1ADE
         PICN            =   "frmProcesso.frx":1AFA
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
         Height          =   270
         Left            =   6885
         TabIndex        =   99
         ToolTipText     =   "Gerar guia"
         Top             =   180
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   476
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
         MICON           =   "frmProcesso.frx":1C54
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
         Height          =   270
         Left            =   6570
         TabIndex        =   100
         ToolTipText     =   "Exibir cidadão gravado no processo original"
         Top             =   180
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   476
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
         MICON           =   "frmProcesso.frx":1C70
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodCid 
         BackStyle       =   0  'Transparent
         Caption         =   "523888"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   300
         Width           =   675
      End
      Begin VB.Label lblNomeCid 
         BackStyle       =   0  'Transparent
         Caption         =   "MARCELA DE SOUZA BRITO CARVALHO"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   780
         TabIndex        =   11
         Top             =   300
         Width           =   4755
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
      Left            =   660
      TabIndex        =   76
      Top             =   1830
      Visible         =   0   'False
      Width           =   6525
      Begin VB.TextBox txtObsData 
         Appearance      =   0  'Flat
         Height          =   1905
         Left            =   90
         MaxLength       =   250
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   77
         Top             =   360
         Width           =   6315
      End
      Begin prjChameleon.chameleonButton cmdGravarObs 
         Height          =   345
         Left            =   5580
         TabIndex        =   78
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
         MICON           =   "frmProcesso.frx":1C8C
         PICN            =   "frmProcesso.frx":1CA8
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
         TabIndex        =   79
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
         MICON           =   "frmProcesso.frx":1E02
         PICN            =   "frmProcesso.frx":1E1E
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
         TabIndex        =   95
         Top             =   2400
         Width           =   4365
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
      Left            =   1050
      TabIndex        =   80
      Top             =   1980
      Width           =   6165
      Begin MSComctlLib.ListView lvDoc 
         Height          =   2055
         Left            =   90
         TabIndex        =   81
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
      Left            =   990
      TabIndex        =   82
      Top             =   3120
      Width           =   6225
      Begin VB.TextBox txtNumLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         TabIndex        =   87
         Top             =   1425
         Width           =   855
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   86
         Top             =   1425
         Width           =   3990
      End
      Begin VB.TextBox txtNumeroLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         MaxLength       =   15
         TabIndex        =   85
         Top             =   1740
         Width           =   1305
      End
      Begin VB.ListBox lstNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         ItemData        =   "frmProcesso.frx":1F78
         Left            =   2040
         List            =   "frmProcesso.frx":1F7A
         TabIndex        =   83
         Top             =   585
         Visible         =   0   'False
         Width           =   3990
      End
      Begin prjChameleon.chameleonButton cmdAlterar2 
         Height          =   315
         Left            =   4260
         TabIndex        =   84
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
         MICON           =   "frmProcesso.frx":1F7C
         PICN            =   "frmProcesso.frx":1F98
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
         TabIndex        =   88
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
         TabIndex        =   89
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
         MICON           =   "frmProcesso.frx":20F2
         PICN            =   "frmProcesso.frx":210E
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
         TabIndex        =   90
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
         MICON           =   "frmProcesso.frx":2268
         PICN            =   "frmProcesso.frx":2284
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
         TabIndex        =   93
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
         MICON           =   "frmProcesso.frx":23DE
         PICN            =   "frmProcesso.frx":23FA
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
         TabIndex        =   94
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
         MICON           =   "frmProcesso.frx":249C
         PICN            =   "frmProcesso.frx":24B8
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
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   92
         Top             =   1455
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número.......:"
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   91
         Top             =   1770
         Width           =   975
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
      Height          =   750
      Left            =   15
      TabIndex        =   42
      Top             =   3675
      Width           =   7350
      Begin VB.ComboBox cmbReq 
         Height          =   315
         Left            =   135
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   210
         Width           =   7035
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
      Left            =   540
      TabIndex        =   101
      Top             =   2340
      Visible         =   0   'False
      Width           =   6495
      Begin prjChameleon.chameleonButton cmdFecharCO 
         Height          =   315
         Left            =   5400
         TabIndex        =   104
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
         MICON           =   "frmProcesso.frx":285D
         PICN            =   "frmProcesso.frx":2879
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCOCidade 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4005
         TabIndex        =   116
         Top             =   1665
         Width           =   2355
      End
      Begin VB.Label lblCOBairro 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   115
         Top             =   1665
         Width           =   2130
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade..:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   3285
         TabIndex        =   114
         Top             =   1665
         Width           =   735
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro.........:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   113
         Top             =   1665
         Width           =   915
      End
      Begin VB.Label lblCOCompl 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   112
         Top             =   1350
         Width           =   5280
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Complem....:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   111
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label lblCOEnd 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   110
         Top             =   1035
         Width           =   5280
      End
      Begin VB.Label lblCORG 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3330
         TabIndex        =   109
         Top             =   720
         Width           =   3075
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "R.G..:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2790
         TabIndex        =   108
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblCODoc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   107
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   106
         Top             =   1035
         Width           =   915
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ.:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   105
         Top             =   720
         Width           =   960
      End
      Begin VB.Label lblCONome 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1080
         TabIndex        =   103
         Top             =   405
         Width           =   5280
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome.........:"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   135
         TabIndex        =   102
         Top             =   405
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim Evento As String, Evento2 As String, bExec As Boolean, sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer, evEsp As Integer, evImp As Integer, evAne As Integer, evCan As Integer, evRea As Integer, evSus As Integer, evArq As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean, bEsp As Boolean, bImp As Boolean, bAne As Boolean, bCan As Boolean, bRea As Boolean, bSus As Boolean, bArq As Boolean
Dim bevNew As Boolean, bevEdit As Boolean, bevDel As Boolean, bEvPrint As Boolean
Dim bEvArquivar As Boolean, bEvCancdelar As Boolean, bEvSuspender As Boolean, bEvReativar As Boolean, bEvAnexos As Boolean
Dim bEvCorrigir As Boolean, bEvArquivos As Boolean, nLoginId As Integer

Private Sub chkInterno_Click()
If chkInterno.value = 1 Then
    cmbOrigem.ListIndex = 0
Else
    cmbOrigem.ListIndex = 1
End If
End Sub

Private Sub cmbAssunto_Click()
If cmbAssunto.ListIndex = -1 Then Exit Sub
'If txtCompl.text = "" Then
    txtCompl.Text = cmbAssunto.Text
'End If

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

'If IsDate(lblDtArquivamento.Caption) Then
'   MsgBox "Processo Arquivado.", vbExclamation, "Atenção"
'   Exit Sub
'End If

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

Private Sub cmdAnexar_Click()
If lblAno.Caption = "" Or lblNumProc = "" Then
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

'frmAnexaProc.show
'frmAnexaProc.ZOrder 0

frmAnexoLog.show vbModal

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
                    Sql = "UPDATE PROCESSOGTI SET DATAARQUIVA='" & Format(z, "mm/dd/yyyy") & "',OBSA='" & Mask(CStr(t)) & "' WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                    cn.Execute Sql, rdExecDirect
                    lblDtArquivamento = z
                 End If
            End If
         End If
     End If
End If

End Sub

Private Sub cmdArquivos_Click()

If NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Em desenvolvimento.", vbInformation, "Atenção"
    Exit Sub
End If

If lblNumProc.Caption <> "" Then
    frmProcessoDigital.show: frmProcessoDigital.ZOrder 0
Else
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
End If

End Sub

Private Sub cmdCancel_Click()

Eventos "INICIAR"
Evento = ""
If lblNumProc.Caption <> "" Then
   CodProcesso = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
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
                Sql = "UPDATE PROCESSOGTI SET DATACANCEL='" & Format(z, "mm/dd/yyyy") & "', OBSC='" & Mask(CStr(t)) & "' WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                cn.Execute Sql, rdExecDirect
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
Dim Sql As String, RdoAux As rdoResultset, nNumproc As Integer, sDoc As String

If Val(lblAno.Caption) = 0 Then Exit Sub
If cmbReq.ListIndex > -1 Then Exit Sub
If frmProcesso.lblNumProc.Caption = "" Then Exit Sub
nNumproc = Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
Sql = "SELECT processocidadao.anoproc, processocidadao.numproc, processocidadao.codcidadao, processocidadao.nomecidadao, processocidadao.doc,processocidadao.RG,processocidadao.ORGAO, "
Sql = Sql & "processocidadao.numimovel, processocidadao.complemento, processocidadao.codbairro, processocidadao.codcidade, processocidadao.siglauf,"
Sql = Sql & "processocidadao.cep, vwLOGRADOURO.abrevtipolog, vwLOGRADOURO.abrevtitlog, vwLOGRADOURO.nomelogradouro, processocidadao.codlogradouro,"
Sql = Sql & "Tributacao.dbo.cidade.desccidade , Tributacao.dbo.bairro.DescBairro FROM processocidadao INNER JOIN vwLOGRADOURO ON processocidadao.codlogradouro = vwLOGRADOURO.codlogradouro LEFT OUTER JOIN "
Sql = Sql & "Tributacao.dbo.bairro ON processocidadao.siglauf = Tributacao.dbo.bairro.siglauf AND processocidadao.codcidade = Tributacao.dbo.bairro.codcidade AND "
Sql = Sql & "processocidadao.codbairro = Tributacao.dbo.bairro.codbairro LEFT OUTER JOIN Tributacao.dbo.cidade ON processocidadao.siglauf = Tributacao.dbo.cidade.siglauf AND processocidadao.codcidade = Tributacao.dbo.cidade.codcidade WHERE ANOPROC=" & Val(lblAno.Caption) & " AND NUMPROC=" & nNumproc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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
frmCnsProcesso2.show
frmCnsProcesso2.ZOrder 0
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
            Sql = "SELECT ASSUNTODOC.CODDOC, DOCUMENTO.NOME FROM DOCUMENTO INNER JOIN "
            Sql = Sql & "ASSUNTODOC ON DOCUMENTO.CODIGO = ASSUNTODOC.CODDOC INNER JOIN ASSUNTO ON "
            Sql = Sql & "ASSUNTODOC.CODASSUNTO = ASSUNTO.CODIGO Where ASSUNTO.Codigo = " & cmbAssunto.ItemData(cmbAssunto.ListIndex)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                    Set itmX = lvDoc.ListItems.Add(, "C" & Format(!CODDOC, "000"), !Nome)
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

Private Sub cmdEditInsc_Click()
Dim z As Variant

If lblAno.Caption = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

z = InputBox("Digite um número de inscrição.", "Alteração", txtInsc.Text)
If Val(z) > 0 And lblNumProc.Caption <> "" Then
    Sql = "UPDATE PROCESSOGTI SET INSC=" & Val(z)
    Sql = Sql & " WHERE  ANO=" & Val(lblAno.Caption)
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    cn.Execute Sql, rdExecDirect
    txtInsc.Text = Val(z)
End If

End Sub

Private Sub cmdExcluir_Click()

If lblAno.Caption = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

Sql = "SELECT ANO,NUMERO,SEQ FROM TRAMITACAO WHERE ANO=" & Val(lblAno.Caption)
Sql = Sql & " AND NUMERO=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        .Close
        MsgBox "Processo possue tramite e não pode ser excluido.", vbExclamation, "Atenção"
        Exit Sub
    End If
   .Close
End With

If MsgBox("Excluir este  Processo ???", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE EXCLUSÃO") = vbNo Then Exit Sub

Sql = "DELETE FROM PROCESSOEND WHERE ANO=" & lblAno.Caption & " AND NUMPROCESSO=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM PROCESSODOC WHERE ANO=" & lblAno.Caption & " AND NUMERO=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM PROCESSOGTI WHERE ANO=" & lblAno.Caption & " AND NUMERO=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
cn.Execute Sql, rdExecDirect
Log Form, Me.Caption, Exclusão, "Excluído processo '" & frmProcesso.lblNumProc.Caption & "'"

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

Private Sub cmdExpandirObs_Click()
frObs.ZOrder 0
If cmdExpandirObs.value = True Then
    frObs.Top = 540
    frObs.Height = 5355
    txtObs.Height = 4700
Else
    frObs.Top = 2460
    frObs.Height = 1215
    txtObs.Height = 675
End If

End Sub

Private Sub cmdFecharCO_Click()
frCidadao.Visible = False
End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset
If bLocal Then
    Exit Sub
End If


If cmbOrigem.ListIndex = -1 Then
    MsgBox "Selecione a origem do processo.", vbExclamation, "Atenção"
    cmbOrigem.SetFocus
    Exit Sub
End If

If cmbAssunto.ListIndex = -1 Then
    MsgBox "Selecione o assunto do processo.", vbExclamation, "Atenção"
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
        Sql = "SELECT * FROM CENTROCUSTO WHERE CODIGO=" & cmbReq.ItemData(cmbReq.ListIndex)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux!Ativo = 0 Then
            MsgBox "Este centro de custos esta desativado e não pode ser utilizado.", vbCritical, "Atenção"
            Exit Sub
        End If
    End If
End If
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
Sql = "UPDATE PROCESSOGTI SET "
Select Case Left(lblOcor.Caption, 1)
        Case "A"
                Sql = Sql & "OBSA='"
        Case "C"
                Sql = Sql & "OBSC='"
        Case "S"
                Sql = Sql & "OBSS='"
        Case "R"
                Sql = Sql & "OBSR='"
End Select
Sql = Sql & Mask(txtObsData.Text) & "' WHERE  ANO=" & Val(lblAno.Caption)
Sql = Sql & " AND NUMERO=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
cn.Execute Sql, rdExecDirect

pnlObs.Visible = False
End Sub

Private Sub cmdGuia_Click()
CodCidadao = lblCodCid.Caption
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
lblAtendente.Tag = nLoginId

Evento = "Novo"
Eventos "INCLUIR"
chkFisico.SetFocus
If cmbOrigem.ListCount > 0 Then cmbOrigem.ListIndex = 1
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
    Sql = "SELECT ANO,NUMERO,OBSA FROM PROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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
    Sql = "SELECT ANO,NUMERO,OBSC FROM PROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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

If chkP1.value = 0 And chkP2.value = 0 And chkP3.value = 0 And chkP4.value = 0 And chkP5.value = 0 Then
    MsgBox "Nenhum relatório selecionado.", vbExclamation, "Atenção"
    Exit Sub
End If

If chkP1.value = 1 Then
    frmReport.ShowReport "PROTOCOLOENTRADA", frmMdi.hwnd, Me.hwnd
End If
If chkP2.value = 1 Then
    frmReport.ShowReport "REQUERIMENTO", frmMdi.hwnd, Me.hwnd
End If
If chkP3.value = 1 Then
    Sql = "SELECT * From PROCESSODOC  Where ANO = " & Val(frmProcesso.lblAno.Caption) & " And Numero = " & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2)) & " And Data Is Null"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
'        If .RowCount = 0 Then
 '           MsgBox "Este processo não possue documentos pendentes.", vbExclamation, "Atenção"
  '      Else
            frmReport.ShowReport "COMUNICADODOC", frmMdi.hwnd, Me.hwnd
   '     End If
       .Close
    End With
End If
If chkP4.value = 1 Then
    Sql = "SELECT * From PROCESSODOC  Where ANO = " & Val(frmProcesso.lblAno.Caption) & " And Numero = " & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2)) & " And Data Is not Null"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "Este processo não possue documentos entregues.", vbExclamation, "Atenção"
        Else
            frmReport.ShowReport "COMPROVANTEDOC", frmMdi.hwnd, Me.hwnd
        End If
       .Close
    End With
End If
If chkP5.value = 1 Then
    frmReport.ShowReport "REQUERIMENTOCANCEL", frmMdi.hwnd, Me.hwnd
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
    Sql = "SELECT ANO,NUMERO,OBSR FROM PROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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
    Sql = "SELECT ANO,NUMERO,OBSS FROM PROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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
                    Sql = "UPDATE PROCESSOGTI SET DATAREATIVA='" & Format(z, "mm/dd/yyyy") & "',DATAARQUIVA=NULL,DATASUSPENSO=NULL,DATACANCEL=NULL,OBSR='" & Mask(CStr(t)) & "' "
                    Sql = Sql & "WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                    cn.Execute Sql, rdExecDirect
                    lblDtReativacao.Caption = z
                    lblDtArquivamento.Caption = "  /  /    "
                    lblDtCancelamento.Caption = "  /  /    "
                    lblDtSuspencao.Caption = "  /  /    "
                End If
        End If
     End If
End If

End Sub

Private Sub cmdRepair_Click()
Dim nAno As Integer, nNumero As Long
If lblAno.Caption = "" Or lblNumProc = "" Then
    MsgBox "Selecione um processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If NomeDeLogin = "RENATA" Or NomeDeLogin = "PEDROS" Or NomeDeLogin = "TANIA" Or NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RODRIGOG" Or NomeDeLogin = "GLEISE" Then
Else
    MsgBox "ACESSO NEGADO!!!", vbCritical, "ATENÇÃO"
    Exit Sub
End If

If MsgBox("Deseja corrigir a tramitação deste processo?" & vbCrLf & "Qualquer alteração na ordem da tramitação será perdida." & vbCrLf & "Deseja continuar?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub


nAno = Val(lblAno.Caption)
nNumero = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))

Sql = "DELETE FROM TRAMITACAOCC WHERE ANO=" & nAno & " AND NUMERO=" & nNumero
cn.Execute Sql, rdExecDirect

If cn.RowsAffected > 0 Then
    MsgBox "Tramitação corrigida.", vbExclamation, "Atenção"
Else
    MsgBox "Tramitação não pode ser corrigida.", vbExclamation, "Atenção"
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
                    Sql = "UPDATE PROCESSOGTI SET DATASUSPENSO='" & Format(z, "mm/dd/yyyy") & "',OBSS='" & Mask(CStr(t)) & "' WHERE NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " AND ANO=" & Val(lblAno.Caption)
                    cn.Execute Sql, rdExecDirect
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

frmTramite.show
frmTramite.ZOrder 0
End Sub

Private Sub Form_Activate()
On Error Resume Next
If cmdConsultar.Visible = True And CodProcesso > 0 Then
    Limpa
    Le
End If

End Sub

Private Sub Form_Load()
frReq2.Visible = False
CodCidadao = 0
nLoginId = RetornaUsuarioID(NomeDeLogin)
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
Sql = "SELECT processogti.ANO, processogti.NUMERO, processogti.FISICO, processogti.ORIGEM, processogti.INTERNO, processogti.CODASSUNTO, processogti.COMPLEMENTO, "
Sql = Sql & "processogti.OBSERVACAO, processogti.DATAENTRADA, processogti.DATAREATIVA, processogti.DATACANCEL, processogti.DATAARQUIVA,"
Sql = Sql & "processogti.DATASUSPENSO, processogti.ETIQUETA, processogti.CODCIDADAO, processogti.MOTIVOCANCEL,"
Sql = Sql & "processogti.CENTROCUSTO, processogti.OBSA, processogti.OBSC, processogti.OBSS, processogti.OBSR, processogti.HORA, processogti.INSC,"
Sql = Sql & " ProcessoGTI.TIPOEND, ProcessoGTI.USERID, Usuario.NomeLogin, Usuario.NomeCompleto "
Sql = Sql & "FROM processogti LEFT OUTER JOIN usuario ON processogti.USERID = usuario.Id WHERE NUMERO=" & CodProcesso & " AND ANO=" & AnoProcesso
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblNumProc.Caption = Format(CodProcesso & RetornaDVProcesso(CodProcesso), "000000-0")
        lblAno.Caption = AnoProcesso
        chkFisico.value = IIf(!FISICO, 1, 0)
        chkInterno.value = IIf(!INTERNO, 1, 0)
        If chkInterno.value = 1 Then
            cmbOrigem.ListIndex = 0
        Else
            cmbOrigem.ListIndex = 1
        End If
        
        For k = 0 To cmbAssunto.ListCount - 1
            If cmbAssunto.ItemData(k) = !CODASSUNTO Then
               cmbAssunto.ListIndex = k
               Exit For
            End If
        Next
        If chkInterno.value = 1 Then
            For k = 0 To cmbReq.ListCount - 1
                cmbReq.ListIndex = k
                If cmbReq.ItemData(cmbReq.ListIndex) = !CENTROCUSTO Then
                    Exit For
                End If
            Next
        End If
        If IsNull(!tipoend) Then
            optEnd(0).value = True
        Else
            If !tipoend = "R" Then
                optEnd(0).value = True
            Else
                optEnd(1).value = True
            End If
        End If
        
        txtCompl.Text = SubNull(!Complemento)
        txtObs.Text = SubNull(!OBSERVACAO)
        txtInsc.Text = SubNull(!INSC)
        'lblAtendente.Caption = SubNull(!RESPONSAVEL)
        lblAtendente.Caption = SubNull(!NomeLogin)
        lblAtendente.Tag = Val(SubNull(!userid))
        lblCodCid.Caption = !CodCidadao
        Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !CodCidadao
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        lblNomeCid.Caption = SubNull(RdoAux2!nomecidadao)
        RdoAux2.Close
        If Not IsNull(!HORA) Then lblHora.Caption = !HORA
        If Not IsNull(!DATAENTRADA) Then lblDtEntrada.Caption = Format(!DATAENTRADA, "dd/mm/yyyy")
        If Not IsNull(!DATAREATIVA) Then lblDtReativacao.Caption = Format(!DATAREATIVA, "dd/mm/yyyy")
        If Not IsNull(!DataCancel) Then lblDtCancelamento.Caption = Format(!DataCancel, "dd/mm/yyyy")
        If Not IsNull(!DATAARQUIVA) Then lblDtArquivamento.Caption = Format(!DATAARQUIVA, "dd/mm/yyyy")
        If Not IsNull(!DATASUSPENSO) Then lblDtSuspencao.Caption = Format(!DATASUSPENSO, "dd/mm/yyyy")
        
       'CARREGA ENDEREÇO
        Sql = "SELECT PROCESSOEND.CODLOGR, vwLOGRADOURO.NOMETIPOLOG, vwLOGRADOURO.NOMETITLOG, "
        Sql = Sql & "vwLOGRADOURO.NomeLogradouro , PROCESSOEND.Numero FROM PROCESSOEND INNER JOIN "
        Sql = Sql & "vwLOGRADOURO ON PROCESSOEND.CODLOGR = vwLOGRADOURO.CODLOGRADOURO "
        Sql = Sql & "Where PROCESSOEND.ANO = " & AnoProcesso & " And PROCESSOEND.NUMPROCESSO = " & CodProcesso
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sNomeLog = Trim$(SubNull(!NomeTipoLog)) & " " & Trim$(SubNull(!NomeTitLog)) & " " & !NomeLogradouro
                grdEnd.AddItem !CodLogr & Chr(9) & sNomeLog & Chr(9) & !Numero
               .MoveNext
            Loop
           .Close
        End With
       
       'CARREGA DOC
       Sql = "SELECT PROCESSODOC.CODDOC, DOCUMENTO.NOME, PROCESSODOC.DATA FROM PROCESSODOC INNER JOIN "
       Sql = Sql & "DOCUMENTO ON PROCESSODOC.CODDOC = DOCUMENTO.CODIGO Where PROCESSODOC.ANO = " & AnoProcesso & " And PROCESSODOC.Numero = " & CodProcesso
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            If .RowCount = 0 Then GoTo DOC2
            Do Until .EOF
                Set itmX = lvDoc.ListItems.Add(, "C" & Format(!CODDOC, "000"), !Nome)
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
    Sql = "SELECT ASSUNTODOC.CODDOC, DOCUMENTO.NOME FROM DOCUMENTO INNER JOIN "
    Sql = Sql & "ASSUNTODOC ON DOCUMENTO.CODIGO = ASSUNTODOC.CODDOC INNER JOIN ASSUNTO ON "
    Sql = Sql & "ASSUNTODOC.CODASSUNTO = ASSUNTO.CODIGO Where ASSUNTO.Codigo = " & cmbAssunto.ItemData(cmbAssunto.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            Set itmX = lvDoc.ListItems.Add(, "C" & Format(!CODDOC, "000"), !Nome)
             nNao = nNao + 1
           .MoveNext
        Loop
       .Close
    End With
    lblDoc1.Caption = nSim: lblDoc2.Caption = nNao
    
   
End With

fim:
Sql = "SELECT * FROM ANEXO WHERE ANO=" & AnoProcesso & " AND NUMERO=" & CodProcesso
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblAnexo.Caption = CStr(.RowCount) & " Anexo(s) "
    Else
        lblAnexo.Caption = "Nenhum"
    End If
   .Close
End With
CodProcesso = 0
bExec = True
End Sub

Private Sub CarregaCombo(bAtivo As Boolean)

cmbAssunto.Clear
Sql = "SELECT CODIGO,NOME FROM ASSUNTO"
If bAtivo Then
    Sql = Sql & " WHERE ATIVO=1"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbAssunto.AddItem !Nome
       cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With

cmbOrigem.Clear
Sql = "SELECT CODIGO,NOME FROM ORIGEM"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbOrigem.AddItem !Nome
       cmbOrigem.ItemData(cmbOrigem.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With
cmbReq.Clear
Sql = "SELECT CODIGO,DESCRICAO FROM CENTROCUSTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbReq.AddItem !descricao
       cmbReq.ItemData(cmbReq.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Limpa()
On Error Resume Next
lblNumProc.Caption = ""
lblAno.Caption = ""
cmbOrigem.ListIndex = -1
cmbAssunto.ListIndex = -1
chkFisico.value = 1
chkInterno.value = 0
txtInsc.Text = ""
txtCompl.Text = ""
lblAtendente.Caption = ""
lblAtendente.Tag = "0"
lblDtArquivamento.Caption = "  /  /    "
lblDtCancelamento.Caption = "  /  /    "
lblDtEntrada.Caption = "  /  /    "
lblDtReativacao.Caption = "  /  /    "
lblDtSuspencao.Caption = "  /  /    "
txtObs.Text = ""
grdEnd.Rows = 1
lblAnexo.Caption = ""
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
    txtObs.BackColor = Kde
    txtInsc.BackColor = Kde
    cmbReq.Enabled = False
    cmbOrigem.Enabled = False
    cmbAssunto.Enabled = False
    chkFisico.Enabled = False
    chkInterno.Enabled = False
    txtCompl.Locked = True
    txtObs.Locked = True
    txtInsc.Locked = True
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
    optEnd(0).Enabled = False
    optEnd(1).Enabled = False
ElseIf Tipo = "INCLUIR" Then
    optEnd(0).Enabled = True
    optEnd(1).Enabled = True
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
    cmbReq.Enabled = True
    If Evento = "Novo" Then
        cmbAssunto.BackColor = Branco
        chkFisico.BackColor = Kde
        chkInterno.BackColor = Kde
        cmbAssunto.Enabled = True
        '
        chkInterno.Enabled = True
        txtCompl.BackColor = Branco
        txtCompl.Locked = False
        txtObs.BackColor = Branco
        txtObs.Locked = False
        txtInsc.BackColor = Branco
        txtInsc.Locked = False
        'cmdEditCid.Enabled = True
        cmdEditEnd.Enabled = True
    Else
        txtCompl.Locked = True
        txtCompl.BackColor = Kde
        If chkInterno.value = False Then
            txtObs.Locked = True
            txtObs.BackColor = Kde
            txtInsc.Locked = True
            txtInsc.BackColor = Kde
        Else
            txtObs.Locked = False
            txtObs.BackColor = Branco
            txtInsc.Locked = False
            txtInsc.BackColor = Branco
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
FormHagana
cmdEditInsc.Enabled = cmdAlterar.Enabled
If cmdAlterar.Enabled = False Then
    If NomeDeLogin = "LUIZH" Or NomeDeLogin = "ROSANGELA" Or NomeDeLogin = "MARTA" Or NomeDeLogin = "RITA" Or NomeDeLogin = "DANIELAR" Or NomeDeLogin = "ANAP" Or NomeDeLogin = "MORITA" Or NomeDeLogin = "PATRICIAG" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "EDUARDO" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "PAULA" Or NomeDeLogin = "NOELI" Or NomeDeLogin = "MARILIA" Or NomeDeLogin = "RODRIGOC" Then
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
    Sql = "SELECT MAX(NUMERO) AS MAXIMO FROM PROCESSOGTI WHERE ANO=" & Val(lblAno.Caption)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            MaxCod = 1
        Else
            MaxCod = !maximo + 1
        End If
       .Close
    End With
Else
    MaxCod = Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
End If

nCodCidadao = Val(lblCodCid.Caption)
If cmbReq.ListIndex > -1 Then
    nCodReq = cmbReq.ItemData(cmbReq.ListIndex)
'    nCodCidadao = 0
End If

If Evento = "Novo" Then
    Sql = "INSERT PROCESSOGTI(ANO,NUMERO,FISICO,ORIGEM,INTERNO,CODASSUNTO,COMPLEMENTO,"
    Sql = Sql & "OBSERVACAO,DATAENTRADA,DATAREATIVA,DATACANCEL,DATAARQUIVA,DATASUSPENSO,"
    Sql = Sql & "ETIQUETA,CODCIDADAO,MOTIVOCANCEL,CENTROCUSTO,HORA,INSC,TIPOEND,userid) VALUES(" & Val(lblAno.Caption) & ","
    'Sql = Sql & "ETIQUETA,CODCIDADAO,RESPONSAVEL,MOTIVOCANCEL,CENTROCUSTO,HORA,INSC,TIPOEND,userid) VALUES(" & Val(lblAno.Caption) & ","
    Sql = Sql & MaxCod & "," & chkFisico.value & "," & cmbOrigem.ItemData(cmbOrigem.ListIndex) & ","
    Sql = Sql & chkInterno.value & "," & cmbAssunto.ItemData(cmbAssunto.ListIndex) & ",'" & Mask(txtCompl.Text) & "','"
    Sql = Sql & Mask(txtObs.Text) & "','" & Format(lblDtEntrada.Caption, "mm/dd/yyyy") & "'," & "Null" & ","
    Sql = Sql & "Null" & "," & "Null" & "," & "Null" & "," & "0" & "," & nCodCidadao & ","
    Sql = Sql & "Null" & "," & nCodReq & ",'" & lblHora.Caption & "'," & Val(txtInsc.Text) & ",'" & IIf(optEnd(0).value = True, "R", "C") & "'," & Val(lblAtendente.Tag) & ")"
    lblNumProc.Caption = CStr(MaxCod) & RetornaDVProcesso(MaxCod)
    lblNumProc.Caption = Format(lblNumProc.Caption, "000000-0")
Else
    Sql = "UPDATE PROCESSOGTI SET FISICO=" & chkFisico.value & ",ORIGEM=" & cmbOrigem.ItemData(cmbOrigem.ListIndex) & ","
    Sql = Sql & "INTERNO=" & chkInterno.value & ",CODASSUNTO=" & cmbAssunto.ItemData(cmbAssunto.ListIndex) & ","
    Sql = Sql & "COMPLEMENTO='" & Mask(txtCompl.Text) & "',OBSERVACAO='" & Mask(txtObs.Text) & "',INSC=" & Val(txtInsc.Text) & ","
    Sql = Sql & "CODCIDADAO=" & nCodCidadao & ",CENTROCUSTO=" & nCodReq & ",TIPOEND='" & IIf(optEnd(0).value = True, "R", "C") & "'"
    Sql = Sql & " WHERE NUMERO=" & MaxCod & " AND ANO=" & Val(lblAno.Caption)
End If
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM PROCESSOEND WHERE ANO=" & Val(lblAno.Caption) & " AND NUMPROCESSO=" & MaxCod
cn.Execute Sql, rdExecDirect

With grdEnd
    For p = 1 To .Rows - 1
        Sql = "INSERT PROCESSOEND (ANO,NUMPROCESSO,CODLOGR,NUMERO) VALUES("
        Sql = Sql & Val(lblAno.Caption) & "," & MaxCod & "," & .TextMatrix(p, 0) & ",'" & .TextMatrix(p, 2) & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With

Sql = "DELETE FROM PROCESSODOC WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & MaxCod
cn.Execute Sql, rdExecDirect

With lvDoc
    For p = 1 To .ListItems.Count
        Sql = "INSERT PROCESSODOC (ANO,NUMERO,CODDOC,DATA) VALUES("
        Sql = Sql & Val(lblAno.Caption) & "," & MaxCod & "," & Val(Right$(.ListItems(p).Key, 3)) & ","
        Sql = Sql & IIf(.ListItems(p).SubItems(1) = "", "Null", "'" & Format(.ListItems(p).SubItems(1), "mm/dd/yyyy") & "'") & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

'GRAVA CIDADÃO NO PROCESSO A PARTIR DE 01/2011
If nCodCidadao > 0 Then
    
    If Evento <> "Novo" Then
        Sql = "DELETE FROM PROCESSOCIDADAO WHERE ANOPROC=" & Val(lblAno.Caption) & " AND NUMPROC=" & MaxCod
        cn.Execute Sql, rdExecDirect
    End If
    Sql = "SELECT * FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If Not IsNull(!CPF) And SubNull(!CPF) <> "" Then
        
            sDoc = !CPF
        Else
            If Not IsNull(!Cnpj) Then
                sDoc = !Cnpj
            Else
                sDoc = ""
            End If
        End If
        Sql = "INSERT PROCESSOCIDADAO (ANOPROC,NUMPROC,CODCIDADAO,NOMECIDADAO,DOC,CODLOGRADOURO,NUMIMOVEL,COMPLEMENTO,CODBAIRRO,CODCIDADE,SIGLAUF,CEP,RG,ORGAO) VALUES("
        Sql = Sql & Val(lblAno.Caption) & "," & MaxCod & "," & nCodCidadao & ",'" & Mask(!nomecidadao) & "','" & sDoc & "'," & Val(SubNull(!CodLogradouro)) & ","
        Sql = Sql & Val(SubNull(!NUMIMOVEL)) & ",'" & Mask(SubNull(!Complemento)) & "'," & Val(SubNull(!CodBairro)) & "," & Val(SubNull(!CodCidade)) & ",'" & SubNull(!SiglaUF) & "'," & Val(SubNull(!Cep)) & ",'"
        Sql = Sql & SubNull(!rg) & "','" & Mask(SubNull(!ORGAO)) & "')"
        cn.Execute Sql, rdExecDirect
    End With
    
End If

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

Private Sub txtInsc_KeyPress(KeyAscii As Integer)
Tweak txtInsc, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumLog_KeyPress(KeyAscii As Integer)
Tweak txtNumLog, KeyAscii, IntegerPositive
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
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLog) & "%' "
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

bevNew = cmdNovo.Enabled
bevEdit = cmdAlterar.Enabled
bevDel = cmdExcluir.Enabled
bEvPrint = cmdImprimir.Enabled
bEvArquivar = cmdArquivar.Enabled
bEvCancdelar = cmdCancelar.Enabled
bEvReativar = cmdReativar.Enabled
bEvSuspender = cmdSuspender.Enabled
bEvAnexos = cmdAnexar.Enabled
bEvCorrigir = cmdRepair.Enabled
bEvArquivos = cmdArquivos.Enabled

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
cmdAnexar.Enabled = False

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

cmdImprimir.Enabled = bEvPrint
cmdNovo.Enabled = bevNew
cmdAlterar.Enabled = bevEdit
cmdExcluir.Enabled = bevDel
cmdArquivar.Enabled = bEvArquivar
cmdCancelar.Enabled = bEvCancdelar
cmdReativar.Enabled = bEvReativar
cmdSuspender.Enabled = bEvSuspender
cmdAnexar.Enabled = bEvAnexos
cmdRepair.Enabled = bEvCorrigir
cmdArquivos.Enabled = bEvArquivos

If Evento <> "" Then
    chkInterno.Enabled = True: chkFisico.Enabled = True: cmbOrigem.Enabled = True: txtCompl.Locked = False: txtObs.Locked = False: cmbAssunto.Enabled = True
    cmbOrigem.BackColor = Branco: txtCompl.BackColor = Branco: txtObs.BackColor = Branco: cmbAssunto.BackColor = Branco
End If
End Sub

Private Sub FormHagana()
evNew = 2
evEdit = 3
evDel = 4
evEsp = 11
evImp = 15
evAne = 9
evCan = 19
evRea = 18
evSus = 17
evArq = 16

bNew = False: bEdit = False: bDel = False: bImp = False: bEsp = False: bAne = False: bCan = False: bRea = False: bSus = False: bArq = False
If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True
If InStr(1, sRet, Format(evEsp, "000"), vbBinaryCompare) > 0 Then bEsp = True
If InStr(1, sRet, Format(evImp, "000"), vbBinaryCompare) > 0 Then bImp = True
If InStr(1, sRet, Format(evAne, "000"), vbBinaryCompare) > 0 Then bAne = True
If InStr(1, sRet, Format(evCan, "000"), vbBinaryCompare) > 0 Then bCan = True
If InStr(1, sRet, Format(evRea, "000"), vbBinaryCompare) > 0 Then bRea = True
If InStr(1, sRet, Format(evSus, "000"), vbBinaryCompare) > 0 Then bSus = True
If InStr(1, sRet, Format(evArq, "000"), vbBinaryCompare) > 0 Then bArq = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False
If Not bImp Then cmdImprimir.Enabled = False
'If Not bAne Then cmdAnexar.Enabled = False
If Not bCan Then cmdCancelar.Enabled = False
If Not bRea Then cmdReativar.Enabled = False
If Not bSus Then cmdSuspender.Enabled = False
If Not bArq Then cmdArquivar.Enabled = False

If cmdNovo.Enabled = False Then
    cmdEditCid.Enabled = False
    cmdEditDoc.Enabled = False
    'cmdEditEnd.Enabled = False
    
End If

If NomeDeLogin = "ROSE" Then
    cmdEditCid.Enabled = True
End If


End Sub

Private Function VerificaTramite() As Boolean

Dim aSeq() As Integer

'CARREGA TODOS OS TRAMITES
ReDim aSeq(0)
Sql = "SELECT * FROM tramitacaocc Where ano = " & lblAno.Caption & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        Sql = "SELECT ASSUNTOCC.SEQ,CENTROCUSTO.CODIGO, CENTROCUSTO.DESCRICAO FROM ASSUNTOCC INNER JOIN "
        Sql = Sql & "CENTROCUSTO ON ASSUNTOCC.CODCC = CENTROCUSTO.CODIGO "
        Sql = Sql & "WHERE ASSUNTOCC.CODASSUNTO =" & frmProcesso.cmbAssunto.ItemData(frmProcesso.cmbAssunto.ListIndex)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ReDim Preserve aSeq(UBound(aSeq) + 1)
                aSeq(UBound(aSeq)) = !Seq
               .MoveNext
            Loop
           .Close
        End With
    Else
        Sql = "SELECT tramitacaocc.seq, tramitacaocc.ccusto, CENTROCUSTO.DESCRICAO "
        Sql = Sql & "FROM tramitacaocc INNER JOIN CENTROCUSTO ON tramitacaocc.ccusto = CENTROCUSTO.CODIGO "
        Sql = Sql & "Where tramitacaocc.ano = " & lblAno.Caption & " And tramitacaocc.Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        Sql = Sql & " order by TRAMITACAOCC.SEQ"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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
If Val(frmProcesso.lblNumProc.Caption) > 0 Then
    For x = 1 To UBound(aSeq)
        Sql = "SELECT CCUSTO,DESCRICAO,DATAHORA,NOMECOMPLETO,DESCDESPACHO FROM VWTRAMITACAO WHERE ANO=" & lblAno.Caption
        Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        Sql = Sql & " AND SEQ=" & aSeq(x)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
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

