VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTramite 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tramitação de Processos"
   ClientHeight    =   5700
   ClientLeft      =   8535
   ClientTop       =   3375
   ClientWidth     =   10725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   10725
   Begin VB.Frame PnlInserir 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selecione o local a ser inserido abaixo do local desejado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1185
      Left            =   1350
      TabIndex        =   56
      Top             =   1590
      Visible         =   0   'False
      Width           =   6285
      Begin VB.ComboBox cmbLocal 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   330
         Width           =   5565
      End
      Begin prjChameleon.chameleonButton cmdOK 
         Height          =   345
         Left            =   4950
         TabIndex        =   58
         ToolTipText     =   "Inserir local selecionado"
         Top             =   720
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":0000
         PICN            =   "frmTramite.frx":001C
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
         Height          =   345
         Left            =   5340
         TabIndex        =   59
         ToolTipText     =   "Cancelar operação"
         Top             =   720
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":0176
         PICN            =   "frmTramite.frx":0192
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
   Begin VB.Frame pnlDespacho 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Selecione o novo despacho para este trâmite"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1185
      Left            =   1020
      TabIndex        =   43
      Top             =   1410
      Visible         =   0   'False
      Width           =   6285
      Begin VB.ComboBox cmbDespacho3 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   330
         Width           =   5565
      End
      Begin prjChameleon.chameleonButton cmdOK4 
         Height          =   345
         Left            =   4950
         TabIndex        =   45
         ToolTipText     =   "Alterar Despacho"
         Top             =   720
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":02EC
         PICN            =   "frmTramite.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancel4 
         Height          =   345
         Left            =   5340
         TabIndex        =   46
         ToolTipText     =   "Cancelar operação"
         Top             =   720
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":0462
         PICN            =   "frmTramite.frx":047E
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
   Begin VB.TextBox txtObs2 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   45
      Locked          =   -1  'True
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   54
      Top             =   4035
      Width           =   9465
   End
   Begin VB.PictureBox pic1 
      Height          =   375
      Left            =   9810
      Picture         =   "frmTramite.frx":05D8
      ScaleHeight     =   315
      ScaleWidth      =   360
      TabIndex        =   51
      Top             =   180
      Visible         =   0   'False
      Width           =   420
   End
   Begin prjChameleon.chameleonButton cmdObs 
      Height          =   705
      Left            =   9630
      TabIndex        =   48
      ToolTipText     =   "Observação do trâmite"
      Top             =   4290
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1244
      BTYPE           =   3
      TX              =   "Observ."
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
      MICON           =   "frmTramite.frx":0722
      PICN            =   "frmTramite.frx":073E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancelObs 
      Height          =   315
      Left            =   9630
      TabIndex        =   49
      ToolTipText     =   "Cancelar Edição"
      Top             =   4665
      Visible         =   0   'False
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
      MICON           =   "frmTramite.frx":0829
      PICN            =   "frmTramite.frx":0845
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
      Left            =   9630
      TabIndex        =   50
      ToolTipText     =   "Gravar observação"
      Top             =   4290
      Visible         =   0   'False
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
      MICON           =   "frmTramite.frx":099F
      PICN            =   "frmTramite.frx":09BB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   45
      Locked          =   -1  'True
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Top             =   4035
      Width           =   9465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3840
      TabIndex        =   17
      Top             =   4890
      Width           =   1185
   End
   Begin prjChameleon.chameleonButton cmdInserir 
      Height          =   345
      Left            =   45
      TabIndex        =   9
      ToolTipText     =   "Inserir um novo local para a tramitação"
      Top             =   5295
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Inserir Local"
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
      MICON           =   "frmTramite.frx":0D60
      PICN            =   "frmTramite.frx":0D7C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRemover 
      Height          =   345
      Left            =   1680
      TabIndex        =   10
      ToolTipText     =   "Remover um local de tramitação"
      Top             =   5295
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Remover Local"
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
      MICON           =   "frmTramite.frx":0ED6
      PICN            =   "frmTramite.frx":0EF2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdReceber 
      Height          =   345
      Left            =   6930
      TabIndex        =   11
      ToolTipText     =   "Receber um Processo"
      Top             =   5295
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Receber"
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
      MICON           =   "frmTramite.frx":104C
      PICN            =   "frmTramite.frx":1068
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
      Height          =   345
      Left            =   9450
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   5295
      Width           =   1215
      _ExtentX        =   2143
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmTramite.frx":1242
      PICN            =   "frmTramite.frx":125E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAcima 
      Height          =   345
      Left            =   3300
      TabIndex        =   13
      ToolTipText     =   "Mover um local acima"
      Top             =   5295
      Width           =   345
      _ExtentX        =   609
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmTramite.frx":13B8
      PICN            =   "frmTramite.frx":13D4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAbaixo 
      Height          =   345
      Left            =   3690
      TabIndex        =   14
      ToolTipText     =   "Mover um local abaixo"
      Top             =   5295
      Width           =   345
      _ExtentX        =   609
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmTramite.frx":152E
      PICN            =   "frmTramite.frx":154A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEmprestimo 
      Height          =   345
      Left            =   4095
      TabIndex        =   15
      ToolTipText     =   "Remover um local de tramitação"
      Top             =   5295
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Empréstimos"
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
      MICON           =   "frmTramite.frx":16A4
      PICN            =   "frmTramite.frx":16C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEnviar 
      Height          =   345
      Left            =   8190
      TabIndex        =   16
      ToolTipText     =   "Enviar um Processo"
      Top             =   5295
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Enviar"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTramite.frx":173C
      PICN            =   "frmTramite.frx":1758
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
      Height          =   345
      Left            =   5490
      TabIndex        =   42
      ToolTipText     =   "Alterar o despacho de processo aberto"
      Top             =   5295
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Alterar"
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
      MICON           =   "frmTramite.frx":17CC
      PICN            =   "frmTramite.frx":17E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame PnlRec 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Recebimento de Processo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2235
      Left            =   990
      TabIndex        =   29
      Top             =   1110
      Visible         =   0   'False
      Width           =   6705
      Begin VB.ComboBox cmbFunc 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1020
         Width           =   5010
      End
      Begin VB.ComboBox cmbDespacho2 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1410
         Width           =   5010
      End
      Begin prjChameleon.chameleonButton cmdOK2 
         Height          =   345
         Left            =   5745
         TabIndex        =   30
         ToolTipText     =   "Inserir local selecionado"
         Top             =   1815
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":1942
         PICN            =   "frmTramite.frx":195E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancel2 
         Height          =   345
         Left            =   6120
         TabIndex        =   33
         ToolTipText     =   "Cancelar operação"
         Top             =   1815
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":1AB8
         PICN            =   "frmTramite.frx":1AD4
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
         Caption         =   "Funcionário........:"
         Height          =   225
         Left            =   195
         TabIndex        =   39
         Top             =   1095
         Width           =   1365
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1530
         TabIndex        =   38
         Top             =   735
         Width           =   2520
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Data e Hora.......:"
         Height          =   225
         Left            =   195
         TabIndex        =   37
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Local..................:"
         Height          =   225
         Left            =   195
         TabIndex        =   36
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label lblLocal 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1530
         TabIndex        =   35
         Top             =   390
         Width           =   4890
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Despacho..........:"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   34
         Top             =   1440
         Width           =   1365
      End
   End
   Begin VB.Frame PnlEnv 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Envio de Processos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   2295
      Left            =   1200
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ComboBox cmbDespacho 
         Height          =   315
         Left            =   1485
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   1425
         Width           =   5010
      End
      Begin prjChameleon.chameleonButton cmdOk3 
         Height          =   345
         Left            =   5775
         TabIndex        =   20
         ToolTipText     =   "Inserir local selecionado"
         Top             =   1875
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":1C2E
         PICN            =   "frmTramite.frx":1C4A
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
         Height          =   345
         Left            =   6150
         TabIndex        =   21
         ToolTipText     =   "Cancelar operação"
         Top             =   1875
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTramite.frx":1DA4
         PICN            =   "frmTramite.frx":1DC0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblLocal2 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1515
         TabIndex        =   28
         Top             =   450
         Width           =   4890
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Local..................:"
         Height          =   225
         Left            =   180
         TabIndex        =   27
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Data e Hora.......:"
         Height          =   225
         Left            =   180
         TabIndex        =   26
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label lblData2 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1515
         TabIndex        =   25
         Top             =   795
         Width           =   2520
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário........:"
         Height          =   225
         Left            =   180
         TabIndex        =   24
         Top             =   1155
         Width           =   1365
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Despacho..........:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1515
         TabIndex        =   22
         Top             =   1110
         Width           =   4905
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTramite 
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   1020
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   4366
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmTramite.frx":1F1A
   End
   Begin VB.Frame pnlEmprestimo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Consulta de Empréstimos e Devolução do Sistema Antigo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   2865
      Left            =   300
      TabIndex        =   40
      Top             =   720
      Visible         =   0   'False
      Width           =   9585
      Begin MSFlexGridLib.MSFlexGrid grdEmp 
         Height          =   2475
         Left            =   90
         TabIndex        =   41
         Top             =   300
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   4366
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmTramite.frx":1FEF
      End
   End
   Begin VB.Label lblTitObs 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   55
      Top             =   3840
      Width           =   9195
   End
   Begin VB.Label lblObsInterna 
      Caption         =   "ObservaçãoInterna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5310
      MouseIcon       =   "frmTramite.frx":20A3
      MousePointer    =   99  'Custom
      TabIndex        =   53
      Top             =   3600
      Width           =   1545
   End
   Begin VB.Label lblObsGeral 
      Caption         =   "Observação Geral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3450
      MouseIcon       =   "frmTramite.frx":21F5
      MousePointer    =   99  'Custom
      TabIndex        =   52
      Top             =   3600
      Width           =   1545
   End
   Begin VB.Label lblRequerente 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   8
      Top             =   720
      Width           =   6495
   End
   Begin VB.Label lblAssunto 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   7
      Top             =   420
      Width           =   6495
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2970
      TabIndex        =   6
      Top             =   120
      Width           =   705
   End
   Begin VB.Label lblNumProc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1500
      TabIndex        =   5
      Top             =   120
      Width           =   915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requerente.........:"
      Height          =   225
      Index           =   7
      Left            =   90
      TabIndex        =   4
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto...............:"
      Height          =   225
      Index           =   6
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   225
      Index           =   1
      Left            =   2490
      TabIndex        =   2
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo...:"
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, bExec As Boolean
Dim evRem As Integer, bRem As Boolean, sRet As String, sLogin As String, sTipoObs As String
Dim nLoginId As Integer

Private Sub cmdAbaixo_Click()
Dim sTemp As String
With grdTramite
    If .Rows = 1 Then
        MsgBox "Não há local a ser movido.", vbExclamation, "Atenção"
    Else
        If .Row = .Rows - 1 Then
            MsgBox "Não é possível mover para baixo o último local", vbExclamation, "Atenção"
        Else
            If .TextMatrix(.Row, 3) <> "" Then
                MsgBox "Não é possível mover este local porque já houve recebimento de processo no mesmo.", vbExclamation, "Atenção"
            Else
                If .TextMatrix(.Row + 1, 3) <> "" Then
                    MsgBox "Não é possível mover este local porque já houve recebimento de processo no local abaixo.", vbExclamation, "Atenção"
                Else
                    sTemp = .TextMatrix(.Row, 1)
                   .TextMatrix(.Row, 1) = .TextMatrix(.Row + 1, 1)
                   .TextMatrix(.Row + 1, 1) = sTemp
                    sTemp = .TextMatrix(.Row, 2)
                   .TextMatrix(.Row, 2) = .TextMatrix(.Row + 1, 2)
                   .TextMatrix(.Row + 1, 2) = sTemp
                   .Row = .Row + 1
                   .ColSel = 6
                    GravaMovimentoCC
                End If
            End If
        End If
    End If
End With

End Sub

Private Sub cmdAcima_Click()
Dim sTemp As String
With grdTramite
    If .Rows = 1 Then
        MsgBox "Não há local a ser movido.", vbExclamation, "Atenção"
    Else
        If .Row = 1 Then
            MsgBox "Não é possível mover para cima o 1º local", vbExclamation, "Atenção"
        Else
            If .TextMatrix(.Row, 3) <> "" Then
                MsgBox "Não é possível mover este local porque já houve recebimento de processo no mesmo.", vbExclamation, "Atenção"
            Else
                If .TextMatrix(.Row - 1, 3) <> "" Then
                    MsgBox "Não é possível mover este local porque já houve recebimento de processo no local acima.", vbExclamation, "Atenção"
                Else
                    sTemp = .TextMatrix(.Row, 1)
                   .TextMatrix(.Row, 1) = .TextMatrix(.Row - 1, 1)
                   .TextMatrix(.Row - 1, 1) = sTemp
                    sTemp = .TextMatrix(.Row, 2)
                   .TextMatrix(.Row, 2) = .TextMatrix(.Row - 1, 2)
                   .TextMatrix(.Row - 1, 2) = sTemp
                   .Row = .Row - 1
                   .ColSel = 6
                    GravaMovimentoCC
                End If
            End If
        End If
    End If
End With

End Sub

Private Sub cmdAlterar_Click()
Dim Sql As String, RdoAux As rdoResultset

If IsDate(frmProcesso.lblDtArquivamento) Then
    MsgBox "O Processo está arquivado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmProcesso.lblDtCancelamento) Then
    MsgBox "O Processo está Cancelado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmProcesso.lblDtSuspencao) Then
    MsgBox "O Processo está Suspenso e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

With grdTramite
    If .Row = 0 Then
        Exit Sub
    Else
        If .Row < .Rows - 1 Then
            If .TextMatrix(.Row + 1, 3) <> "" Then
                MsgBox "Não é possível alterar este despacho, pois o próximo local já foi tramitado.", vbExclamation, "Atenção"
                Exit Sub
            End If
        End If
    End If
End With

If grdTramite.TextMatrix(grdTramite.Row, 3) = "" Then
    MsgBox "Este local ainda não foi tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If


'Sql = "SELECT NOME,CODIGOCC FROM USUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Sql = "SELECT USERID,CODIGOCC FROM USUARIOCC WHERE USERID=" & nLoginId & " AND CODIGOCC=" & Val(grdTramite.TextMatrix(grdTramite.Row, 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Este usuário não tem permissão para alterar o despacho deste local.", vbCritical, "Atenção"
    RdoAux.Close:    Exit Sub
End If

pnlDespacho.Visible = True: pnlDespacho.ZOrder 0
PreencheListas
If grdTramite.TextMatrix(grdTramite.Row, 6) <> "" Then
    cmbDespacho3.Text = grdTramite.TextMatrix(grdTramite.Row, 6)
End If
End Sub

Private Sub cmdCancel_Click()
LiberarTela
PnlInserir.Visible = False

End Sub

Private Sub cmdCancel2_Click()
PnlRec.Visible = False
LiberarTela
End Sub

Private Sub cmdCancel4_Click()
pnlDespacho.Visible = False
End Sub

Private Sub cmdCancelar3_Click()
PnlEnv.Visible = False
LiberarTela

End Sub

Private Sub cmdCancelObs_Click()
EventosObs False
End Sub

Private Sub cmdEmprestimo_Click()

If cmdEmprestimo.value = True Then
    cmdSair.Enabled = False
    cmdInserir.Enabled = False
    cmdReceber.Enabled = False
Else
    cmdSair.Enabled = True
    cmdInserir.Enabled = True
    cmdReceber.Enabled = True
End If

grdEmp.Rows = 1
Sql = "SELECT   EMPRESTIMO.SEQ, EMPRESTIMO.SEQ2, EMPRESTIMO.SEQ3, EMPRESTIMO.TIPO, "
Sql = Sql & " EMPRESTIMO.CODCC , CENTROCUSTO.DESCRICAO, EMPRESTIMO.DATAHORA, EMPRESTIMO.TEMPO "
Sql = Sql & "FROM  EMPRESTIMO INNER JOIN  CENTROCUSTO ON EMPRESTIMO.CODCC = CENTROCUSTO.CODIGO "
Sql = Sql & "WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         grdEmp.AddItem Format(!Seq, "00") & Chr(9) & IIf(!Tipo = 1, "EMPRÉSTIMO", "DEVOLUÇÃO") & Chr(9) & !descricao & Chr(9) & Format(!DATAHORA, "dd/mm/yyyy hh:mm") & Chr(9) & !TEMPO
        .MoveNext
    Loop
   .Close
End With

If grdEmp.Rows = 1 Then
    MsgBox "Processo não possue empréstimos ou devoluções cadastrados no sistema antigo de protocolo.", vbExclamation, "Atenção"
    cmdEmprestimo.value = False
    cmdSair.Enabled = True
    cmdInserir.Enabled = True
    cmdReceber.Enabled = True
Else
    pnlEmprestimo.Visible = cmdEmprestimo.value
    If pnlEmprestimo.Visible = True Then pnlEmprestimo.ZOrder 0
End If

End Sub

Private Sub cmdEnviar_Click()

If IsDate(frmProcesso.lblDtArquivamento) Then
    MsgBox "O Processo está arquivado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmProcesso.lblDtCancelamento) Then
    MsgBox "O Processo está Cancelado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmProcesso.lblDtSuspencao) Then
    MsgBox "O Processo está Suspenso e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

With grdTramite
    If .Rows = 1 Then
        MsgBox "Selecione um local.", vbExclamation, "Atenção"
    Else
        If .Row = 1 Then
            If .TextMatrix(.Row, 8) <> "" Then
                MsgBox "Este local ja foi tramitado.", vbExclamation, "Atenção"
            Else
                If .TextMatrix(.Row, 3) = "" Then
                    MsgBox "Este local ainda não foi tramitado.", vbExclamation, "Atenção"
                Else
                    Enviar
                End If
            End If
        Else
            If .TextMatrix(.Row, 8) <> "" Then
                MsgBox "Este local já foi tramitado.", vbExclamation, "Atenção"
            Else
                If .TextMatrix(.Row - 1, 8) = "" And .TextMatrix(.Row - 1, 4) = "" Then
                    MsgBox "O local anterior ainda não foi tramitado.", vbExclamation, "Atenção"
                Else
                    If .TextMatrix(.Row, 3) = "" Then
                        MsgBox "Este local ainda não foi tramitado.", vbExclamation, "Atenção"
                    Else
                        Enviar
                    End If
                End If
            End If
        End If
    End If
End With

End Sub

Private Sub cmdGravarObs_Click()

If MsgBox("Gravar a observação?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    

    Sql = "UPDATE TRAMITACAO SET OBS='" & Mask(txtObs.Text) & "',OBSINTERNA='" & Mask(txtObs2.Text) & "' WHERE ANO=" & lblAno.Caption
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(grdTramite.Row, 0)
    cn.Execute Sql, rdExecDirect

End If

EventosObs False
End Sub

Private Sub cmdInserir_Click()

If IsDate(frmProcesso.lblDtArquivamento.Caption) Then
    MsgBox "Não é possível inserir local,processo arquivado.", vbExclamation, "Atenção"
    Exit Sub
End If


With grdTramite
    If .Row = 0 Then
        Exit Sub
    Else
        If .Row < .Rows - 1 Then
            If .TextMatrix(.Row + 1, 3) <> "" Then
                MsgBox "Não é possível inserir um local, pois o próximo local já foi tramitado.", vbExclamation, "Atenção"
                Exit Sub
            End If
        End If
    End If
End With

cmbLocal.Clear
Sql = "SELECT CODIGO, DESCRICAO FROM CENTROCUSTO WHERE ATIVO = 1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbLocal.AddItem !descricao
        cmbLocal.ItemData(cmbLocal.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
If cmbLocal.ListCount > 0 Then cmbLocal.ListIndex = 0
BloquearTela
PnlInserir.Visible = True
End Sub

Private Sub cmdObs_Click()

'Sql = "SELECT NOME,CODIGOCC FROM USUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Sql = "SELECT USERID,CODIGOCC FROM USUARIOCC WHERE USERID=" & nLoginId & " AND CODIGOCC=" & Val(grdTramite.TextMatrix(grdTramite.Row, 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Este usuário não tem permissão para alterar a observação deste local.", vbCritical, "Atenção"
    RdoAux.Close
Else
    EventosObs True
End If

End Sub

Private Sub EventosObs(bEdit As Boolean)

cmdGravarObs.Visible = bEdit
cmdCancelObs.Visible = bEdit
cmdObs.Visible = Not bEdit
txtObs.Locked = Not bEdit
txtObs2.Locked = Not bEdit

End Sub

Private Sub cmdOK_Click()

grdTramite.AddItem 9 & Chr(9) & cmbLocal.ItemData(cmbLocal.ListIndex) & Chr(9) & cmbLocal.Text, grdTramite.Row + 1
Renumerar
LiberarTela
PnlInserir.Visible = False

GravaMovimentoCC

End Sub

Private Sub cmdOK2_Click()
Dim RdoAux As rdoResultset

If cmbFunc.ListIndex = -1 Then
    MsgBox "Selecione o Funcionário.", vbExclamation, "Atenção"
    Exit Sub
End If
'If cmbDespacho.ListIndex = -1 Then
'    MsgBox "Selecione o Despacho.", vbExclamation, "Atenção"
'    Exit Sub
'End If
PnlRec.Visible = False
LiberarTela

With grdTramite
   .TextMatrix(.Row, 3) = Left$(lblData.Caption, 10)
   .TextMatrix(.Row, 4) = Right$(lblData.Caption, 5)
   .TextMatrix(.Row, 5) = cmbFunc.Text
   .TextMatrix(.Row, 6) = cmbDespacho2.Text

    Sql = "DELETE FROM TRAMITACAO WHERE ANO=" & lblAno.Caption
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Sql = Sql & " AND SEQ=" & .TextMatrix(.Row, 0)
    cn.Execute Sql, rdExecDirect

'    Sql = "SELECT NOMELOGIN FROM USUARIO WHERE NOMECOMPLETO='" & Mask(cmbFunc.Text) & "'"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If .RowCount > 0 Then
'            sLogin = !NomeLogin
'        End If
'       .Close
'    End With

    If cmbDespacho2.ListIndex > -1 Then
        'Sql = "INSERT TRAMITACAO (ANO,NUMERO,SEQ,CCUSTO,DATAHORA,DESPACHO,USUARIO) VALUES("
        Sql = "INSERT TRAMITACAO (ANO,NUMERO,SEQ,CCUSTO,DATAHORA,DESPACHO,USERID) VALUES("
        Sql = Sql & lblAno.Caption & "," & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & ","
        Sql = Sql & .TextMatrix(.Row, 0) & "," & .TextMatrix(.Row, 1) & ",'" & Format(lblData.Caption, "mm/dd/yyyy hh:mm") & "',"
        Sql = Sql & cmbDespacho2.ItemData(cmbDespacho2.ListIndex) & "," & cmbFunc.ItemData(cmbFunc.ListIndex) & ")"
        'Sql = Sql & cmbDespacho2.ItemData(cmbDespacho2.ListIndex) & ",'" & Trim$(sLogin) & "')"
     Else
        'Sql = "INSERT TRAMITACAO (ANO,NUMERO,SEQ,CCUSTO,DATAHORA,USUARIO) VALUES("
        Sql = "INSERT TRAMITACAO (ANO,NUMERO,SEQ,CCUSTO,DATAHORA,USERID) VALUES("
        Sql = Sql & lblAno.Caption & "," & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & ","
        Sql = Sql & .TextMatrix(.Row, 0) & "," & .TextMatrix(.Row, 1) & ",'" & Format(lblData.Caption, "mm/dd/yyyy hh:mm") & "'," & cmbFunc.ItemData(cmbFunc.ListIndex) & ")"
        'Sql = Sql & .TextMatrix(.Row, 0) & "," & .TextMatrix(.Row, 1) & ",'" & Format(lblData.Caption, "mm/dd/yyyy hh:mm") & "','" & Trim$(sLogin) & "')"
    End If
    cn.Execute Sql, rdExecDirect

End With

CalculaDias

End Sub

Private Sub cmdOk3_Click()
'Dim bFiscal As Boolean
If cmbDespacho.ListIndex = -1 Then
    MsgBox "Selecione o despacho.", vbExclamation, "Atenção"
    Exit Sub
End If

grdTramite.TextMatrix(grdTramite.Row, 6) = cmbDespacho.Text
grdTramite.TextMatrix(grdTramite.Row, 8) = Format(Now, "dd/mm/yyyy")

'bFiscal = False
'Sql = "SELECT * FROM PRODUTIVIDADEFISCAL WHERE NOME='" & NomeDeLogin & "'"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If .RowCount > 0 Then
'        bFiscal = True
'    End If
'   .Close
'End With

Sql = "UPDATE TRAMITACAO SET DESPACHO=" & cmbDespacho.ItemData(cmbDespacho.ListIndex)
''Sql = Sql & " ,USUARIO='" & Trim$(NomeDeLogin) & "'  ,DATAENVIO='" & Format(Now, "mm/dd/yyyy") & "',USUARIO2='" & NomeDeLogin & "' "
'Sql = Sql & " ,DATAENVIO='" & Format(Now, "mm/dd/yyyy") & "',USUARIO2='" & NomeDeLogin & "' "
Sql = Sql & " ,DATAENVIO='" & Format(Now, "mm/dd/yyyy") & "',USERID2=" & nLoginId
Sql = Sql & " WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(grdTramite.Row, 0)
cn.Execute Sql, rdExecDirect
CalculaDias

PnlEnv.Visible = False
LiberarTela

'If bFiscal Then
'    If Not ProdIsBossLogin Then
'        frmProdutividadeTarefa.show
'        frmProdutividadeTarefa.lblNumProc.Caption = lblNumProc.Caption & "/" & lblAno.Caption
'        frmProdutividadeTarefa.lblDataTramite.Caption = Left(lblData2.Caption, 10)
'        frmProdutividadeTarefa.ProdutCarregaLista
'        frmProdutividadeTarefa.lvMain.SetFocus
'        Unload Me
'    End If
'End If

End Sub

Private Sub cmdOK4_Click()
If cmbDespacho3.ListIndex = -1 Then Exit Sub
Sql = "UPDATE TRAMITACAO SET DESPACHO=" & cmbDespacho.ItemData(cmbDespacho3.ListIndex)
Sql = Sql & " WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(grdTramite.Row, 0)
cn.Execute Sql, rdExecDirect
grdTramite.TextMatrix(grdTramite.Row, 6) = cmbDespacho3.Text
pnlDespacho.Visible = False

End Sub

Private Sub cmdReceber_Click()

If IsDate(frmProcesso.lblDtArquivamento) Then
    MsgBox "O Processo está arquivado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmProcesso.lblDtCancelamento) Then
    MsgBox "O Processo está Cancelado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmProcesso.lblDtSuspencao) Then
    MsgBox "O Processo está Suspenso e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

With grdTramite
    If .Rows = 1 Then
        MsgBox "Selecione um local.", vbExclamation, "Atenção"
    Else
        If .Row = 1 Then
            If .TextMatrix(.Row, 3) <> "" And .TextMatrix(.Row, 5) <> "" Then
                MsgBox "Este local ja foi tramitado.", vbExclamation, "Atenção"
            Else
                Receber
            End If
        Else
            If .TextMatrix(.Row, 3) <> "" And .TextMatrix(.Row, 5) <> "" Then
                MsgBox "Este local já foi tramitado.", vbExclamation, "Atenção"
            Else
                If .TextMatrix(.Row - 1, 3) = "" Then
                    MsgBox "O local anterior ainda não foi tramitado.", vbExclamation, "Atenção"
                Else
                    Receber
                End If
            End If
        End If
    End If
End With
End Sub

Private Sub cmdRemover_Click()

With grdTramite
    If .Rows = 1 Then
        MsgBox "Não há local a ser removido.", vbExclamation, "Atenção"
    Else
        If .TextMatrix(.Row, 3) <> "" Then
            MsgBox "Não é possível remover este local pois já houve tramitação nele.", vbExclamation, "Atenção"
        Else
            If MsgBox("Remover o local " & .TextMatrix(.Row, 2) & " ?", vbQuestion + vbYesNo, Confirmação) = vbYes Then
                If .Rows > 2 Then
                    .RemoveItem (.Row)
                Else
                    .Rows = 1
                End If
                GravaMovimentoCC
                Renumerar
            End If
        End If
    End If
End With

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim RdoAux2 As rdoResultset
Exit Sub
Sql = "SELECT    distinct  tramitacao.ano, tramitacao.numero, tramitacaocc.seq, tramitacaocc.ccusto "
Sql = Sql & "FROM tramitacao LEFT OUTER JOIN tramitacaocc ON tramitacao.ano = tramitacaocc.ano AND tramitacao.numero = tramitacaocc.numero "
Sql = Sql & "WHERE     (tramitacao.ano > 2004) AND (tramitacaocc.ccusto IS NULL)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cGetInputState() <> 0 Then DoEvents
'        Sql = "SELECT ANO,NUMERO FROM TRAMITACAOCC WHERE ANO=" & !ANO & " AND NUMERO=" & !Numero
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux2.RowCount = 0 Then
            Sql = "INSERT TRAMITACAOCC SELECT ANO,NUMERO,SEQ,CCUSTO FROM TRAMITACAO WHERE ANO=" & !Ano & " AND NUMERO=" & !Numero
            cn.Execute Sql, rdExecDirect
'        End If
'        RdoAux2.Close
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Form_Activate()
frmProcesso.Enabled = False
End Sub

Private Sub Form_Deactivate()
Me.ZOrder 0
End Sub

Private Sub Form_Load()
bExec = False
Ocupado
Centraliza Me
nLoginId = RetornaUsuarioID(NomeDeLogin)
sRet = RetEventUserForm(Me.Name)
grdTramite.COLWIDTH(1) = 0
lblNumProc.Caption = frmProcesso.lblNumProc.Caption
lblAno.Caption = frmProcesso.lblAno.Caption
lblAssunto.Caption = frmProcesso.cmbAssunto.Text
lblRequerente.Caption = frmProcesso.lblNomeCid.Caption
CarregaTramite
txtObs2.Visible = False
lblTitObs.Caption = "Observação Geral"
FormHagana
Liberado
End Sub

Private Sub CarregaTramite()
Dim bAchou As Boolean, nMax As Integer
If Val(lblNumProc.Caption) = 0 Then
    MsgBox "Sem numero de processo.", vbCritica, "ERRO"
    Exit Sub
End If
bExec = False
'CARREGA TODOS OS TRAMITES
Sql = "SELECT * FROM tramitacaocc Where ano = " & lblAno.Caption & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        Sql = "SELECT ano, numero, seq, ccusto, DESCRICAO From vwTRAMITACAO2 Where ano =" & lblAno.Caption & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " order by seq"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                grdTramite.AddItem !Seq & Chr(9) & !ccusto & Chr(9) & SubNull(!descricao)
                nMax = !Seq
               .MoveNext
            Loop
           .Close
        End With
        
    
        Sql = "SELECT ASSUNTOCC.SEQ,CENTROCUSTO.CODIGO, CENTROCUSTO.DESCRICAO FROM ASSUNTOCC INNER JOIN "
        Sql = Sql & "CENTROCUSTO ON ASSUNTOCC.CODCC = CENTROCUSTO.CODIGO "
        Sql = Sql & "WHERE ASSUNTOCC.CODASSUNTO =" & frmProcesso.cmbAssunto.ItemData(frmProcesso.cmbAssunto.ListIndex)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                bAchou = False
                For x = 1 To grdTramite.Rows - 1
                    If grdTramite.TextMatrix(x, 1) = !Codigo Then
                        bAchou = True
                        Exit For
                    End If
                Next
'                If Not bAchou Then
                    nMax = nMax + 1
                    grdTramite.AddItem nMax & Chr(9) & !Codigo & Chr(9) & SubNull(!descricao)
 '               End If
               .MoveNext
            Loop
           .Close
        End With
        GravaMovimentoCC
    Else
        Sql = "SELECT tramitacaocc.seq, tramitacaocc.ccusto, CENTROCUSTO.DESCRICAO "
        Sql = Sql & "FROM tramitacaocc INNER JOIN CENTROCUSTO ON tramitacaocc.ccusto = CENTROCUSTO.CODIGO "
        Sql = Sql & "Where tramitacaocc.ano = " & lblAno.Caption & " And tramitacaocc.Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        Sql = Sql & " order by TRAMITACAOCC.SEQ"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                grdTramite.AddItem !Seq & Chr(9) & !ccusto & Chr(9) & SubNull(!descricao)
               .MoveNext
            Loop
           .Close
        End With
    End If
   .Close
End With

'VERIFICA OS TRAMITES CONCLUIDOS
If Val(frmProcesso.lblNumProc.Caption) > 0 Then
    For x = 1 To grdTramite.Rows - 1
        Sql = "SELECT CCUSTO,DESCRICAO,DATAHORA,NOMECOMPLETO,DESCDESPACHO,dataenvio,nomelogin2 FROM vwTRAMITACAO2 WHERE ANO=" & lblAno.Caption
        Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(x, 0)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                grdTramite.TextMatrix(x, 3) = Format(!DATAHORA, "dd/mm/yyyy")
                grdTramite.TextMatrix(x, 4) = Format(!DATAHORA, "hh:mm")
                grdTramite.TextMatrix(x, 5) = SubNull(!NomeCompleto)
                grdTramite.TextMatrix(x, 6) = SubNull(!DESCDESPACHO)
                If Not IsNull(!DATAENVIO) Then
                    grdTramite.TextMatrix(x, 8) = Format(!DATAENVIO, "dd/mm/yyyy")
                End If
                'grdTramite.TextMatrix(x, 9) = SubNull(!Usuario2)
                grdTramite.TextMatrix(x, 9) = SubNull(!nomelogin2)
            End If
           .Close
        End With
        Sql = "SELECT * FROM TRAMITACAO WHERE ANO=" & lblAno.Caption
        Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(x, 0)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        
        grdTramite.Row = x
        grdTramite.col = 0
        grdTramite.ColAlignment(1) = flexAlignRightCenter
        grdTramite.CellForeColor = vbWhite
        If Not IsNull(RdoAux!obs) And Not RdoAux.RowCount = 0 Then
            Set grdTramite.CellPicture = pic1.Image
            grdTramite.CellPictureAlignment = flexAlignLeftCenter
        End If
    Next
End If

CalculaDias
For x = 1 To grdTramite.Rows - 1
    If grdTramite.TextMatrix(x, 8) = "" Then
        grdTramite.Row = x
        Exit For
    End If
Next

'Marca trâmite atual
For x = 1 To grdTramite.Rows - 1
    If grdTramite.TextMatrix(x, 8) = "" Then
        If x < grdTramite.Rows - 1 Then
            If grdTramite.TextMatrix(x + 1, 8) <> "" Then
                GoTo ProximoX
            End If
        End If
        Exit For
    End If
ProximoX:
Next
If x = grdTramite.Rows Then
    x = x - 1
End If

For Y = 0 To grdTramite.Cols - 1
    grdTramite.Row = x
    grdTramite.col = Y
   ' grdTramite.CellBackColor = vbRed
   ' grdTramite.CellForeColor = vbWhite
   grdTramite.CellFontBold = True
Next

grdTramite.Row = 1
grdTramite.RowSel = 1
grdTramite.col = 0
grdTramite.ColSel = grdTramite.Cols - 1

bExec = True
grdTramite_RowColChange
End Sub

Private Sub Renumerar()
Dim x As Integer

With grdTramite
    For x = 1 To .Rows - 1
        .TextMatrix(x, 0) = x
    Next
End With

End Sub

Private Sub BloquearTela()

cmdAbaixo.Enabled = False
cmdAcima.Enabled = False
cmdInserir.Enabled = False
cmdRemover.Enabled = False
cmdEnviar.Enabled = False
cmdReceber.Enabled = False
cmdSair.Enabled = False
grdTramite.Enabled = False

End Sub

Private Sub LiberarTela()

cmdAbaixo.Enabled = True
cmdAcima.Enabled = True
cmdInserir.Enabled = True
If bRem Then
    cmdRemover.Enabled = True
 End If
 cmdEnviar.Enabled = True
cmdReceber.Enabled = True
cmdSair.Enabled = True
grdTramite.Enabled = True

End Sub

Private Sub Receber()

'Sql = "SELECT NOME,CODIGOCC FROM USUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Sql = "SELECT userid,CODIGOCC FROM USUARIOCC WHERE USERID=" & nLoginId & " AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Este usuário não tem permissão para receber processo deste local.", vbCritical, "Atenção"
    RdoAux.Close:    Exit Sub
End If

PnlRec.Visible = True
PnlRec.ZOrder 0
BloquearTela
PreencheListas
End Sub

Private Sub Enviar()
Dim x As Integer

'Sql = "SELECT NOME,CODIGOCC FROM USUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Sql = "SELECT userid,CODIGOCC FROM USUARIOCC WHERE USERID=" & nLoginId & " AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Este usuário não tem permissão para enviar processo deste local.", vbCritical, "Atenção"
    RdoAux.Close:    Exit Sub
End If


PnlEnv.Visible = True
PnlEnv.ZOrder 0
BloquearTela
PreencheListas
For x = 0 To cmbDespacho.ListCount - 1
    If cmbDespacho.List(x) = grdTramite.TextMatrix(grdTramite.Row, 6) Then
        cmbDespacho.ListIndex = x
        Exit For
    End If
Next
End Sub

Private Sub PreencheListas()
Dim sNome As String

cmbDespacho.Clear: cmbFunc.Clear: cmbDespacho2.Clear: cmbDespacho3.Clear
Sql = "SELECT CODIGO,DESCRICAO FROM DESPACHO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbDespacho.AddItem !descricao
        cmbDespacho.ItemData(cmbDespacho.NewIndex) = !Codigo
        cmbDespacho2.AddItem !descricao
        cmbDespacho2.ItemData(cmbDespacho2.NewIndex) = !Codigo
        cmbDespacho3.AddItem !descricao
        cmbDespacho3.ItemData(cmbDespacho3.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
If grdTramite.TextMatrix(grdTramite.Row, 5) = "" Then
    'Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE NOMELOGIN='" & NomeDeLogin & "'"
    Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE ID=" & nLoginId
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
       sNome = RdoAux!NomeCompleto
    Else
       sNome = ""
    End If
    RdoAux.Close
Else
    sNome = grdTramite.TextMatrix(grdTramite.Row, 5)
End If
cmbFunc.AddItem sNome
cmbFunc.ItemData(cmbFunc.NewIndex) = RetornaUsuarioID(RetornaUsuarioLoginName(sNome))
'Sql = "SELECT nomelogin, usuariofunc.funclogin, USUARIO.NOMECOMPLETO FROM usuariofunc INNER JOIN "
'Sql = Sql & "USUARIO ON usuariofunc.funclogin = USUARIO.NOMELOGIN "
'Sql = Sql & "WHERE     usuariofunc.userlogin = '" & NomeDeLogin & "'"
Sql = "SELECT nomelogin, usuariofunc.funclogin, USUARIO.NOMECOMPLETO FROM usuariofunc INNER JOIN "
Sql = Sql & "USUARIO ON usuariofunc.funclogin = USUARIO.NOMELOGIN "
Sql = Sql & "WHERE     usuariofunc.userid = '" & RetornaUsuarioID(NomeDeLogin) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFunc.AddItem !NomeCompleto
        'cmbFunc.ItemData(cmbFunc.NewIndex) = Val(Right$(!NomeLogin, 3))
        cmbFunc.ItemData(cmbFunc.NewIndex) = RetornaUsuarioID(RetornaUsuarioLoginName(!NomeCompleto))
       .MoveNext
    Loop
   .Close
End With

For x = 0 To cmbFunc.ListCount - 1
    If cmbFunc.List(x) = sNome Then
        cmbFunc.ListIndex = x
        Exit For
    End If
Next
lblFunc.Caption = sNome
lblData.Caption = Right$(frmMdi.Sbar.Panels(6).Text, 10) & " " & Format(Now, "hh:mm")
lblLocal.Caption = grdTramite.TextMatrix(grdTramite.Row, 2)
lblData2.Caption = Right$(frmMdi.Sbar.Panels(6).Text, 10) & " " & Format(Now, "hh:mm")
lblLocal2.Caption = grdTramite.TextMatrix(grdTramite.Row, 2)

End Sub

Private Sub Form_Unload(Cancel As Integer)
frmProcesso.Enabled = True
End Sub

Private Sub CalculaDias()
Dim dDataIni As Date, dDataFim As Date, t As Integer

With grdTramite
    For t = 1 To .Rows - 2
        If .TextMatrix(t + 1, 3) <> "" Then
            dDataIni = CDate(.TextMatrix(t, 3))
            dDataFim = CDate(.TextMatrix(t + 1, 3))
           .TextMatrix(t, 7) = CStr(DateDiff("d", dDataIni, dDataFim))
        Else
            Exit For
        End If
    Next
    If .TextMatrix(.Rows - 1, 3) <> "" Then
       .TextMatrix(.Rows - 1, 7) = 0
    End If
End With

End Sub

Private Sub GravaMovimentoCC()
If Val(lblNumProc.Caption) = 0 Then
    MsgBox "Sem numero de processo.", vbCritica, "ERRO"
    Exit Sub
End If

Sql = "DELETE FROM TRAMITACAOCC WHERE ANO=" & lblAno.Caption & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
cn.Execute Sql, rdExecDirect

With grdTramite
    For x = 1 To .Rows - 1
        Sql = "INSERT TRAMITACAOCC(ANO,NUMERO,SEQ,CCUSTO) VALUES("
        Sql = Sql & lblAno.Caption & "," & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & ","
        Sql = Sql & x & "," & .TextMatrix(x, 1) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

End Sub

Private Sub FormHagana()

evRem = 14

bRem = False
If InStr(1, sRet, Format(evRem, "000"), vbBinaryCompare) > 0 Then bRem = True

If Not bRem Then
    cmdRemover.Enabled = False
Else
    cmdRemover.Enabled = True
End If

End Sub

Private Sub grdTramite_RowColChange()
Dim RdoAux As rdoResultset, Sql As String
If Not bExec Then Exit Sub
txtObs.Text = ""
Sql = "SELECT OBS,obsinterna FROM TRAMITACAO WHERE ANO=" & lblAno.Caption
Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(grdTramite.Row, 0)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtObs.Text = SubNull(!obs)
        txtObs2.Text = SubNull(!obsinterna)
    End If
   .Close
End With

End Sub

Private Sub Label3_Click()

End Sub

Private Sub lblObsGeral_Click()
txtObs.Visible = True
txtObs2.Visible = False
lblTitObs.Caption = "Observação Geral"
End Sub

Private Sub lblObsInterna_Click()
txtObs.Visible = False
txtObs2.Visible = True
lblTitObs.Caption = "Observação Interna"
End Sub

Private Sub mnuObsGeral_Click()


End Sub


