VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCPTramite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tramitação de Processos (Setor de Compras)"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   10755
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
      Left            =   315
      TabIndex        =   39
      Top             =   900
      Visible         =   0   'False
      Width           =   9585
      Begin MSFlexGridLib.MSFlexGrid grdEmp 
         Height          =   2475
         Left            =   90
         TabIndex        =   40
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
         FormatString    =   $"frmCPTramite.frx":0000
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   3855
      TabIndex        =   30
      Top             =   4290
      Width           =   1185
   End
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
      Left            =   2235
      TabIndex        =   26
      Top             =   1650
      Visible         =   0   'False
      Width           =   6285
      Begin VB.ComboBox cmbLocal 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   330
         Width           =   5565
      End
      Begin prjChameleon.chameleonButton cmdOK 
         Height          =   345
         Left            =   4950
         TabIndex        =   28
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
         MICON           =   "frmCPTramite.frx":00B4
         PICN            =   "frmCPTramite.frx":00D0
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
         TabIndex        =   29
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
         MICON           =   "frmCPTramite.frx":022A
         PICN            =   "frmCPTramite.frx":0246
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
      Left            =   2955
      TabIndex        =   15
      Top             =   1020
      Visible         =   0   'False
      Width           =   6735
      Begin VB.ComboBox cmbDespacho 
         Height          =   315
         Left            =   1485
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1425
         Width           =   5010
      End
      Begin prjChameleon.chameleonButton cmdOk3 
         Height          =   345
         Left            =   5775
         TabIndex        =   17
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
         MICON           =   "frmCPTramite.frx":03A0
         PICN            =   "frmCPTramite.frx":03BC
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
         TabIndex        =   18
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
         MICON           =   "frmCPTramite.frx":0516
         PICN            =   "frmCPTramite.frx":0532
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1515
         TabIndex        =   25
         Top             =   1110
         Width           =   4905
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Despacho..........:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   1500
         Width           =   1365
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário........:"
         Height          =   225
         Left            =   180
         TabIndex        =   23
         Top             =   1155
         Width           =   1365
      End
      Begin VB.Label lblData2 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1515
         TabIndex        =   22
         Top             =   795
         Width           =   2520
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Data e Hora.......:"
         Height          =   225
         Left            =   180
         TabIndex        =   21
         Top             =   795
         Width           =   1365
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Local..................:"
         Height          =   225
         Left            =   180
         TabIndex        =   20
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label lblLocal2 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1515
         TabIndex        =   19
         Top             =   450
         Width           =   4890
      End
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
      Left            =   2955
      TabIndex        =   4
      Top             =   1050
      Visible         =   0   'False
      Width           =   6705
      Begin VB.ComboBox cmbDespacho2 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1410
         Width           =   5010
      End
      Begin VB.ComboBox cmbFunc 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   5010
      End
      Begin prjChameleon.chameleonButton cmdOK2 
         Height          =   345
         Left            =   5745
         TabIndex        =   7
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
         MICON           =   "frmCPTramite.frx":068C
         PICN            =   "frmCPTramite.frx":06A8
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
         TabIndex        =   8
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
         MICON           =   "frmCPTramite.frx":0802
         PICN            =   "frmCPTramite.frx":081E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Despacho..........:"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   1440
         Width           =   1365
      End
      Begin VB.Label lblLocal 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1530
         TabIndex        =   13
         Top             =   390
         Width           =   4890
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Local..................:"
         Height          =   225
         Left            =   195
         TabIndex        =   12
         Top             =   390
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Data e Hora.......:"
         Height          =   225
         Left            =   195
         TabIndex        =   11
         Top             =   735
         Width           =   1365
      End
      Begin VB.Label lblData 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1530
         TabIndex        =   10
         Top             =   735
         Width           =   2520
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Funcionário........:"
         Height          =   225
         Left            =   195
         TabIndex        =   9
         Top             =   1095
         Width           =   1365
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
      Left            =   2205
      TabIndex        =   0
      Top             =   1710
      Visible         =   0   'False
      Width           =   6285
      Begin VB.ComboBox cmbDespacho3 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   330
         Width           =   5565
      End
      Begin prjChameleon.chameleonButton cmdOK4 
         Height          =   345
         Left            =   4950
         TabIndex        =   2
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
         MICON           =   "frmCPTramite.frx":0978
         PICN            =   "frmCPTramite.frx":0994
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
         TabIndex        =   3
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
         MICON           =   "frmCPTramite.frx":0AEE
         PICN            =   "frmCPTramite.frx":0B0A
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
   Begin MSFlexGridLib.MSFlexGrid grdTramite 
      Height          =   2475
      Left            =   45
      TabIndex        =   31
      Top             =   990
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   4366
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmCPTramite.frx":0C64
   End
   Begin prjChameleon.chameleonButton cmdInserir 
      Height          =   345
      Left            =   75
      TabIndex        =   32
      ToolTipText     =   "Inserir um novo local para a tramitação"
      Top             =   3660
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
      MICON           =   "frmCPTramite.frx":0D22
      PICN            =   "frmCPTramite.frx":0D3E
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
      Left            =   1695
      TabIndex        =   33
      ToolTipText     =   "Remover um local de tramitação"
      Top             =   3660
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
      MICON           =   "frmCPTramite.frx":0E98
      PICN            =   "frmCPTramite.frx":0EB4
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
      Left            =   6945
      TabIndex        =   34
      ToolTipText     =   "Receber um Processo"
      Top             =   3660
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCPTramite.frx":100E
      PICN            =   "frmCPTramite.frx":102A
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
      Left            =   9465
      TabIndex        =   35
      ToolTipText     =   "Sair da Tela"
      Top             =   3660
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
      MICON           =   "frmCPTramite.frx":1204
      PICN            =   "frmCPTramite.frx":1220
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
      Left            =   3315
      TabIndex        =   36
      ToolTipText     =   "Mover um local acima"
      Top             =   3660
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
      MICON           =   "frmCPTramite.frx":137A
      PICN            =   "frmCPTramite.frx":1396
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
      Left            =   3705
      TabIndex        =   37
      ToolTipText     =   "Mover um local abaixo"
      Top             =   3660
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
      MICON           =   "frmCPTramite.frx":14F0
      PICN            =   "frmCPTramite.frx":150C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEnviar 
      Height          =   345
      Left            =   8205
      TabIndex        =   38
      ToolTipText     =   "Enviar um Processo"
      Top             =   3660
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCPTramite.frx":1666
      PICN            =   "frmCPTramite.frx":1682
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
      Left            =   5505
      TabIndex        =   41
      ToolTipText     =   "Alterar o despacho de processo aberto"
      Top             =   3660
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
      MICON           =   "frmCPTramite.frx":16F6
      PICN            =   "frmCPTramite.frx":1712
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
      Caption         =   "Nº do Processo...:"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   49
      Top             =   90
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   225
      Index           =   1
      Left            =   2505
      TabIndex        =   48
      Top             =   90
      Width           =   435
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Assunto...............:"
      Height          =   225
      Index           =   6
      Left            =   105
      TabIndex        =   47
      Top             =   390
      Width           =   1365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requerente.........:"
      Height          =   225
      Index           =   7
      Left            =   105
      TabIndex        =   46
      Top             =   690
      Width           =   1365
   End
   Begin VB.Label lblNumProc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1515
      TabIndex        =   45
      Top             =   90
      Width           =   915
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2985
      TabIndex        =   44
      Top             =   90
      Width           =   705
   End
   Begin VB.Label lblAssunto 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1515
      TabIndex        =   43
      Top             =   390
      Width           =   6495
   End
   Begin VB.Label lblRequerente 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1515
      TabIndex        =   42
      Top             =   690
      Width           =   6495
   End
End
Attribute VB_Name = "frmCPTramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim evRem As Integer, bRem As Boolean, sRet As String, sLogin As String

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

If IsDate(frmCPProcesso.lblDtArquivamento) Then
    MsgBox "O Processo está arquivado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmCPProcesso.lblDtCancelamento) Then
    MsgBox "O Processo está Cancelado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmCPProcesso.lblDtSuspencao) Then
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


Sql = "SELECT NOME,CODIGOCC FROM CPUSUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
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

Private Sub cmdEnviar_Click()

If IsDate(frmCPProcesso.lblDtArquivamento) Then
    MsgBox "O Processo está arquivado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmCPProcesso.lblDtCancelamento) Then
    MsgBox "O Processo está Cancelado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmCPProcesso.lblDtSuspencao) Then
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
                Enviar
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

Private Sub cmdInserir_Click()

If IsDate(frmCPProcesso.lblDtArquivamento.Caption) Then
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
Sql = "SELECT CODIGO, DESCRICAO FROM CPCENTROCUSTO WHERE ATIVO = 1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        cmbLocal.AddItem !DESCRICAO
        cmbLocal.ItemData(cmbLocal.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
If cmbLocal.ListCount > 0 Then cmbLocal.ListIndex = 0
BloquearTela
PnlInserir.Visible = True
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

    Sql = "DELETE FROM CPTRAMITACAO WHERE ANO=" & lblAno.Caption
    Sql = Sql & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
    Sql = Sql & " AND SEQ=" & .TextMatrix(.Row, 0)
    cn.Execute Sql, rdExecDirect

    Sql = "SELECT NOMELOGIN FROM USUARIO WHERE NOMECOMPLETO='" & cmbFunc.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            sLogin = !NomeLogin
        End If
       .Close
    End With

    If cmbDespacho2.ListIndex > -1 Then
        Sql = "INSERT CPTRAMITACAO (ANO,NUMERO,SEQ,CCUSTO,DATAHORA,DESPACHO,USUARIO) VALUES("
        Sql = Sql & lblAno.Caption & "," & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & ","
        Sql = Sql & .TextMatrix(.Row, 0) & "," & .TextMatrix(.Row, 1) & ",'" & Format(lblData.Caption, "mm/dd/yyyy hh:mm") & "',"
        Sql = Sql & cmbDespacho2.ItemData(cmbDespacho2.ListIndex) & ",'" & Trim$(sLogin) & "')"
     Else
        Sql = "INSERT CPTRAMITACAO (ANO,NUMERO,SEQ,CCUSTO,DATAHORA,USUARIO) VALUES("
        Sql = Sql & lblAno.Caption & "," & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & ","
        Sql = Sql & .TextMatrix(.Row, 0) & "," & .TextMatrix(.Row, 1) & ",'" & Format(lblData.Caption, "mm/dd/yyyy hh:mm") & "','" & Trim$(sLogin) & "')"
    End If
    cn.Execute Sql, rdExecDirect

End With

CalculaDias

End Sub

Private Sub cmdOk3_Click()

If cmbDespacho.ListIndex = -1 Then
    MsgBox "Selecione o despacho.", vbExclamation, "Atenção"
    Exit Sub
End If

grdTramite.TextMatrix(grdTramite.Row, 6) = cmbDespacho.Text
grdTramite.TextMatrix(grdTramite.Row, 8) = Format(Now, "dd/mm/yyyy")

Sql = "SELECT NOMELOGIN FROM USUARIO WHERE NOMECOMPLETO='" & lblFunc.Caption & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        sLogin = !NomeLogin
    End If
   .Close
End With

Sql = "UPDATE CPTRAMITACAO SET DESPACHO=" & cmbDespacho.ItemData(cmbDespacho.ListIndex)
Sql = Sql & " ,USUARIO='" & Trim$(sLogin) & "'  ,DATAENVIO='" & Format(Now, "mm/dd/yyyy") & "' "
Sql = Sql & "WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(grdTramite.Row, 0)
cn.Execute Sql, rdExecDirect
CalculaDias
PnlEnv.Visible = False
LiberarTela

End Sub

Private Sub cmdOK4_Click()
If cmbDespacho3.ListIndex = -1 Then Exit Sub
Sql = "UPDATE CPTRAMITACAO SET DESPACHO=" & cmbDespacho.ItemData(cmbDespacho3.ListIndex)
Sql = Sql & " WHERE ANO=" & Val(lblAno.Caption) & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Sql = Sql & " AND SEQ=" & grdTramite.TextMatrix(grdTramite.Row, 0)
cn.Execute Sql, rdExecDirect
grdTramite.TextMatrix(grdTramite.Row, 6) = cmbDespacho3.Text
pnlDespacho.Visible = False

End Sub

Private Sub cmdReceber_Click()

If IsDate(frmCPProcesso.lblDtArquivamento) Then
    MsgBox "O Processo está arquivado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmCPProcesso.lblDtCancelamento) Then
    MsgBox "O Processo está Cancelado e não pode ser tramitado.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(frmCPProcesso.lblDtSuspencao) Then
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
Sql = "SELECT    distinct  cptramitacao.ano, cptramitacao.numero, cptramitacaocc.seq, cptramitacaocc.ccusto "
Sql = Sql & "FROM cptramitacao LEFT OUTER JOIN cptramitacaocc ON cptramitacao.ano = cptramitacaocc.ano AND cptramitacao.numero = cptramitacaocc.numero "
Sql = Sql & "WHERE     (cptramitacao.ano > 2004) AND (cptramitacaocc.ccusto IS NULL)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cGetInputState() <> 0 Then DoEvents
'        Sql = "SELECT ANO,NUMERO FROM TRAMITACAOCC WHERE ANO=" & !ANO & " AND NUMERO=" & !Numero
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux2.RowCount = 0 Then
            Sql = "INSERT CPTRAMITACAOCC SELECT ANO,NUMERO,SEQ,CCUSTO FROM TRAMITACAO WHERE ANO=" & !Ano & " AND NUMERO=" & !Numero
            cn.Execute Sql, rdExecDirect
'        End If
'        RdoAux2.Close
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Form_Activate()
frmCPProcesso.Enabled = False
End Sub

Private Sub Form_Deactivate()
Me.ZOrder 0
End Sub

Private Sub Form_Load()
Ocupado
Centraliza Me
sRet = RetEventUserForm(Me.Name)
grdTramite.ColWidth(1) = 0
lblNumProc.Caption = frmCPProcesso.lblNumProc.Caption
lblAno.Caption = frmCPProcesso.lblAno.Caption
lblAssunto.Caption = frmCPProcesso.cmbAssunto.Text
lblRequerente.Caption = frmCPProcesso.lblNomeCid.Caption
CarregaTramite
FormHagana
Liberado
End Sub

Private Sub CarregaTramite()
Dim bAchou As Boolean, nMax As Integer
If Val(lblNumProc.Caption) = 0 Then
    MsgBox "Sem numero de processo.", vbCritica, "ERRO"
    Exit Sub
End If
'CARREGA TODOS OS TRAMITES
Sql = "SELECT * FROM cptramitacaocc Where ano = " & lblAno.Caption & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        Sql = "SELECT ano, numero, seq, ccusto, DESCRICAO From vwcptramitacao Where ano =" & lblAno.Caption & " And Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2)) & " order by seq"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                grdTramite.AddItem !Seq & Chr(9) & !ccusto & Chr(9) & SubNull(!DESCRICAO)
                nMax = !Seq
               .MoveNext
            Loop
           .Close
        End With
        
    
        Sql = "SELECT CPASSUNTOCC.SEQ,CPCENTROCUSTO.CODIGO, CPCENTROCUSTO.DESCRICAO FROM CPASSUNTOCC INNER JOIN "
        Sql = Sql & "CPCENTROCUSTO ON CPASSUNTOCC.CODCC = CPCENTROCUSTO.CODIGO "
        Sql = Sql & "WHERE CPASSUNTOCC.CODASSUNTO =" & frmCPProcesso.cmbAssunto.ItemData(frmCPProcesso.cmbAssunto.ListIndex)
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
                    grdTramite.AddItem nMax & Chr(9) & !Codigo & Chr(9) & SubNull(!DESCRICAO)
 '               End If
               .MoveNext
            Loop
           .Close
        End With
        GravaMovimentoCC
    Else
        Sql = "SELECT cptramitacaocc.seq, cptramitacaocc.ccusto, cpCENTROCUSTO.DESCRICAO "
        Sql = Sql & "FROM cptramitacaocc INNER JOIN cpCENTROCUSTO ON cptramitacaocc.ccusto = cpCENTROCUSTO.CODIGO "
        Sql = Sql & "Where cptramitacaocc.ano = " & lblAno.Caption & " And cptramitacaocc.Numero = " & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
        Sql = Sql & " order by cpTRAMITACAOCC.SEQ"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                grdTramite.AddItem !Seq & Chr(9) & !ccusto & Chr(9) & SubNull(!DESCRICAO)
               .MoveNext
            Loop
           .Close
        End With
    End If
   .Close
End With

'VERIFICA OS TRAMITES CONCLUIDOS
If Val(frmCPProcesso.lblNumProc.Caption) > 0 Then
    For x = 1 To grdTramite.Rows - 1
        Sql = "SELECT CCUSTO,DESCRICAO,DATAHORA,NOMECOMPLETO,DESCDESPACHO,dataenvio FROM VWcpTRAMITACAO WHERE ANO=" & lblAno.Caption
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
            End If
           .Close
        End With
    Next
End If
CalculaDias

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

'Sql = "SELECT NOME,CODIGOCC FROM USUARIOCC WHERE NOME='" & Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Sql = "SELECT NOME,CODIGOCC FROM CPUSUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Este usuário não tem permissão para receber processo deste local.", vbCritical, "Atenção"
    RdoAux.Close:    Exit Sub
End If

PnlRec.Visible = True
BloquearTela
PreencheListas
End Sub

Private Sub Enviar()
Dim x As Integer

'Sql = "SELECT NOME,CODIGOCC FROM USUARIOCC WHERE NOME='" & Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Sql = "SELECT NOME,CODIGOCC FROM CPUSUARIOCC WHERE NOME='" & NomeDeLogin & "' AND CODIGOCC=" & grdTramite.TextMatrix(grdTramite.Row, 1)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Este usuário não tem permissão para receber processo deste local.", vbCritical, "Atenção"
    RdoAux.Close:    Exit Sub
End If


PnlEnv.Visible = True
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
Sql = "SELECT CODIGO,DESCRICAO FROM CPDESPACHO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbDespacho.AddItem !DESCRICAO
        cmbDespacho.ItemData(cmbDespacho.NewIndex) = !Codigo
        cmbDespacho2.AddItem !DESCRICAO
        cmbDespacho2.ItemData(cmbDespacho2.NewIndex) = !Codigo
        cmbDespacho3.AddItem !DESCRICAO
        cmbDespacho3.ItemData(cmbDespacho3.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
If grdTramite.TextMatrix(grdTramite.Row, 5) = "" Then
    Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE NOMELOGIN='" & Mid(frmMdi.Sbar.Panels(2).Text, 10, Len(frmMdi.Sbar.Panels(2).Text) - 8) & "'"
    Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE NOMELOGIN='" & NomeDeLogin & "'"
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
cmbFunc.ItemData(cmbFunc.NewIndex) = 999
Sql = "SELECT nomelogin, usuariofunc.funclogin, USUARIO.NOMECOMPLETO FROM usuariofunc INNER JOIN "
Sql = Sql & "USUARIO ON usuariofunc.funclogin = USUARIO.NOMELOGIN "
Sql = Sql & "WHERE     usuariofunc.userlogin = '" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFunc.AddItem !NomeCompleto
        cmbFunc.ItemData(cmbFunc.NewIndex) = Val(Right$(!NomeLogin, 3))
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
frmCPProcesso.Enabled = True
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

Sql = "DELETE FROM CPTRAMITACAOCC WHERE ANO=" & lblAno.Caption & " AND NUMERO=" & Val(Left$(lblNumProc.Caption, Len(lblNumProc.Caption) - 2))
cn.Execute Sql, rdExecDirect

With grdTramite
    For x = 1 To .Rows - 1
        Sql = "INSERT CPTRAMITACAOCC(ANO,NUMERO,SEQ,CCUSTO) VALUES("
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


