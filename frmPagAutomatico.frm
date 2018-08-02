VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmPagAutomatico 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retorno de Arquivo Bancário / Baixa Automática"
   ClientHeight    =   5820
   ClientLeft      =   5625
   ClientTop       =   2610
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   8190
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   6765
      TabIndex        =   9
      ToolTipText     =   "Sair da Tela"
      Top             =   5370
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Sair"
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
      MICON           =   "frmPagAutomatico.frx":0000
      PICN            =   "frmPagAutomatico.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView Mv 
      Height          =   2310
      Left            =   60
      TabIndex        =   10
      Top             =   3210
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   15658734
      Appearance      =   0
      StartOfWeek     =   159907841
      TitleBackColor  =   192
      TitleForeColor  =   12648447
      CurrentDate     =   37439
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   8
      Left            =   5460
      TabIndex        =   8
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   1980
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":008A
      PICN            =   "frmPagAutomatico.frx":03A4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   7
      Left            =   5460
      TabIndex        =   7
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   990
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":0DC4
      PICN            =   "frmPagAutomatico.frx":10DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   6
      Left            =   5460
      TabIndex        =   6
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":1E0F
      PICN            =   "frmPagAutomatico.frx":2129
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   2
      Left            =   0
      TabIndex        =   2
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   1980
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BTYPE           =   9
      TX              =   "Outros Bancos"
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":29D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":2CF0
      PICN            =   "frmPagAutomatico.frx":300A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   1
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   990
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":4455
      PICN            =   "frmPagAutomatico.frx":476F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   3
      Left            =   2730
      TabIndex        =   3
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   0
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":548C
      PICN            =   "frmPagAutomatico.frx":57A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   4
      Left            =   2730
      TabIndex        =   4
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   990
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":6862
      PICN            =   "frmPagAutomatico.frx":6B7C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   5
      Left            =   2730
      TabIndex        =   5
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   1980
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmPagAutomatico.frx":75EB
      PICN            =   "frmPagAutomatico.frx":7905
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblAux 
      Caption         =   "Label3"
      Height          =   375
      Left            =   6060
      TabIndex        =   42
      Top             =   5790
      Width           =   1485
   End
   Begin VB.Label lblData 
      Caption         =   "Label3"
      Height          =   225
      Left            =   6120
      TabIndex        =   41
      Top             =   4380
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Disponíveis"
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
      Left            =   4125
      TabIndex        =   40
      Top             =   3165
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banespa..............:"
      Height          =   225
      Index           =   0
      Left            =   2730
      TabIndex        =   39
      Top             =   3450
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco do Brasil...:"
      Height          =   225
      Index           =   1
      Left            =   2730
      TabIndex        =   38
      Top             =   3690
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Outros Bancos.....:"
      Height          =   225
      Index           =   2
      Left            =   2730
      TabIndex        =   37
      Top             =   3915
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bradesco.............:"
      Height          =   225
      Index           =   3
      Left            =   2730
      TabIndex        =   36
      Top             =   4155
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Caixa Federal.......:"
      Height          =   225
      Index           =   4
      Left            =   2730
      TabIndex        =   35
      Top             =   4380
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HSBC..................:"
      Height          =   225
      Index           =   5
      Left            =   2730
      TabIndex        =   34
      Top             =   4620
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Itau......................:"
      Height          =   225
      Index           =   6
      Left            =   2730
      TabIndex        =   33
      Top             =   4860
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nossa Caixa........:"
      Height          =   225
      Index           =   7
      Left            =   2730
      TabIndex        =   32
      Top             =   5085
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unibanco............:"
      Height          =   225
      Index           =   8
      Left            =   2730
      TabIndex        =   31
      Top             =   5325
      Width           =   1365
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   4140
      TabIndex        =   30
      Top             =   3450
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   4140
      TabIndex        =   29
      Top             =   3690
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   2
      Left            =   4140
      TabIndex        =   28
      Top             =   3915
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   4140
      TabIndex        =   27
      Top             =   4155
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   4140
      TabIndex        =   26
      Top             =   4380
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   5
      Left            =   4140
      TabIndex        =   25
      Top             =   4620
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   6
      Left            =   4140
      TabIndex        =   24
      Top             =   4860
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   7
      Left            =   4140
      TabIndex        =   23
      Top             =   5085
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   8
      Left            =   4140
      TabIndex        =   22
      Top             =   5325
      Width           =   315
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione a Data de Geração dos Arquivos"
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
      Height          =   465
      Left            =   5910
      TabIndex        =   21
      Top             =   3675
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Baixados"
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
      Left            =   5115
      TabIndex        =   20
      Top             =   3165
      Width           =   690
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   5145
      TabIndex        =   19
      Top             =   3450
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   5145
      TabIndex        =   18
      Top             =   3690
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   2
      Left            =   5145
      TabIndex        =   17
      Top             =   3915
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   5145
      TabIndex        =   16
      Top             =   4155
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   5145
      TabIndex        =   15
      Top             =   4380
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   5
      Left            =   5145
      TabIndex        =   14
      Top             =   4620
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   6
      Left            =   5145
      TabIndex        =   13
      Top             =   4860
      Width           =   315
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   7
      Left            =   5145
      TabIndex        =   12
      Top             =   5085
      Width           =   360
   End
   Begin VB.Label lblBaixa 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   8
      Left            =   5145
      TabIndex        =   11
      Top             =   5325
      Width           =   315
   End
End
Attribute VB_Name = "frmPagAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FebrabanA
    CodigoRegistro As String * 1
    CodigoRemessa As String * 1
    CodigoConvenio As String * 20
    NomeEmpresa As String * 20
    CodigoBanco As String * 3
    NomeBanco As String * 20
    DataGeracao As String * 8
    NumeroSeq As String * 6
    VersaoLayout As String * 2
    Filler As String * 69
End Type

Private Type FebrabanG
   CodigoRegistro As String * 1
   ContaPrefeitura As String * 20
   DataPagamento As String * 8
   DataCredito As String * 8
   PreCodBarra As String * 4
   ValorRecebido As String * 11
   CodigoMunic As String * 4
   DataVencto As String * 8
   NumDocumento As String * 9
   NumParcela As String * 2
   SituacaoRetorno As String * 2
   FillerSmar As String * 4
   ValorRetornado As String * 12
   ValorTarifa As String * 7
   NumSeq As String * 8
   CodAgencia As String * 8
   FormaPagamento As String * 1
   NumAutentica As String * 23
   Filler As String * 10
End Type

Private Sub cmdBanco_Click(Index As Integer)
Dim RdoAux As rdoResultset, nCodBanco As Integer

lblAux.Caption = Index
'If NomeDoComputador <> "SCORPION" Then
'    If Val(lblBanco(Index).Caption) = 0 And Val(lblBaixa(Index).Caption) = 0 Then
'        MsgBox "Não existem arquivos para este " & vbCrLf & "banco na data especificada.", vbInformation, "Atenção"
'        Exit Sub
'    End If
'    frmPagBanco.show vbModeless, frmMdi
'
'    Select Case Index
'        Case 0
'            frmPagBanco.lblBanco.Caption = "033 - BANESPA"
'        Case 1
'            frmPagBanco.lblBanco.Caption = "001 - B.BRASIL"
'        Case 2
'            frmPagBanco.lblBanco.Caption = "641 - BBV BANCO"
'        Case 3
'            frmPagBanco.lblBanco.Caption = "237 - BRADESCO"
'        Case 4
'            frmPagBanco.lblBanco.Caption = "104 - CAIXA FED"
'        Case 5
''            frmPagBanco.lblBanco.Caption = "399 - HSBC"
''        Case 6
''            frmPagBanco.lblBanco.Caption = "341 - ITAU"
''        Case 7
''            frmPagBanco.lblBanco.Caption = "151 - N.CAIXA"
''        Case 8
''            frmPagBanco.lblBanco.Caption = "409 - UNIBANCO"
''    End Select
''
''    frmPagBanco.grdArq.Rows = 1
''
''    Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(lblData.Caption, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(frmPagBanco.lblBanco.Caption, 3))
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        Do Until .EOF
'           frmPagBanco.grdArq.AddItem sPathArqBanco & "\" & Mv.Year & "\" & Format(Mv.Month, "00") & "\" & Format(Mv.Day, "00") & "\" & Chr(9) & !NOMEARQ
'          .MoveNext
'        Loop
'    End With
'Else
    frmBaixaBancaria.show vbModeless
'    If frmMdi.frTeste.Visible = True Then
'        If frmMdi.frTeste.Caption = "ACESSANDO OS DADOS LOCAIS" Then
'            sPathArqBanco = "C:\Trabalho\GTI\Bancos"
'        End If
'    End If
    If frmBaixaBancaria.lblBanco.Caption = "000-OUTROS BANCOS" Then
        nCodBanco = 90
    Else
        nCodBanco = Val(Left$(frmBaixaBancaria.lblBanco.Caption, 3))
    End If
    If nCodBanco > 0 Then
        frmBaixaBancaria.txtPath.Text = sPathArqBanco & "\" & Mv.Year & "\" & Format(Mv.Month, "00") & "\" & Format(Mv.Day, "00") & "\"
        lblData.Caption = Format(Mv.Day, "00") & "\" & Format(Mv.Month, "00") & "\" & Mv.Year
        If nCodBanco = 33 Then
            Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(Mv.Month, "00") & "/" & Format(Mv.Day, "00") & "/" & Mv.Year & "' AND (CODBANCO=33 or codbanco=8)"
        ElseIf nCodBanco = 90 Then
            Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(Mv.Month, "00") & "/" & Format(Mv.Day, "00") & "/" & Mv.Year & "' AND CODBANCO NOT IN (1,8,33,104,151,237,341,399,409,641)"
        Else
            Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(Mv.Month, "00") & "/" & Format(Mv.Day, "00") & "/" & Mv.Year & "' AND CODBANCO=" & nCodBanco
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
               frmBaixaBancaria.lstArq.AddItem !NOMEARQ
              .MoveNext
            Loop
           .Close
        End With
        If frmBaixaBancaria.lstArq.ListCount > 0 Then frmBaixaBancaria.lstArq.ListIndex = 0
        frmBaixaBancaria.ZOrder 0
    Else
        Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(Mv.Month, "00") & "/" & Format(Mv.Day, "00") & "/" & Mv.Year & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux
             Do Until .EOF
                frmBaixaBancaria.lstArq.AddItem RdoAux!NOMEARQ
               .MoveNext
             Loop
        End With
        
    
    
    
    End If
End Sub

Private Sub cmdSair_Click()

Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long
Ocupado
Sql = "SELECT DISTINCT  debitopago.codreduzido, debitopago.contacorrente, debitoparcela.statuslanc, debitopago.anoexercicio, debitopago.seqlancamento, debitopago.numparcela,"
Sql = Sql & "debitopago.CODCOMPLEMENTO , debitopago.CodLancamento FROM  debitopago INNER JOIN debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
Sql = Sql & "debitopago.NumParcela = debitoparcela.NumParcela And debitopago.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO Where (debitoparcela.statuslanc = 3)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        Sql = "update debitoparcela set statuslanc=2 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, statuslanc From debitoparcela "
Sql = Sql & "Where (statuslanc = 1) And (NumParcela = 0) And (CODREDUZIDO < 100000) And (AnoExercicio = 2018) And (CodLancamento = 1)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        Sql = "update debitoparcela set statuslanc=5 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=0  and statuslanc=3"
        cn.Execute Sql, rdExecDirect
        DoEvents
       .MoveNext
    Loop
   .Close
End With


Liberado



Unload Me
End Sub

Private Sub Form_Activate()
Mv_DateClick Now
End Sub

Private Sub Form_Load()
If frmMdi.frTeste.Visible = False Then
    frmServico.show 1
    frmServico.btVerificar_Click
    Unload frmServico
End If

Screen.MousePointer = vbHourglass
Centraliza Me
'txtAno.text = Year(Now)
Mv.Day = Day(Now)
Mv.Year = Year(Now)
Mv.Month = Month(Now)

Screen.MousePointer = vbDefault
End Sub

Private Sub Mv_Click()
'txtAno.text = Mv.Year
End Sub

Private Sub Mv_DateClick(ByVal DateClicked As Date)
Dim RdoAux  As rdoResultset
lblData.Caption = Format(Mv.value, "dd/mm/yyyy")
Screen.MousePointer = vbHourglass
lblMsg.ForeColor = vbRed
lblMsg.Caption = "Aguarde... Lendo Arquivos."
lblMsg.Refresh
LimpaContador
If Len(Trim$(CStr(Mv.Year))) = 4 Then
   sAno = CStr(Mv.Year)
Else
   sAno = "20" & CStr(Mv.Year)
End If

Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(lblData.Caption, "mm/dd/yyyy") & "' "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Select Case !CodBanco
            Case 1 'BB
                If IsNull(!DATABAIXA) Then
                   lblBanco(1).Caption = Val(lblBanco(1).Caption) + 1
                Else
                   lblBaixa(1).Caption = Val(lblBaixa(1).Caption) + 1
                End If
            Case 8, 33 'BANESPA
                If IsNull(!DATABAIXA) Then
                   lblBanco(0).Caption = Val(lblBanco(0).Caption) + 1
                Else
                   lblBaixa(0).Caption = Val(lblBaixa(0).Caption) + 1
                End If
            Case 104 'CAIXA FEDERAL
                If IsNull(!DATABAIXA) Then
                   lblBanco(4).Caption = Val(lblBanco(4).Caption) + 1
                Else
                   lblBaixa(4).Caption = Val(lblBaixa(4).Caption) + 1
                End If
            Case 151 'NOSSA CAIXA
                If IsNull(!DATABAIXA) Then
                   lblBanco(7).Caption = Val(lblBanco(7).Caption) + 1
                Else
                   lblBaixa(7).Caption = Val(lblBaixa(7).Caption) + 1
                End If
            Case 237 'BRADESCO
                If IsNull(!DATABAIXA) Then
                   lblBanco(3).Caption = Val(lblBanco(3).Caption) + 1
                Else
                   lblBaixa(3).Caption = Val(lblBaixa(3).Caption) + 1
                End If
            Case 341 'ITAU
                If IsNull(!DATABAIXA) Then
                   lblBanco(6).Caption = Val(lblBanco(6).Caption) + 1
                Else
                   lblBaixa(6).Caption = Val(lblBaixa(6).Caption) + 1
                End If
            Case 399 'HSBC
                If IsNull(!DATABAIXA) Then
                   lblBanco(5).Caption = Val(lblBanco(5).Caption) + 1
                Else
                   lblBaixa(5).Caption = Val(lblBaixa(5).Caption) + 1
                End If
            Case 409 'UNIBANCO
                If IsNull(!DATABAIXA) Then
                   lblBanco(8).Caption = Val(lblBanco(8).Caption) + 1
                Else
                   lblBaixa(8).Caption = Val(lblBaixa(8).Caption) + 1
                End If
            Case Else 'Outros bancos
                If IsNull(!DATABAIXA) Then
                   lblBanco(2).Caption = Val(lblBanco(2).Caption) + 1
                Else
                   lblBaixa(2).Caption = Val(lblBaixa(2).Caption) + 1
                End If
       End Select
      .MoveNext
    Loop
End With

Screen.MousePointer = vbDefault
lblMsg.ForeColor = &HC00000
lblMsg.Caption = "Selecione a Data de Geração dos Arquivos"
lblMsg.Refresh
   
End Sub

Private Sub LimpaContador()
For x = 0 To 8
    lblBanco(x).Caption = 0
    lblBaixa(x).Caption = 0
Next
End Sub

