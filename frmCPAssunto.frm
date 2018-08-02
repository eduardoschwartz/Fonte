VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCPAssunto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de Assuntos (Divida Ativa)"
   ClientHeight    =   5955
   ClientLeft      =   1590
   ClientTop       =   1875
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   9210
   Begin VB.ListBox lstDoc2 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   4905
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   4290
      Width           =   4245
   End
   Begin VB.ListBox lstDoc1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   4290
      Width           =   4245
   End
   Begin VB.ListBox lstCC1 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   2430
      Width           =   4245
   End
   Begin VB.TextBox txtAssunto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   45
      MaxLength       =   150
      TabIndex        =   4
      Top             =   150
      Width           =   3795
   End
   Begin VB.ListBox lstAssunto 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   510
      Width           =   4245
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5565
      MaxLength       =   150
      TabIndex        =   2
      Top             =   1140
      Width           =   3435
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5565
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   795
      Width           =   840
   End
   Begin VB.CheckBox chkInativo 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Inativo"
      Height          =   255
      Left            =   7425
      TabIndex        =   0
      Top             =   795
      Visible         =   0   'False
      Width           =   1215
   End
   Begin prjChameleon.chameleonButton cmdDR1 
      Height          =   255
      Left            =   4395
      TabIndex        =   8
      ToolTipText     =   "Atualiza Lista"
      Top             =   5430
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   450
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
      MICON           =   "frmCPAssunto.frx":0000
      PICN            =   "frmCPAssunto.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdR2 
      Height          =   285
      Left            =   4395
      TabIndex        =   9
      ToolTipText     =   "Remove documento"
      Top             =   4890
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
      MICON           =   "frmCPAssunto.frx":00BB
      PICN            =   "frmCPAssunto.frx":00D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdR1 
      Height          =   285
      Left            =   4395
      TabIndex        =   10
      ToolTipText     =   "Adiciona documento"
      Top             =   4560
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
      MICON           =   "frmCPAssunto.frx":0231
      PICN            =   "frmCPAssunto.frx":024D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvCC2 
      Height          =   1590
      Left            =   4905
      TabIndex        =   11
      Top             =   2430
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   2805
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Descricao"
         Object.Width           =   7232
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdCR1 
      Height          =   255
      Left            =   4395
      TabIndex        =   12
      ToolTipText     =   "Atualiza Lista"
      Top             =   3600
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   450
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
      MICON           =   "frmCPAssunto.frx":03A7
      PICN            =   "frmCPAssunto.frx":03C3
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
      Left            =   4395
      TabIndex        =   13
      ToolTipText     =   "Remove centro de custos"
      Top             =   3060
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
      MICON           =   "frmCPAssunto.frx":0462
      PICN            =   "frmCPAssunto.frx":047E
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
      Left            =   4395
      TabIndex        =   14
      ToolTipText     =   "Adiciona centro de custos"
      Top             =   2730
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
      MICON           =   "frmCPAssunto.frx":05D8
      PICN            =   "frmCPAssunto.frx":05F4
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
      Height          =   300
      Left            =   3900
      TabIndex        =   15
      ToolTipText     =   "Consultar Assuntos por parte do nome"
      Top             =   150
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
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
      MICON           =   "frmCPAssunto.frx":074E
      PICN            =   "frmCPAssunto.frx":076A
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
      Left            =   4605
      TabIndex        =   16
      ToolTipText     =   "Novo Registro"
      Top             =   1605
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
      MICON           =   "frmCPAssunto.frx":08C4
      PICN            =   "frmCPAssunto.frx":08E0
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
      Left            =   5505
      TabIndex        =   17
      ToolTipText     =   "Editar Registro"
      Top             =   1605
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
      MICON           =   "frmCPAssunto.frx":0A3A
      PICN            =   "frmCPAssunto.frx":0A56
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
      Left            =   6405
      TabIndex        =   18
      ToolTipText     =   "Excluir Registro"
      Top             =   1605
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCPAssunto.frx":0BB0
      PICN            =   "frmCPAssunto.frx":0BCC
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
      Left            =   7305
      TabIndex        =   19
      ToolTipText     =   "Gravar os Dados"
      Top             =   1605
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
      MICON           =   "frmCPAssunto.frx":0C6E
      PICN            =   "frmCPAssunto.frx":0C8A
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
      Left            =   8235
      TabIndex        =   20
      ToolTipText     =   "Cancelar Edição"
      Top             =   1590
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
      MICON           =   "frmCPAssunto.frx":102F
      PICN            =   "frmCPAssunto.frx":104B
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
      Left            =   8235
      TabIndex        =   21
      ToolTipText     =   "Sair da Tela"
      Top             =   1590
      Width           =   885
      _ExtentX        =   1561
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
      MICON           =   "frmCPAssunto.frx":11A5
      PICN            =   "frmCPAssunto.frx":11C1
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
      Height          =   300
      Left            =   4350
      TabIndex        =   22
      ToolTipText     =   "Imprimir os documentos necessarios"
      Top             =   135
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   529
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
      MICON           =   "frmCPAssunto.frx":122F
      PICN            =   "frmCPAssunto.frx":124B
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
      BackColor       =   &H00000080&
      Caption         =   "Selecione os documentos exigidos para este assunto."
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
      Left            =   45
      TabIndex        =   26
      Top             =   4050
      Width           =   9105
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Selecione os centros de custo para a tramitação deste assunto."
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
      Left            =   45
      TabIndex        =   25
      Top             =   2190
      Width           =   9105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição...:"
      Height          =   195
      Index           =   11
      Left            =   4605
      TabIndex        =   24
      Top             =   1215
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código........:"
      Height          =   195
      Index           =   2
      Left            =   4605
      TabIndex        =   23
      Top             =   840
      Width           =   1155
   End
End
Attribute VB_Name = "frmCPAssunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset
Dim sEvento As String, z As Long, sDesc As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub cmdAlterar_Click()
Eventos "INCLUIR"
sEvento = "Alterar"
End Sub

Private Sub cmdC1_Click()
If lstCC1.ListIndex = -1 Then
    MsgBox "Selecione um centro de custos.", vbExclamation, "Atenção"
Else
    Set itmX = lvCC2.ListItems.Add(, "C" & Format(lvCC2.ListItems.Count + 1, "00") & Format(lstCC1.ItemData(lstCC1.ListIndex), "000"), lstCC1.Text)
End If

End Sub

Private Sub cmdC2_Click()

If lvCC2.SelectedItem = "" Then
    MsgBox "Selecione um centro de custos a remover.", vbExclamation, "Atenção"
Else
    lvCC2.ListItems.Remove (lvCC2.SelectedItem.Index)
    RefreshKey
End If

End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
sEvento = ""
lstAssunto_Click
End Sub

Private Sub cmdCR1_Click()
CarregaCC
End Sub

Private Sub cmdDR1_Click()
CarregaDoc
End Sub

Private Sub cmdGravar_Click()
Dim v As Integer
sDesc = UCase$(Trim$(txtDesc.Text))

If sDesc = "" Then
    MsgBox "Digite o nome do assunto.", vbExclamation, "Atenção"
    Exit Sub
End If

If sEvento = "Novo" Then
    For v = 0 To lstAssunto.ListCount - 1
        If lstAssunto.List(v) = sDesc Then
            MsgBox "Assunto já cadastrado.", vbExclamation, "Atenção"
            Exit Sub
        End If
    Next
End If

If lvCC2.ListItems.Count = 0 Then
    MsgBox "Inclua as secretarias na tramitação.", vbExclamation, "Atenção"
    Exit Sub
End If

Grava

Eventos "INICIAR"
sEvento = ""
End Sub

Private Sub cmdNovo_Click()
Eventos "INCLUIR"
sEvento = "Novo"
Limpa
CarregaLista
End Sub

Private Sub cmdPrint_Click()
frmReport.ShowReport "DOCUMENTOASSUNTO", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub cmdR1_Click()

If lstDoc1.ListIndex = -1 Then
    MsgBox "Selecione um documento.", vbExclamation, "Atenção"
Else
    lstDoc2.AddItem lstDoc1.Text
    lstDoc2.ItemData(lstDoc2.NewIndex) = lstDoc1.ItemData(lstDoc1.ListIndex)
    lstDoc1.RemoveItem lstDoc1.ListIndex
End If

End Sub

Private Sub cmdR2_Click()
If lstDoc2.ListIndex = -1 Then
    MsgBox "Selecione um documento a remover.", vbExclamation, "Atenção"
Else
    lstDoc1.AddItem lstDoc2.Text
    lstDoc1.ItemData(lstDoc1.NewIndex) = lstDoc2.ItemData(lstDoc2.ListIndex)
    lstDoc2.RemoveItem lstDoc2.ListIndex
End If

End Sub

Private Sub cmdRefresh_Click()
CarregaAssunto
CarregaLista
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
CarregaAssunto
CarregaLista
Eventos "INICIAR"
lstAssunto.ListIndex = 0
lstAssunto_Click
End Sub

Private Sub CarregaLista()

lstCC1.Clear
Sql = "SELECT CODIGO,DESCRICAO,ATIVO FROM CPCENTROCUSTO WHERE ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstCC1.AddItem !Descricao
        lstCC1.ItemData(lstCC1.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
lstDoc1.Clear
Sql = "SELECT CODIGO,NOME FROM CPDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstDoc1.AddItem !nome
        lstDoc1.ItemData(lstDoc1.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lstAssunto_Click()
CarregaLista
CarregaCC
CarregaDoc
txtCod.Text = Format(lstAssunto.ItemData(lstAssunto.ListIndex), "000")
txtDesc.Text = lstAssunto.Text
End Sub

Private Sub CarregaCC()

If lstAssunto.ListIndex = -1 Then Exit Sub
z = SendMessage(lvCC2.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT CPASSUNTOCC.SEQ, CPASSUNTOCC.CODCC, CPCENTROCUSTO.DESCRICAO "
Sql = Sql & "FROM CPASSUNTOCC INNER JOIN CPCENTROCUSTO ON CPASSUNTOCC.CODCC = CPCENTROCUSTO.CODIGO "
Sql = Sql & "Where CPCENTROCUSTO.Ativo = 1 And CPASSUNTOCC.CODASSUNTO =" & lstAssunto.ItemData(lstAssunto.ListIndex)
Sql = Sql & " ORDER BY CPASSUNTOCC.SEQ "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvCC2.ListItems.Add(, "C" & Format(lvCC2.ListItems.Count + 1, "00") & Format(!CODCC, "000"), !Descricao)
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaDoc()
Dim x As Integer
If lstAssunto.ListIndex = -1 Then Exit Sub
lstDoc2.Clear
Sql = "SELECT CPASSUNTODOC.CODDOC,CPDOCUMENTO.NOME FROM CPASSUNTODOC INNER JOIN "
Sql = Sql & "CPDOCUMENTO ON CPASSUNTODOC.CODDOC = CPDOCUMENTO.CODIGO "
Sql = Sql & "Where CPASSUNTODOC.CODASSUNTO = " & lstAssunto.ItemData(lstAssunto.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstDoc2.AddItem !nome
        lstDoc2.ItemData(lstDoc2.NewIndex) = Format(!CODDOC, "00")
        For x = 0 To lstDoc1.ListCount - 1
            If lstDoc1.ItemData(x) = !CODDOC Then
                lstDoc1.RemoveItem (x)
                Exit For
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lvCC2_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objDrag As ListItem
    Dim objDrop As ListItem
    Dim objNew As ListItem
    Dim objSub As ListSubItem
    Dim intIndex As Integer
    If sEvento = "" Then Exit Sub
    'Retrieve the original items
    Set objDrop = lvCC2.HitTest(x, y)
    Set objDrag = lvCC2.SelectedItem
    If (objDrop Is Nothing) Or (objDrag Is Nothing) Then
        Set lvCC2.DropHighlight = Nothing
        Set objDrop = Nothing
        Set objDrag = Nothing
        Exit Sub
    End If
    
    'Retrieve the drop position
    intIndex = objDrop.Index
    
    'Remove the dragged item
    lvCC2.ListItems.Remove objDrag.Index
    'Add it back into the dropped position
    Set objNew = lvCC2.ListItems.Add(intIndex, objDrag.Key, objDrag.Text, objDrag.Icon, objDrag.SmallIcon)
    'Copy the original subitems to the new item
    If objDrag.ListSubItems.Count > 0 Then
        For Each objSub In objDrag.ListSubItems
            objNew.ListSubItems.Add objSub.Index, objSub.Key, objSub.Text, objSub.ReportIcon, objSub.ToolTipText
        Next
    End If
    'Reselect the item
    objNew.Selected = True
    
    'Destroy all objects
    Set objNew = Nothing
    Set objDrag = Nothing
    Set objDrop = Nothing
    Set lvCC2.DropHighlight = Nothing
    
    RefreshKey
End Sub

Private Sub lvCC2_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'Highlight the item below the drag so the user knows where it will fall
Set lvCC2.DropHighlight = lvCC2.HitTest(x, y)
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmCPAssunto
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   lstAssunto.Enabled = True
   cmdC1.Enabled = False
   cmdC2.Enabled = False
   cmdCR1.Enabled = False
   cmdDR1.Enabled = False
   cmdR1.Enabled = False
   cmdR2.Enabled = False
   lstCC1.Enabled = False
   lstDoc1.Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmCPAssunto
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   txtCod.Locked = True
   txtCod.BackColor = Kde
   lstAssunto.Enabled = False
   cmdC1.Enabled = True
   cmdC2.Enabled = True
   cmdCR1.Enabled = True
   cmdDR1.Enabled = True
   cmdR1.Enabled = True
   cmdR2.Enabled = True
   lstCC1.Enabled = True
   lstDoc1.Enabled = True
End If

txtAssunto.Enabled = True
txtAssunto.BackColor = Branco

End Sub

Private Sub RefreshKey()
Dim R As Integer, sKey As String, nCount As Integer
nCount = 99

For R = lvCC2.ListItems.Count To 1 Step -1
    sKey = lvCC2.ListItems(R).Key
    sKey = "C" & Format(nCount, "00") & Right$(lvCC2.ListItems(R).Key, 3)
    lvCC2.ListItems(R).Key = sKey
    nCount = nCount - 1
Next

For R = 1 To lvCC2.ListItems.Count
    sKey = lvCC2.ListItems(R).Key
    sKey = "C" & Format(R, "00") & Right$(lvCC2.ListItems(R).Key, 3)
    lvCC2.ListItems(R).Key = sKey
Next

End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDesc.Text = ""
z = SendMessage(lvCC2.hwnd, LVM_DELETEALLITEMS, 0, 0)
lstDoc2.Clear
CarregaLista
End Sub

Private Sub Grava()
Dim nLastCod As Integer, z As Integer

If sEvento = "Novo" Then
   Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM CPASSUNTO"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nLastCod = RdoAux!MAXIMO + 1
   RdoAux.Close
   txtCod.Text = Format(nLastCod, "000")
   lstAssunto.AddItem sDesc
   lstAssunto.ItemData(lstAssunto.NewIndex) = nLastCod
   
   Sql = "INSERT CPASSUNTO (CODIGO,NOME,ATIVO) VALUES("
   Sql = Sql & nLastCod & ",'" & Mask(Left$(sDesc, 70)) & "',1)"
Else
   nLastCod = Val(txtCod.Text)
   Sql = "DELETE FROM CPASSUNTOCC WHERE CODASSUNTO=" & nLastCod
   cn.Execute Sql, rdExecDirect
   Sql = "DELETE FROM CPASSUNTODOC WHERE CODASSUNTO=" & nLastCod
   cn.Execute Sql, rdExecDirect
   Sql = "UPDATE CPASSUNTO SET NOME='" & Mask(sDesc) & "' "
   Sql = Sql & "WHERE CODIGO=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect

For z = 1 To lvCC2.ListItems.Count
    Sql = "INSERT CPASSUNTOCC (CODASSUNTO,SEQ,CODCC) VALUES("
    Sql = Sql & nLastCod & "," & Val(Mid$(lvCC2.ListItems(z).Key, 2, 2)) & ","
    Sql = Sql & Val(Right$(lvCC2.ListItems(z).Key, 3)) & ")"
    cn.Execute Sql, rdExecDirect
Next

For z = 0 To lstDoc2.ListCount - 1
    Sql = "INSERT CPASSUNTODOC (CODASSUNTO,CODDOC) VALUES("
    Sql = Sql & nLastCod & "," & lstDoc2.ItemData(z) & ")"
    cn.Execute Sql, rdExecDirect
Next

End Sub


Private Sub CarregaAssunto()
z = SendMessage(lvCC2.hwnd, LVM_DELETEALLITEMS, 0, 0)
lstDoc2.Clear
lstAssunto.Clear
Sql = "SELECT CODIGO,NOME,ATIVO FROM CPASSUNTO WHERE ATIVO=1 AND NOME LIKE '%" & Mask(txtAssunto.Text) & "%'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstAssunto.AddItem !nome
        lstAssunto.ItemData(lstAssunto.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
If lstAssunto.ListCount > 0 Then
    lstAssunto.ListIndex = 0
End If
End Sub

Private Sub txtAssunto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    cmdRefresh_Click
End If
End Sub

