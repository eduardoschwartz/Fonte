VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGuiaPratico1 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recolhimento aos cofres municipais"
   ClientHeight    =   2895
   ClientLeft      =   4920
   ClientTop       =   4635
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   8595
   Begin VB.ComboBox cmbPag 
      Height          =   315
      ItemData        =   "frmGuiaPratico1.frx":0000
      Left            =   1980
      List            =   "frmGuiaPratico1.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2430
      Width           =   1365
   End
   Begin VB.TextBox txtParc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3870
      TabIndex        =   21
      Top             =   2430
      Width           =   510
   End
   Begin VB.TextBox txtCod2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5580
      TabIndex        =   18
      Top             =   1980
      Width           =   1005
   End
   Begin VB.TextBox txtProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1260
      TabIndex        =   16
      Top             =   1980
      Width           =   1455
   End
   Begin VB.TextBox txtNot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6885
      TabIndex        =   14
      Top             =   1485
      Width           =   1455
   End
   Begin VB.TextBox txtArea 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   12
      Top             =   1485
      Width           =   1005
   End
   Begin VB.ComboBox cmbCateg 
      Height          =   315
      ItemData        =   "frmGuiaPratico1.frx":0022
      Left            =   1080
      List            =   "frmGuiaPratico1.frx":002F
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1485
      Width           =   1185
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmGuiaPratico1.frx":0047
      Left            =   6705
      List            =   "frmGuiaPratico1.frx":0054
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1035
      Width           =   1680
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3735
      TabIndex        =   6
      Top             =   1035
      Width           =   1005
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3735
      TabIndex        =   3
      Top             =   585
      Width           =   1905
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   225
      TabIndex        =   0
      Top             =   225
      Width           =   6180
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   7335
      TabIndex        =   20
      Top             =   2385
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   609
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
      MICON           =   "frmGuiaPratico1.frx":007C
      PICN            =   "frmGuiaPratico1.frx":0098
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
      Caption         =   "para pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   315
      TabIndex        =   25
      Top             =   2475
      Width           =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "em"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   3465
      TabIndex        =   24
      Top             =   2475
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "vezes."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   4500
      TabIndex        =   23
      Top             =   2475
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ","
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   6615
      TabIndex        =   19
      Top             =   2025
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "tendo sido lançado no código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   2835
      TabIndex        =   17
      Top             =   2025
      Width           =   2760
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", Processo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   225
      TabIndex        =   15
      Top             =   2025
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "m², conforme Notificação nº"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   4365
      TabIndex        =   13
      Top             =   1530
      Width           =   2490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", área de "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   2340
      TabIndex        =   11
      Top             =   1530
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", padrão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   225
      TabIndex        =   9
      Top             =   1530
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "sendo a construção"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4815
      TabIndex        =   7
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Construção Civil do imóvel de código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   5
      Top             =   1080
      Width           =   3390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", referente ao ISS incidente na"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5715
      TabIndex        =   4
      Top             =   630
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "cofres municipais, a importância de R$"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   630
      Width           =   3525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", deverá recolher aos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   6480
      TabIndex        =   1
      Top             =   270
      Width           =   2130
   End
End
Attribute VB_Name = "frmGuiaPratico1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
frmReport.ShowReport "GUIAPRATICO1", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
Tweak txtArea, KeyAscii, DecimalPositive
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod2_KeyPress(KeyAscii As Integer)
Tweak txtCod2, KeyAscii, IntegerPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub
