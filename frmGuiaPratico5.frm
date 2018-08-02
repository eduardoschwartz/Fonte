VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGuiaPratico5 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lançamento de IPTU proporcional"
   ClientHeight    =   1920
   ClientLeft      =   870
   ClientTop       =   5445
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1920
   ScaleWidth      =   8445
   Begin VB.TextBox txtPerc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6705
      TabIndex        =   5
      Top             =   990
      Width           =   510
   End
   Begin VB.TextBox txtParc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4905
      TabIndex        =   7
      Top             =   1395
      Width           =   510
   End
   Begin VB.ComboBox cmbPag 
      Height          =   315
      ItemData        =   "frmGuiaPratico5.frx":0000
      Left            =   3015
      List            =   "frmGuiaPratico5.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1395
      Width           =   1365
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5805
      TabIndex        =   3
      Top             =   585
      Width           =   2445
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2430
      TabIndex        =   4
      Top             =   990
      Width           =   1185
   End
   Begin VB.TextBox txtNot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox txtProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2970
      TabIndex        =   2
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4860
      TabIndex        =   0
      Top             =   225
      Width           =   915
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   7200
      TabIndex        =   8
      Top             =   1440
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
      MICON           =   "frmGuiaPratico5.frx":0022
      PICN            =   "frmGuiaPratico5.frx":003E
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
      Caption         =   "%, tendo"
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
      Left            =   7290
      TabIndex        =   19
      Top             =   1035
      Width           =   1050
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
      Index           =   6
      Left            =   5535
      TabIndex        =   18
      Top             =   1440
      Width           =   735
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
      Index           =   1
      Left            =   4500
      TabIndex        =   17
      Top             =   1440
      Width           =   420
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Foi lançado IPTU proporcional para o exercício de"
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
      Left            =   135
      TabIndex        =   16
      Top             =   270
      Width           =   4740
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", conforme Notificação nº"
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
      Left            =   5850
      TabIndex        =   15
      Top             =   270
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", em nome de"
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
      Left            =   4500
      TabIndex        =   14
      Top             =   630
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "para o imóvel de código"
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
      Left            =   135
      TabIndex        =   13
      Top             =   1035
      Width           =   2310
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", em decorrência de isenção de :"
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
      Left            =   3735
      TabIndex        =   12
      Top             =   1035
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   ", Processo nº"
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
      Left            =   1665
      TabIndex        =   11
      Top             =   630
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "sido lançado para pagamento"
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
      Left            =   135
      TabIndex        =   10
      Top             =   1440
      Width           =   2850
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "."
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
      Left            =   6480
      TabIndex        =   9
      Top             =   1530
      Width           =   330
   End
End
Attribute VB_Name = "frmGuiaPratico5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
If cmbPag.ListIndex = 1 And Val(txtParc.Text) = 0 Then
    MsgBox "Digito o numero de parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

frmReport.ShowReport2 "GUIAPRATICO5", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub Form_Load()
Centraliza Me
cmbPag.ListIndex = 0
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtParc_KeyPress(KeyAscii As Integer)
Tweak txtParc, KeyAscii, IntegerPositive
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
Tweak txtPerc, KeyAscii, IntegerPositive
End Sub
