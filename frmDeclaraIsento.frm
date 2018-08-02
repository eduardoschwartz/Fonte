VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDeclaraIsento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Declaração para fins de isenção de IPTU"
   ClientHeight    =   3900
   ClientLeft      =   4440
   ClientTop       =   4440
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3900
   ScaleWidth      =   7605
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "Imprimir Declaração"
      Top             =   3465
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
      MICON           =   "frmDeclaraIsento.frx":0000
      PICN            =   "frmDeclaraIsento.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frF 
      Height          =   1545
      Left            =   45
      TabIndex        =   14
      Top             =   -45
      Width           =   7530
      Begin prjChameleon.chameleonButton cmdCnsImovel 
         Height          =   270
         Left            =   1080
         TabIndex        =   0
         ToolTipText     =   "Consulta Cidadão"
         Top             =   180
         Width           =   465
         _ExtentX        =   820
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDeclaraIsento.frx":0176
         PICN            =   "frmDeclaraIsento.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdBuscaImovel 
         Height          =   270
         Left            =   1080
         TabIndex        =   1
         ToolTipText     =   "Consulta Cidadão"
         Top             =   810
         Width           =   465
         _ExtentX        =   820
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDeclaraIsento.frx":02EC
         PICN            =   "frmDeclaraIsento.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCodImovel 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   1575
         TabIndex        =   27
         Top             =   855
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "Bairro.........:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label lblBairro 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1035
         TabIndex        =   25
         Top             =   1170
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Nº..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   6255
         TabIndex        =   24
         Top             =   855
         Width           =   390
      End
      Begin VB.Label lblNum 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   6705
         TabIndex        =   23
         Top             =   855
         Width           =   705
      End
      Begin VB.Label lblEndereco 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   2295
         TabIndex        =   22
         Top             =   855
         Width           =   3720
      End
      Begin VB.Label Label1 
         Caption         =   "Imóvel........:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   855
         Width           =   885
      End
      Begin VB.Label lblCPF 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4005
         TabIndex        =   20
         Top             =   540
         Width           =   3450
      End
      Begin VB.Label Label1 
         Caption         =   "Nº do CPF..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3015
         TabIndex        =   19
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblRG 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1035
         TabIndex        =   18
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Nº de RG...:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "Requerente.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label lblRequerente 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1620
         TabIndex        =   15
         Top             =   225
         Width           =   5835
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2445
      Left            =   45
      TabIndex        =   29
      Top             =   1440
      Width           =   4335
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   3285
         TabIndex        =   8
         Top             =   2070
         Width           =   870
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   3285
         TabIndex        =   7
         Top             =   1755
         Width           =   870
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   3285
         TabIndex        =   6
         Top             =   1440
         Width           =   870
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   3285
         TabIndex        =   5
         Top             =   1125
         Width           =   870
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   3285
         TabIndex        =   4
         Top             =   810
         Width           =   870
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   3285
         TabIndex        =   3
         Top             =   495
         Width           =   870
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3285
         TabIndex        =   2
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label1 
         Caption         =   "rendas de quaisquer ativ.econômica......R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   90
         TabIndex        =   36
         Top             =   2115
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "rendas provenientes de prop. rural.........R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   90
         TabIndex        =   35
         Top             =   1800
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "juros de capital, inclusive poupança......R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   90
         TabIndex        =   34
         Top             =   1485
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "aluguéis recebidos.................................R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   90
         TabIndex        =   33
         Top             =   1170
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "doações recebidas.................................R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   90
         TabIndex        =   32
         Top             =   855
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "salários...................................................R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   31
         Top             =   540
         Width           =   3180
      End
      Begin VB.Label Label1 
         Caption         =   "aposentadoria, aux. doença ou pensão R$"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   30
         Top             =   225
         Width           =   3180
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   4410
      TabIndex        =   28
      Top             =   1440
      Width           =   3165
      Begin VB.CheckBox chk 
         Caption         =   "pensionistas"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   12
         Top             =   945
         Width           =   1455
      End
      Begin VB.CheckBox chk 
         Caption         =   "aposentados"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   11
         Top             =   675
         Width           =   1455
      End
      Begin VB.CheckBox chk 
         Caption         =   "viúvas e viúvos"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   405
         Width           =   2850
      End
      Begin VB.CheckBox chk 
         Caption         =   "pessoas portadoras de deficiência"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Width           =   2850
      End
   End
End
Attribute VB_Name = "frmDeclaraIsento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscaImovel_Click()
sForm = Me.Name
frmCnsImovel.show
frmCnsImovel.ZOrder 0
End Sub

Private Sub cmdCnsImovel_Click()
Set frm = frmCidadao
frm.sForm = Me.Name
frm.show
frm.ZOrder 0
End Sub

Private Sub cmdPrint_Click()
frmReport.ShowReport "DECLARAISENTOIPTU", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
    lblCodImovel.Caption = Format(Val(Left$(CodImovel, 7)), "000000")
    CodImovel = 0
    LeImovel
End If
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub LeImovel()
Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT * FROM vwFULLIMOVEL2 WHERE CODREDUZIDO=" & Val(lblCodImovel.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblEndereco.Caption = SubNull(!Logradouro)
    lblNum.Caption = SubNull(!Li_Num)
    lblBairro.Caption = SubNull(!DescBairro)
    .Close
End With

End Sub

Private Sub txtValor_KeyPress(Index As Integer, KeyAscii As Integer)
Tweak txtValor(Index), KeyAscii, DecimalPositive, 2
End Sub
