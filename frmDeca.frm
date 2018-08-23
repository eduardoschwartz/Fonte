VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmDeca 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Declaração cadastral (DECA)"
   ClientHeight    =   5385
   ClientLeft      =   2715
   ClientTop       =   2775
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   10050
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   600
      Left            =   45
      TabIndex        =   22
      Top             =   4770
      Width           =   9960
      Begin prjChameleon.chameleonButton cmdVoltar 
         Height          =   315
         Left            =   90
         TabIndex        =   23
         ToolTipText     =   "Voltar a tela anterior"
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Voltar"
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
         MICON           =   "frmDeca.frx":0000
         PICN            =   "frmDeca.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNext 
         Height          =   315
         Left            =   1350
         TabIndex        =   24
         ToolTipText     =   "Avançar para próxima tela"
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Avançar"
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
         MICON           =   "frmDeca.frx":0176
         PICN            =   "frmDeca.frx":0192
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
         Left            =   8610
         TabIndex        =   25
         ToolTipText     =   "Sair da Tela"
         Top             =   180
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "frmDeca.frx":02EC
         PICN            =   "frmDeca.frx":0308
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
         Left            =   3645
         TabIndex        =   26
         ToolTipText     =   "Imprimir DECA"
         Top             =   180
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Imprimir Frente"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDeca.frx":0376
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPrint2 
         Height          =   315
         Left            =   5535
         TabIndex        =   169
         ToolTipText     =   "Imprimir DECA"
         Top             =   180
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Imprimir Verso"
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
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDeca.frx":0392
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
   Begin VB.Frame frTela 
      BackColor       =   &H00EEEEEE&
      Height          =   4830
      Index           =   3
      Left            =   45
      TabIndex        =   92
      Top             =   -45
      Width           =   9960
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   3
         Left            =   90
         TabIndex        =   129
         Top             =   3465
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   7650
            TabIndex        =   221
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   765
            TabIndex        =   109
            Top             =   135
            Width           =   6405
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1710
            TabIndex        =   110
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   7650
            TabIndex        =   112
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   7650
            TabIndex        =   113
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1710
            TabIndex        =   111
            Top             =   675
            Width           =   5010
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   11
            Left            =   6795
            TabIndex        =   222
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   240
            Index           =   3
            Left            =   6750
            TabIndex        =   211
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   133
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   132
            Top             =   450
            Width           =   1545
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   131
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   3
            Left            =   7335
            TabIndex        =   130
            Top             =   180
            Width           =   330
         End
      End
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   2
         Left            =   90
         TabIndex        =   124
         Top             =   2475
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   7650
            TabIndex        =   219
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   765
            TabIndex        =   104
            Top             =   135
            Width           =   6405
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1710
            TabIndex        =   105
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   7650
            TabIndex        =   107
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   7650
            TabIndex        =   108
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1710
            TabIndex        =   106
            Top             =   675
            Width           =   5010
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   10
            Left            =   6795
            TabIndex        =   220
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   240
            Index           =   2
            Left            =   6750
            TabIndex        =   210
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   128
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   127
            Top             =   450
            Width           =   1500
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   126
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   2
            Left            =   7335
            TabIndex        =   125
            Top             =   180
            Width           =   330
         End
      End
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   1
         Left            =   90
         TabIndex        =   119
         Top             =   1485
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   7650
            TabIndex        =   217
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   765
            TabIndex        =   99
            Top             =   135
            Width           =   6405
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1710
            TabIndex        =   100
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   7650
            TabIndex        =   102
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   7650
            TabIndex        =   103
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1710
            TabIndex        =   101
            Top             =   675
            Width           =   5010
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   9
            Left            =   6795
            TabIndex        =   218
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   195
            Index           =   1
            Left            =   6750
            TabIndex        =   209
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   123
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   122
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   121
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   1
            Left            =   7335
            TabIndex        =   120
            Top             =   180
            Width           =   330
         End
      End
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   0
         Left            =   90
         TabIndex        =   93
         Top             =   495
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   7650
            TabIndex        =   215
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1710
            TabIndex        =   96
            Top             =   675
            Width           =   5010
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   7650
            TabIndex        =   98
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   7650
            TabIndex        =   97
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1710
            TabIndex        =   95
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   765
            TabIndex        =   94
            Top             =   135
            Width           =   6405
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   8
            Left            =   6795
            TabIndex        =   216
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   240
            Index           =   0
            Left            =   6795
            TabIndex        =   208
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   0
            Left            =   7335
            TabIndex        =   118
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   116
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   115
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   114
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Label lblTitP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DADOS RELATIVOS À PESSO DO TITULAR, DOS SÓCIOS OU DIRETORES 1"
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
         Index           =   0
         Left            =   945
         TabIndex        =   134
         Top             =   225
         Width           =   7485
      End
   End
   Begin VB.Frame frTela 
      BackColor       =   &H00EEEEEE&
      Height          =   4830
      Index           =   2
      Left            =   45
      TabIndex        =   88
      Top             =   -45
      Width           =   9960
      Begin VB.TextBox txtDescAmbulante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   135
         TabIndex        =   90
         Top             =   4410
         Width           =   9690
      End
      Begin VB.TextBox txtHist 
         Appearance      =   0  'Flat
         Height          =   3255
         Left            =   135
         MultiLine       =   -1  'True
         TabIndex        =   89
         Top             =   540
         Width           =   9690
      End
      Begin VB.Label Label11 
         Caption         =   "Trabalho como comércio ambulante de:"
         Height          =   240
         Left            =   135
         TabIndex        =   232
         Top             =   4185
         Width           =   2850
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "ESPECIFICAÇÕES DA ATIVIDADE DE COMÉRCIO AMBULANTE"
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
         Left            =   135
         TabIndex        =   231
         Top             =   3915
         Width           =   7710
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "HISTÓRICO DA OCORRÊNCIA"
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
         Left            =   180
         TabIndex        =   91
         Top             =   270
         Width           =   4110
      End
   End
   Begin VB.Frame frTela 
      BackColor       =   &H00EEEEEE&
      Height          =   4830
      Index           =   1
      Left            =   45
      TabIndex        =   47
      Top             =   -45
      Width           =   9960
      Begin VB.TextBox txtDescAmb 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7110
         MaxLength       =   30
         TabIndex        =   81
         Top             =   2070
         Width           =   2715
      End
      Begin VB.CheckBox chkAmbulante 
         Alignment       =   1  'Right Justify
         Caption         =   "       AMBULANTE"
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
         Left            =   7200
         TabIndex        =   230
         Top             =   1215
         Width           =   1905
      End
      Begin VB.ComboBox cmbAmbulante 
         BackColor       =   &H00EEEEEE&
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmDeca.frx":03AE
         Left            =   7110
         List            =   "frmDeca.frx":03BB
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   1485
         Width           =   2220
      End
      Begin VB.TextBox txtEndEntrega 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4095
         MaxLength       =   100
         TabIndex        =   76
         Top             =   4275
         Width           =   5685
      End
      Begin VB.CheckBox chkE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Outros"
         Height          =   240
         Index           =   4
         Left            =   7155
         TabIndex        =   86
         Top             =   3825
         Width           =   1950
      End
      Begin VB.CheckBox chkE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Mista"
         Height          =   240
         Index           =   3
         Left            =   7155
         TabIndex        =   85
         Top             =   3555
         Width           =   1950
      End
      Begin VB.CheckBox chkE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Prestação de Serviços"
         Height          =   240
         Index           =   2
         Left            =   7155
         TabIndex        =   84
         Top             =   3285
         Width           =   1950
      End
      Begin VB.CheckBox chkE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Industrial"
         Height          =   240
         Index           =   1
         Left            =   7155
         TabIndex        =   83
         Top             =   3015
         Width           =   1950
      End
      Begin VB.CheckBox chkE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Comercial"
         Height          =   240
         Index           =   0
         Left            =   7155
         TabIndex        =   82
         Top             =   2745
         Width           =   1950
      End
      Begin VB.CheckBox chkT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Pessoa Jurídica"
         Height          =   240
         Index           =   1
         Left            =   7155
         TabIndex        =   79
         Top             =   855
         Width           =   1950
      End
      Begin VB.CheckBox chkT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Pessoa Física"
         Height          =   240
         Index           =   0
         Left            =   7155
         TabIndex        =   78
         Top             =   585
         Width           =   1950
      End
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   0
         Left            =   5310
         TabIndex        =   67
         Top             =   585
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":03E5
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
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "5 - OUTRAS ALTERAÇÕES OU COMUNICAÇÕES"
         Height          =   240
         Index           =   9
         Left            =   270
         TabIndex        =   57
         Top             =   3870
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "4 - TRANSFERÊNCIAS DE ESTABELECIMENTO"
         Height          =   240
         Index           =   8
         Left            =   270
         TabIndex        =   56
         Top             =   3510
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "3 - CANCELAMENTO"
         Height          =   240
         Index           =   7
         Left            =   270
         TabIndex        =   55
         Top             =   3135
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "         -  sócios e diretores"
         Height          =   240
         Index           =   6
         Left            =   270
         TabIndex        =   54
         Top             =   2790
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "         -  de razão social"
         Height          =   240
         Index           =   5
         Left            =   270
         TabIndex        =   53
         Top             =   2430
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "         -  de endereço"
         Height          =   240
         Index           =   4
         Left            =   270
         TabIndex        =   52
         Top             =   2070
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "         -  de capital"
         Height          =   240
         Index           =   3
         Left            =   270
         TabIndex        =   51
         Top             =   1710
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "         -  de atividade"
         Height          =   240
         Index           =   2
         Left            =   270
         TabIndex        =   50
         Top             =   1350
         Width           =   4200
      End
      Begin VB.CheckBox chkO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "1 - ABERTURA"
         Height          =   240
         Index           =   0
         Left            =   270
         TabIndex        =   49
         Top             =   630
         Width           =   4200
      End
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   2
         Left            =   5310
         TabIndex        =   68
         Top             =   1305
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":0401
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   3
         Left            =   5310
         TabIndex        =   69
         Top             =   1665
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":041D
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   4
         Left            =   5310
         TabIndex        =   70
         Top             =   2025
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":0439
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   5
         Left            =   5310
         TabIndex        =   71
         Top             =   2385
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":0455
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   6
         Left            =   5310
         TabIndex        =   72
         Top             =   2745
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":0471
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   7
         Left            =   5310
         TabIndex        =   73
         Top             =   3105
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":048D
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   8
         Left            =   5310
         TabIndex        =   74
         Top             =   3465
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":04A9
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
      Begin esMaskEdit.esMaskedEdit mskO 
         Height          =   285
         Index           =   9
         Left            =   5310
         TabIndex        =   75
         Top             =   3825
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmDeca.frx":04C5
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
      Begin VB.Label Label12 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   7155
         TabIndex        =   233
         Top             =   1845
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "6 - ENDER.P/ENTREGA DE DOCUM.EM GERAL"
         Height          =   195
         Left            =   315
         TabIndex        =   214
         Top             =   4320
         Width           =   3660
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "2 - ALTERAÇÃO"
         Height          =   240
         Left            =   315
         TabIndex        =   194
         Top             =   990
         Width           =   1725
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE EMPRESA"
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
         Index           =   2
         Left            =   7020
         TabIndex        =   87
         Top             =   2520
         Width           =   2400
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "TIPO DE CONTRIBUINTE"
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
         Index           =   1
         Left            =   7020
         TabIndex        =   77
         Top             =   315
         Width           =   2400
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   6660
         X2              =   6660
         Y1              =   450
         Y2              =   4230
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   8
         Left            =   4815
         TabIndex        =   66
         Top             =   3870
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   7
         Left            =   4815
         TabIndex        =   65
         Top             =   3510
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   6
         Left            =   4815
         TabIndex        =   64
         Top             =   3135
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   5
         Left            =   4815
         TabIndex        =   63
         Top             =   2790
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   4
         Left            =   4815
         TabIndex        =   62
         Top             =   2430
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   3
         Left            =   4815
         TabIndex        =   61
         Top             =   2070
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   2
         Left            =   4815
         TabIndex        =   60
         Top             =   1710
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   1
         Left            =   4815
         TabIndex        =   59
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "em"
         Height          =   195
         Index           =   0
         Left            =   4815
         TabIndex        =   58
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "OCORRÊNCIA"
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
         Index           =   0
         Left            =   225
         TabIndex        =   48
         Top             =   315
         Width           =   1275
      End
   End
   Begin VB.Frame frTela 
      BackColor       =   &H00EEEEEE&
      Height          =   4830
      Index           =   0
      Left            =   45
      TabIndex        =   21
      Top             =   -45
      Width           =   9960
      Begin VB.TextBox txtEmailEmpresa 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   945
         MaxLength       =   100
         TabIndex        =   20
         Top             =   4185
         Width           =   8880
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         TabIndex        =   117
         Top             =   225
         Width           =   8340
      End
      Begin VB.TextBox txtCPF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7470
         TabIndex        =   19
         Top             =   3465
         Width           =   2355
      End
      Begin VB.TextBox txtRG 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4500
         TabIndex        =   18
         Top             =   3465
         Width           =   1995
      End
      Begin VB.TextBox txtCapital 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1575
         TabIndex        =   17
         Top             =   3465
         Width           =   1590
      End
      Begin VB.TextBox txtNumReg 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7470
         TabIndex        =   16
         Top             =   3105
         Width           =   2355
      End
      Begin VB.TextBox txtMunicipio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         TabIndex        =   14
         Top             =   3105
         Width           =   2625
      End
      Begin VB.TextBox txtOrgao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   15
         Top             =   3105
         Width           =   1815
      End
      Begin VB.TextBox txtNumemp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8595
         TabIndex        =   13
         Top             =   2745
         Width           =   1230
      End
      Begin VB.TextBox txtArea 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5085
         TabIndex        =   12
         Top             =   2745
         Width           =   1410
      End
      Begin VB.TextBox txtDataAbe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2700
         TabIndex        =   11
         Top             =   2745
         Width           =   1230
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7425
         TabIndex        =   10
         Top             =   2025
         Width           =   2445
      End
      Begin VB.TextBox txtZona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Top             =   2025
         Width           =   2085
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   855
         TabIndex        =   8
         Top             =   2025
         Width           =   2670
      End
      Begin VB.TextBox txtCEP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8685
         TabIndex        =   7
         Top             =   1665
         Width           =   1185
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5265
         TabIndex        =   6
         Top             =   1665
         Width           =   2850
      End
      Begin VB.TextBox txtSala 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3195
         TabIndex        =   5
         Top             =   1665
         Width           =   1185
      End
      Begin VB.TextBox txtAndar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   855
         TabIndex        =   4
         Top             =   1665
         Width           =   1230
      End
      Begin VB.TextBox txtRamo2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   945
         Width           =   9780
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         TabIndex        =   3
         Top             =   1305
         Width           =   8745
      End
      Begin VB.TextBox txtCodAtiv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6255
         TabIndex        =   2
         Top             =   945
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.TextBox txtRamo1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1890
         TabIndex        =   0
         Top             =   585
         Width           =   7980
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EMAIL......:"
         Height          =   285
         Index           =   20
         Left            =   90
         TabIndex        =   234
         Top             =   4230
         Width           =   915
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Index           =   1
         X1              =   90
         X2              =   9780
         Y1              =   4050
         Y2              =   4065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ:"
         Height          =   195
         Index           =   19
         Left            =   6570
         TabIndex        =   46
         Top             =   3510
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RG/INSC EST:"
         Height          =   195
         Index           =   18
         Left            =   3285
         TabIndex        =   45
         Top             =   3510
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CAPITAL SOCIAL:"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   44
         Top             =   3510
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº REG:"
         Height          =   195
         Index           =   16
         Left            =   6795
         TabIndex        =   43
         Top             =   3150
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ÓRGÃO:"
         Height          =   195
         Index           =   15
         Left            =   3960
         TabIndex        =   42
         Top             =   3150
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MUNICÍPIO:"
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   41
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº DE EMPREGADOS:"
         Height          =   195
         Index           =   13
         Left            =   6795
         TabIndex        =   40
         Top             =   2790
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ÁREA:"
         Height          =   195
         Index           =   12
         Left            =   4500
         TabIndex        =   39
         Top             =   2790
         Width           =   600
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DATA DO INÍCIO DE ATIVIDADE:"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   38
         Top             =   2790
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         Index           =   0
         X1              =   90
         X2              =   9780
         Y1              =   2520
         Y2              =   2535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TELEFONE:"
         Height          =   240
         Index           =   10
         Left            =   6480
         TabIndex        =   37
         Top             =   2070
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ZONA:"
         Height          =   240
         Index           =   9
         Left            =   3735
         TabIndex        =   36
         Top             =   2070
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CIDADE:"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   35
         Top             =   2070
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         Height          =   240
         Index           =   7
         Left            =   8235
         TabIndex        =   34
         Top             =   1710
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BAIRRO:"
         Height          =   240
         Index           =   6
         Left            =   4500
         TabIndex        =   33
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SALA/CONJ:"
         Height          =   240
         Index           =   5
         Left            =   2160
         TabIndex        =   32
         Top             =   1710
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ANDAR:"
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   31
         Top             =   1710
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ENDEREÇO:"
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   30
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CÓD ATIV.:"
         Height          =   240
         Index           =   2
         Left            =   5310
         TabIndex        =   29
         Top             =   990
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RAMO OU ATIVIDADE:"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   28
         Top             =   630
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NOME OU FIRMA:"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   270
         Width           =   1410
      End
   End
   Begin VB.Frame frTela 
      BackColor       =   &H00EEEEEE&
      Height          =   4830
      Index           =   5
      Left            =   45
      TabIndex        =   170
      Top             =   -45
      Width           =   9960
      Begin VB.TextBox txtCidadeC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   183
         Top             =   1665
         Width           =   5640
      End
      Begin VB.TextBox txtUFC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7785
         TabIndex        =   187
         Top             =   1665
         Width           =   1995
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6255
         TabIndex        =   190
         Top             =   2025
         Width           =   3525
      End
      Begin VB.TextBox txtAssinatura 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3015
         TabIndex        =   192
         Top             =   4365
         Width           =   5640
      End
      Begin VB.TextBox txtOBSC 
         Appearance      =   0  'Flat
         Height          =   1590
         Left            =   225
         MultiLine       =   -1  'True
         TabIndex        =   191
         Top             =   2655
         Width           =   9555
      End
      Begin VB.TextBox txtOrgaoC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3915
         TabIndex        =   189
         Top             =   2025
         Width           =   1635
      End
      Begin VB.TextBox txtRGC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         TabIndex        =   188
         Top             =   2025
         Width           =   1500
      End
      Begin VB.TextBox txtCEPC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7785
         TabIndex        =   186
         Top             =   1305
         Width           =   1995
      End
      Begin VB.TextBox txtnumC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7785
         TabIndex        =   185
         Top             =   945
         Width           =   1995
      End
      Begin VB.TextBox txtFoneC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7785
         TabIndex        =   184
         Top             =   585
         Width           =   1995
      End
      Begin VB.TextBox txtBairroC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   182
         Top             =   1305
         Width           =   5640
      End
      Begin VB.TextBox txtEndC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   181
         Top             =   945
         Width           =   5640
      End
      Begin VB.TextBox txtNomeC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1170
         TabIndex        =   180
         Top             =   585
         Width           =   5640
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   236
         Top             =   1710
         Width           =   645
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "UF:"
         Height          =   195
         Index           =   10
         Left            =   7020
         TabIndex        =   235
         Top             =   1710
         Width           =   645
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Email..:"
         Height          =   195
         Index           =   9
         Left            =   5670
         TabIndex        =   213
         Top             =   2070
         Width           =   600
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Contribuinte ou Procurador:"
         Height          =   195
         Index           =   8
         Left            =   270
         TabIndex        =   212
         Top             =   4410
         Width           =   2715
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVAÇÕES"
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
         Left            =   1800
         TabIndex        =   193
         Top             =   2430
         Width           =   5775
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "RG...:"
         Height          =   195
         Index           =   7
         Left            =   3330
         TabIndex        =   179
         Top             =   2070
         Width           =   600
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Número do CRC:"
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   178
         Top             =   2070
         Width           =   1320
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         Height          =   195
         Index           =   5
         Left            =   7020
         TabIndex        =   177
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Número:"
         Height          =   195
         Index           =   4
         Left            =   7020
         TabIndex        =   176
         Top             =   990
         Width           =   645
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone:"
         Height          =   195
         Index           =   3
         Left            =   7020
         TabIndex        =   175
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro:"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   174
         Top             =   1350
         Width           =   645
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço:"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   173
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   172
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DADOS REFERENTE À PESSOA DO CONTADOR"
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
         Left            =   1935
         TabIndex        =   171
         Top             =   270
         Width           =   5775
      End
   End
   Begin VB.Frame frTela 
      BackColor       =   &H00EEEEEE&
      Height          =   4830
      Index           =   4
      Left            =   45
      TabIndex        =   135
      Top             =   -45
      Width           =   9960
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   5
         Left            =   90
         TabIndex        =   195
         Top             =   2475
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   7650
            TabIndex        =   226
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   765
            TabIndex        =   200
            Top             =   135
            Width           =   6405
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   1710
            TabIndex        =   199
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   7650
            TabIndex        =   198
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   7650
            TabIndex        =   197
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   1710
            TabIndex        =   196
            Top             =   675
            Width           =   5010
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   14
            Left            =   6795
            TabIndex        =   227
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   195
            Index           =   5
            Left            =   6750
            TabIndex        =   206
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   204
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   203
            Top             =   450
            Width           =   1590
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   202
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   5
            Left            =   7335
            TabIndex        =   201
            Top             =   180
            Width           =   330
         End
      End
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   7
         Left            =   90
         TabIndex        =   146
         Top             =   495
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   7650
            TabIndex        =   152
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   765
            TabIndex        =   147
            Top             =   135
            Width           =   6405
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   1710
            TabIndex        =   149
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   7650
            TabIndex        =   148
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   7650
            TabIndex        =   150
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   1710
            TabIndex        =   151
            Top             =   675
            Width           =   5010
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   12
            Left            =   6795
            TabIndex        =   223
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   157
            Top             =   180
            Width           =   1140
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   156
            Top             =   450
            Width           =   1500
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   155
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   7
            Left            =   7335
            TabIndex        =   154
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   195
            Index           =   7
            Left            =   6750
            TabIndex        =   153
            Top             =   450
            Width           =   870
         End
      End
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   6
         Left            =   90
         TabIndex        =   141
         Top             =   1485
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   7650
            TabIndex        =   224
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   1710
            TabIndex        =   161
            Top             =   675
            Width           =   5010
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   7650
            TabIndex        =   163
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   7650
            TabIndex        =   162
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   1710
            TabIndex        =   160
            Top             =   405
            Width           =   5010
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   765
            TabIndex        =   158
            Top             =   135
            Width           =   6405
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   13
            Left            =   6795
            TabIndex        =   225
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   195
            Index           =   4
            Left            =   6750
            TabIndex        =   205
            Top             =   450
            Width           =   825
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   6
            Left            =   7335
            TabIndex        =   145
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   144
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   143
            Top             =   450
            Width           =   1455
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   142
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Frame frProp 
         BackColor       =   &H00EEEEEE&
         Height          =   1005
         Index           =   4
         Left            =   90
         TabIndex        =   136
         Top             =   3465
         Width           =   9780
         Begin VB.TextBox txtTelefone 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   7650
            TabIndex        =   228
            Top             =   675
            Width           =   2040
         End
         Begin VB.TextBox txtBairroP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1710
            TabIndex        =   166
            Top             =   675
            Width           =   4965
         End
         Begin VB.TextBox txtCPFP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   7650
            TabIndex        =   168
            Top             =   405
            Width           =   2040
         End
         Begin VB.TextBox txtRGP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   7650
            TabIndex        =   167
            Top             =   135
            Width           =   2040
         End
         Begin VB.TextBox txtRuaP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1710
            TabIndex        =   165
            Top             =   405
            Width           =   4965
         End
         Begin VB.TextBox txtNomeP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   765
            TabIndex        =   164
            Top             =   135
            Width           =   6405
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "TELEFONE"
            Height          =   240
            Index           =   15
            Left            =   6795
            TabIndex        =   229
            Top             =   720
            Width           =   825
         End
         Begin VB.Label lblCPF 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF/ CNPJ"
            Height          =   240
            Index           =   6
            Left            =   6705
            TabIndex        =   207
            Top             =   450
            Width           =   870
         End
         Begin VB.Label lblRG 
            BackStyle       =   0  'Transparent
            Caption         =   "RG"
            Height          =   195
            Index           =   4
            Left            =   7335
            TabIndex        =   140
            Top             =   180
            Width           =   330
         End
         Begin VB.Label lblBairro 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro, Cidade e UF"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   139
            Top             =   720
            Width           =   1500
         End
         Begin VB.Label lblRua 
            BackStyle       =   0  'Transparent
            Caption         =   "Rua, Número, CEP"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   138
            Top             =   450
            Width           =   1500
         End
         Begin VB.Label lblNome 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome"
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   137
            Top             =   180
            Width           =   1140
         End
      End
      Begin VB.Label lblTitP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DADOS RELATIVOS À PESSO DO TITULAR, DOS SÓCIOS OU DIRETORES 2"
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
         Index           =   1
         Left            =   945
         TabIndex        =   159
         Top             =   225
         Width           =   7485
      End
   End
End
Attribute VB_Name = "frmDeca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodReduz As Long, nTela As Integer

Public Property Let nCodigoEmpresa(nCodigo As Long)
    nCodReduz = nCodigo
End Property

Private Sub chkAmbulante_Click()
If chkAmbulante.value = vbChecked Then
    cmbAmbulante.Enabled = True
    cmbAmbulante.BackColor = Branco
    cmbAmbulante_Click
Else
    cmbAmbulante.Enabled = False
    cmbAmbulante.BackColor = Kde
    txtDescAmb.Text = ""
    txtDescAmbulante.Text = ""
End If

End Sub

Private Sub cmbAmbulante_Click()
If cmbAmbulante.ListIndex = -1 Then Exit Sub

If cmbAmbulante.ListIndex = 0 Then
    txtDescAmb.BackColor = Kde
    txtDescAmb.Enabled = False
Else
    txtDescAmb.BackColor = Branco
    txtDescAmb.Enabled = True
End If

End Sub

Private Sub cmdNext_Click()
nTela = nTela + 1
MudaTela
End Sub

Private Sub cmdPrint_Click()
If txtRG.Text = "" Then txtRG.Text = " "

Dim Sql As String

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "INSERT REPORTTMP(USUARIO,MEMO1) VALUES('" & NomeDeLogin & "','" & Mask(txtHist.Text) & "')"
cn.Execute Sql, rdExecDirect
frmReport.ShowReport2 "DECA", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdPrint2_Click()
frmReport.ShowReport2 "DECA2", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdVoltar_Click()
nTela = nTela - 1
MudaTela

End Sub

Private Sub Form_Load()
Centraliza Me
nTela = 0
MudaTela
cmbAmbulante.ListIndex = 0
Me.Caption = Me.Caption & " -> Inscrição Municipal nº " & Format(nCodReduz, "000000")
If nCodReduz > 0 Then Le
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset

If nCodReduz >= 100000 And nCodReduz < 300000 Then
    Sql = "SELECT * FROM vwFULLEMPRESA3 WHERE CODIGOMOB=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    With RdoAux
        txtNome.Text = !RazaoSocial
        txtRamo1.Text = !ativextenso
        txtCodAtiv.Text = !codatividade
        txtEnd.Text = !Logradouro & ", " & !Numero
        txtBairro.Text = SubNull(!DescBairro)
        If IsNull(!Cep) Then
            txtCEP.Text = Format(RetornaCEP(!CodLogradouro, !Numero), "00000-000")
        Else
            txtCEP.Text = Format(!Cep, "00000-000")
        End If
        txtCidade.Text = SubNull(!descCidade) & "/" & SubNull(!SiglaUF)
        txtFone.Text = SubNull(!fonecontato)
        txtDataAbe.Text = Format(!DATAABERTURA, "dd/mm/yyyy")
        txtArea.Text = FormatNumber(!areatl, 2)
        txtNumemp.Text = SubNull(!QTDEEMPREGADO)
        txtMunicipio.Text = SubNull(!descCidade)
        txtOrgao.Text = SubNull(!NOMEORGAO)
        txtNumReg.Text = SubNull(!NUMREGISTRORESP)
        txtCapital.Text = FormatNumber(!CAPITALSOCIAL, 2)
        txtRG.Text = SubNull(!rg)
        If txtRG.Text = "" Then
            txtRG.Text = SubNull(!INSCESTADUAL)
        End If
        txtCPF.Text = SubNull(!Cnpj)
        If txtCPF.Text <> "" Then
            'txtCPF.Text = Format(txtCPF.Text, "0#\.###\.###/####-##")
            chkT(0).value = vbUnchecked
            chkT(1).value = vbChecked
        Else
            txtCPF.Text = SubNull(!CPF)
            If txtCPF.Text <> "" Then
                'txtCPF.Text = Format(txtCPF.Text, "00#\.###\.###-##")
                chkT(0).value = vbChecked
                chkT(1).value = vbUnchecked
            End If
        End If
        chkE(0).value = vbUnchecked
        chkE(1).value = vbUnchecked
        chkE(2).value = vbUnchecked
        chkE(3).value = vbUnchecked
        chkE(4).value = vbUnchecked
        
        If Val(txtCodAtiv.Text) > 10000 And Val(txtCodAtiv.Text) < 20000 Then
            chkE(1).value = vbChecked
        ElseIf Val(txtCodAtiv.Text) > 20000 And Val(txtCodAtiv.Text) < 30000 Then
            chkE(0).value = vbChecked
        ElseIf Val(txtCodAtiv.Text) > 30000 And Val(txtCodAtiv.Text) < 40000 Then
            chkE(2).value = vbChecked
        Else
            chkE(4).value = vbChecked
        End If
        
        If Val(SubNull(!RESPCONTABIL)) > 0 Then
            txtNomeC.Text = SubNull(!NOMEESC)
            txtFoneC.Text = SubNull(!telefone)
            txtEndC.Text = SubNull(!RUAESC)
            txtnumC.Text = SubNull(!NUMEROESC)
            txtBairroC.Text = ""
            txtCEPC.Text = SubNull(!CEPESC)
        End If
        
        .Close
    End With
    Sql = "SELECT mobiliarioproprietario.codcidadao, vwFULLCIDADAO.* "
    Sql = Sql & "FROM  mobiliarioproprietario INNER JOIN vwFULLCIDADAO ON mobiliarioproprietario.codcidadao = vwFULLCIDADAO.codcidadao Where mobiliarioproprietario.codmobiliario = " & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            txtNomeP(.AbsolutePosition - 1).Text = !nomecidadao
            'txtCPFP(.AbsolutePosition - 1).Text = Format(SubNull(!CPF), "0#\.###\.###/####-##")
            txtCPFP(.AbsolutePosition - 1).Text = SubNull(!CPF)
            txtRGP(.AbsolutePosition - 1).Text = SubNull(!rg)
            txtRuaP(.AbsolutePosition - 1).Text = SubNull(!Endereco) & ", " & SubNull(!NUMIMOVEL) & ", " & RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!NUMIMOVEL)))
            txtBairroP(.AbsolutePosition - 1).Text = SubNull(!DescBairro) & ", " & SubNull(!descCidade) & " - " & SubNull(!SiglaUF)
            txtTelefone(.AbsolutePosition - 1).Text = SubNull(!telefone)
           .MoveNext
        Loop
       .Close
    End With
Else
    Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    With RdoAux
        txtNome.Text = !nomecidadao
        txtNomeP(0).Text = !nomecidadao
        txtEnd.Text = SubNull(!Endereco) & ", " & SubNull(!NUMIMOVEL)
        txtRuaP(0).Text = SubNull(!Endereco) & ", " & SubNull(!NUMIMOVEL)
        
        txtBairro.Text = SubNull(!DescBairro)
        If IsNull(!Cep) Then
            txtCEP.Text = Format(RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!NUMIMOVEL))), "00000-000")
        Else
            txtCEP.Text = Format(!Cep, "00000-000")
        End If
        txtCidade.Text = SubNull(!descCidade) & "/" & SubNull(!SiglaUF)
        txtBairroP(0).Text = SubNull(!DescBairro) & ", " & SubNull(!descCidade) & " - " & SubNull(!SiglaUF)
        txtFone.Text = SubNull(!telefone)
        txtRG.Text = SubNull(!rg)
'        If txtRG.text = "" Then
'            txtRG.text = SubNull(!INSCESTADUAL)
            
 '       End If
        txtRGP(0).Text = txtRG.Text
        txtCPF.Text = SubNull(!Cnpj)
        If txtCPF.Text <> "" Then
'            txtCPF.Text = Format(txtCPF.Text, "0#\.###\.###/####-##")
            'txtCPFP(0).Text = Format(txtCPF.Text, "0#\.###\.###/####-##")
            txtCPFP(0).Text = txtCPF.Text
            chkT(0).value = vbUnchecked
            chkT(1).value = vbChecked
        Else
            txtCPF.Text = SubNull(!CPF)
            If txtCPF.Text <> "" Then
                'txtCPF.Text = Format(txtCPF.Text, "00#\.###\.###-##")
                txtCPF.Text = txtCPF.Text
                chkT(0).value = vbChecked
                chkT(1).value = vbUnchecked
            End If
        End If
     End With
End If

End Sub

Private Sub MudaTela()
Dim x As Integer
For x = 0 To 5
    If x = nTela Then
        frTela(x).Visible = True
        frTela(x).ZOrder 0
    Else
        frTela(x).Visible = False
    End If
Next
If nTela = 0 Then
    cmdVoltar.Enabled = False
Else
    cmdVoltar.Enabled = True
End If
If nTela = 5 Then
    cmdNext.Enabled = False
Else
    cmdNext.Enabled = True
End If


End Sub

