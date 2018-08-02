VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmEscContab 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Escritórios Contabeis"
   ClientHeight    =   3975
   ClientLeft      =   10530
   ClientTop       =   5415
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7545
   Begin VB.TextBox txtIM 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4410
      MaxLength       =   6
      TabIndex        =   1
      Top             =   150
      Width           =   1005
   End
   Begin VB.TextBox txtRG 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   4
      Top             =   825
      Width           =   2505
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   3195
      TabIndex        =   36
      Top             =   3630
      Width           =   1005
      _ExtentX        =   1773
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
      MICON           =   "frmEscContab.frx":0000
      PICN            =   "frmEscContab.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEtiqueta 
      Height          =   315
      Left            =   5340
      TabIndex        =   35
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   3630
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Etiquetas"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEscContab.frx":0176
      PICN            =   "frmEscContab.frx":0192
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
      Left            =   6480
      TabIndex        =   31
      ToolTipText     =   "Sair da Tela"
      Top             =   3630
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEscContab.frx":0243
      PICN            =   "frmEscContab.frx":025F
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
      Left            =   4230
      TabIndex        =   26
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   3630
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmEscContab.frx":02CD
      PICN            =   "frmEscContab.frx":02E9
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
      Left            =   6390
      TabIndex        =   27
      ToolTipText     =   "Cancelar Edição"
      Top             =   3630
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
      MICON           =   "frmEscContab.frx":0443
      PICN            =   "frmEscContab.frx":045F
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
      Left            =   30
      TabIndex        =   28
      ToolTipText     =   "Novo Registro"
      Top             =   3630
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmEscContab.frx":05B9
      PICN            =   "frmEscContab.frx":05D5
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
      Left            =   1080
      TabIndex        =   29
      ToolTipText     =   "Editar Registro"
      Top             =   3630
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmEscContab.frx":072F
      PICN            =   "frmEscContab.frx":074B
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
      Left            =   2130
      TabIndex        =   30
      ToolTipText     =   "Excluir Registro"
      Top             =   3630
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmEscContab.frx":08A5
      PICN            =   "frmEscContab.frx":08C1
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
      Left            =   5310
      TabIndex        =   32
      ToolTipText     =   "Gravar os Dados"
      Top             =   3630
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEscContab.frx":0963
      PICN            =   "frmEscContab.frx":097F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstEsc 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4050
      Width           =   7515
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   3585
      Left            =   30
      TabIndex        =   14
      Top             =   0
      Width           =   7515
      Begin VB.ListBox lstNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   1785
         ItemData        =   "frmEscContab.frx":0D24
         Left            =   1110
         List            =   "frmEscContab.frx":0D26
         TabIndex        =   43
         Top             =   1500
         Visible         =   0   'False
         Width           =   5835
      End
      Begin VB.ComboBox cmbCidade 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4320
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2160
         Width           =   3150
      End
      Begin VB.ComboBox cmbBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2520
         Width           =   2460
      End
      Begin VB.TextBox txtCRC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   30
         TabIndex        =   0
         Top             =   170
         Width           =   1605
      End
      Begin VB.CheckBox chkCarne 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Aceita receber carnês"
         Height          =   195
         Left            =   5550
         TabIndex        =   2
         Top             =   220
         Width           =   1875
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         MaxLength       =   40
         TabIndex        =   13
         Top             =   3180
         Width           =   5805
      End
      Begin esMaskEdit.esMaskedEdit mskCEP 
         Height          =   285
         Left            =   4335
         TabIndex        =   11
         Top             =   2490
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   503
         MouseIcon       =   "frmEscContab.frx":0D28
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
         MaxLength       =   9
         Mask            =   "99999-999"
         SelText         =   ""
         Text            =   "_____-___"
         HideSelection   =   -1  'True
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   1110
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   210
         Width           =   405
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   3
         Top             =   500
         Width           =   6345
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1500
         Width           =   6345
      End
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1830
         Width           =   975
      End
      Begin VB.ComboBox cmbUF 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2160
         Width           =   2445
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         MaxLength       =   200
         TabIndex        =   12
         Top             =   2850
         Width           =   6345
      End
      Begin prjChameleon.chameleonButton cmdEmail 
         Height          =   345
         Left            =   6960
         TabIndex        =   34
         ToolTipText     =   "Abrir Email"
         Top             =   3150
         Width           =   495
         _ExtentX        =   873
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEscContab.frx":0D44
         PICN            =   "frmEscContab.frx":0D60
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskCPF 
         Height          =   285
         Left            =   1110
         TabIndex        =   5
         Top             =   1170
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         MouseIcon       =   "frmEscContab.frx":0DF3
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
         MaxLength       =   14
         Mask            =   "999.999.999-99"
         SelText         =   ""
         Text            =   "___.___.___-__"
         HideSelection   =   -1  'True
      End
      Begin esMaskEdit.esMaskedEdit mskCNPJ 
         Height          =   285
         Left            =   4980
         TabIndex        =   6
         Top             =   1170
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         MouseIcon       =   "frmEscContab.frx":0E0F
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
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         SelText         =   ""
         Text            =   "__.___.___/____-__"
         HideSelection   =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.:"
         Height          =   225
         Index           =   15
         Left            =   2400
         TabIndex        =   42
         Top             =   1860
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "I.M.:"
         Height          =   225
         Index           =   14
         Left            =   3960
         TabIndex        =   41
         Top             =   210
         Width           =   315
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RG..............:"
         Height          =   225
         Index           =   13
         Left            =   90
         TabIndex        =   40
         Top             =   876
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF.............:"
         Height          =   225
         Index           =   12
         Left            =   90
         TabIndex        =   39
         Top             =   1209
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CNPJ..:"
         Height          =   225
         Index           =   11
         Left            =   4290
         TabIndex        =   38
         Top             =   1200
         Width           =   660
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CRC..:"
         Height          =   225
         Index           =   10
         Left            =   1620
         TabIndex        =   37
         Top             =   210
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email...........:"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   33
         Top             =   3210
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome..........:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   543
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Número.......:"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   1875
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logradouro..:"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   1542
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro..........:"
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   20
         Top             =   2535
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade:"
         Height          =   225
         Index           =   6
         Left            =   3720
         TabIndex        =   19
         Top             =   2190
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...............:"
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   18
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP:"
         Height          =   225
         Index           =   8
         Left            =   3720
         TabIndex        =   17
         Top             =   2520
         Width           =   585
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone.....:"
         Height          =   225
         Index           =   9
         Left            =   90
         TabIndex        =   16
         Top             =   2874
         Width           =   975
      End
   End
   Begin VB.TextBox txtCompl 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      MaxLength       =   30
      TabIndex        =   9
      Top             =   1820
      Width           =   3885
   End
End
Attribute VB_Name = "frmEscContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String, bExec As Boolean, bExec2 As Boolean
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean


Private Sub cmbCidade_Click()
If Not bExec Then Exit Sub
If cmbCidade.ListIndex = -1 Then Exit Sub
cmbBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "' AND CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do While Not .EOF
            cmbBairro.AddItem !DescBairro
            cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
       .MoveNext
    Loop
   .Close
End With
If cmbBairro.ListCount > 0 Then
    bExec = False
    cmbBairro.ListIndex = 0
    bExec = True
End If

End Sub

Private Sub cmbUF_Click()
If Not bExec Then Exit Sub
cmbCidade.Clear
cmbBairro.Clear
Sql = "SELECT CODCIDADE,DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & Left$(cmbUF.Text, 2) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do While Not .EOF
       cmbCidade.AddItem !descCidade
       cmbCidade.ItemData(cmbCidade.NewIndex) = !CodCidade
      .MoveNext
    Loop
   .Close
End With
If cmbCidade.ListCount > 0 Then
    bExec = False
    cmbCidade.ListIndex = 0
    bExec = True
End If
End Sub

Private Sub cmdAlterar_Click()
    If Val(txtCod.Text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = txtNome.Text
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdConsultar_Click()

If Me.Height > 6000 Then
   cmdConsultar.Caption = "Consultar"
   Me.Height = 4420
Else
   cmdConsultar.Caption = "Selecionar"
   Me.Height = 6300
End If
cmdNovo.Enabled = Not cmdNovo.Enabled
cmdAlterar.Enabled = Not cmdAlterar.Enabled
cmdExcluir.Enabled = Not cmdExcluir.Enabled
cmdGravar.Enabled = Not cmdGravar.Enabled
cmdCancel.Enabled = Not cmdCancel.Enabled

Centraliza Me
End Sub

Private Sub cmdEmail_Click()
If Trim(txtEmail.Text) <> "" Then
    Call ShellExecute(0&, vbNullString, "mailto: " & txtEmail.Text, vbNullString, vbNullString, vbNormalFocus)
Else
    MsgBox "Sem email.", vbCritical, "ERRO"
End If
End Sub

Private Sub cmdEtiqueta_Click()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim xId As Long, nNumRec As Long, nCodLogr As Long, sCodInscricao As String, sContribuinte As String
Dim sEnd As String, nNum As Integer, sCep As String, sCompl As String, sBairro As String
Dim sEndEntrega As String, sBairroEntrega As String, sCidEntrega As String, sCepEntrega As String, sUFEntrega As String, sNumEntrega As String

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT * FROM ESCRITORIOCONTABIL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
Do Until .EOF
    sCodInscricao = !codigoesc
    sContribuinte = !NOMEESC
    sEndEntrega = SubNull(!NomeLogradouro) & ", " & SubNull(!Numero)
    sBairroEntrega = SubNull(!NOMEBairro)
    sCidEntrega = SubNull(!NomeCidade)
    sCepEntrega = SubNull(!Cep)
    sComplEntrega = ""
    sUFEntrega = SubNull(!UF)
    
    Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
    Sql = Sql & NomeDeLogin & "'," & xId & ",'" & sCodInscricao & "','" & Mask(sContribuinte) & "','"
    Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
    cn.Execute Sql, rdExecDirect
    xId = xId + 1
proximo:
   .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub

Private Sub cmdExcluir_Click()

On Error GoTo Erro
    If Val(txtCod.Text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir este Escritório de Contabilidade ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM ESCRITORIOCONTABIL WHERE CODIGOESC=" & Val(txtCod.Text)
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtNome.Text
       Limpa
       bExec = True
       CarregaLista
       bExec = False
    End If

Exit Sub
Erro:

MsgBox "Não é possivel excluir este escritório pois existem empresas cadastrados com ela.", vbExclamation, "Atenção"

End Sub

Private Sub cmdGravar_Click()
    If txtNome.Text = "" Then
       MsgBox "Favor digitar o Nome do Estabelecimento.", vbExclamation, "Atenção"
       txtNome.SetFocus
       Exit Sub
    End If
    
    If mskCPF.ClipText <> "" Then
        If Not ValidaCPF(mskCPF.ClipText) Then
           MsgBox "CPF inválido.", vbExclamation, "Atenção"
           Exit Sub
        End If
    End If
    
    If mskCNPJ.ClipText <> "" Then
        If Not ValidaCGC(mskCNPJ.ClipText) Then
           MsgBox "Cnpj inválido.", vbExclamation, "Atenção"
           Exit Sub
        End If
    End If
    
    If cmbCidade.ListIndex = -1 Then
       MsgBox "Selecione a cidade.", vbExclamation, "Atenção"
       Exit Sub
    End If
    
    If cmbBairro.ListIndex = -1 Then
       MsgBox "Selecione o bairro.", vbExclamation, "Atenção"
       Exit Sub
    End If
    
    If cmbCidade.ItemData(cmbCidade.ListIndex) = 413 And Val(txtNomeLog.Tag) = 0 Then
       MsgBox "Selecione o logradouro de Jaboticabal." & vbCrLf & " (Digite parte do nome e clique em ENTER)", vbExclamation, "Atenção"
       Exit Sub
    End If
    
    Grava
    Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
    Limpa
    cmbUF.Text = "SP-SÃO PAULO"
    cmbUF_Click
    cmbCidade.Text = "JABOTICABAL"
    cmbCidade_Click
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub cmdPrint_Click()
frmReport.ShowReport "CONTADOR", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()
Dim RdoAux2 As rdoResultset
Ocupado

Centraliza Me
bExec = True
sRet = RetEventUserForm(Me.Name)

Sql = "Select SIGLAUF,DESCUF From UF"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
       cmbUF.AddItem !SiglaUF & "-" & !DESCUF
      .MoveNext
    Loop
   .Close
End With
bExec2 = True
CarregaLista
bExec2 = False
bExec = True
Le
Eventos "INICIAR"

End Sub

Private Sub CarregaLista()

If Not bExec2 Then Exit Sub

lstEsc.Clear
Sql = "Select CODIGOESC,NOMEESC FROM ESCRITORIOCONTABIL WHERE CODIGOESC>0"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
       lstEsc.AddItem !NOMEESC
       lstEsc.ItemData(lstEsc.NewIndex) = !codigoesc
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdConsultar.Visible = True
   cmdEtiqueta.Visible = True
   cmdPrint.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmEscContab
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
         Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   chkCarne.Enabled = False
   mskCEP.BackColor = Kde
   mskCEP.Enabled = False
   mskCPF.BackColor = Kde
   mskCPF.Enabled = False
   mskCNPJ.BackColor = Kde
   mskCNPJ.Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdConsultar.Visible = False
   cmdEtiqueta.Visible = False
   cmdPrint.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmEscContab
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   chkCarne.Enabled = True
   txtNome.SetFocus
   txtCod.BackColor = Kde
   txtCod.Locked = True
   mskCEP.BackColor = vbWhite
   mskCEP.Enabled = True
   mskCPF.BackColor = vbWhite
   mskCPF.Enabled = True
   mskCNPJ.BackColor = vbWhite
   mskCNPJ.Enabled = True
End If

FormHagana

End Sub

Private Sub Le()
Dim x As Integer
If Val(txtCod.Text) = 0 Then
   Limpa
   Exit Sub
End If

Sql = "SELECT escritoriocontabil.*,logradouro.endereco "
Sql = Sql & "FROM ESCRITORIOCONTABIL  LEFT OUTER JOIN logradouro ON escritoriocontabil.codlogradouro = logradouro.codlogradouro WHERE CODIGOESC=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        .Close
        Exit Sub
    End If
    txtNome.Text = !NOMEESC
    If Val(SubNull(!CodLogradouro)) = 0 Then
        txtNomeLog.Text = SubNull(!NomeLogradouro)
        txtNomeLog.Tag = "0"
    Else
        txtNomeLog.Text = SubNull(!Endereco)
        txtNomeLog.Tag = !CodLogradouro
    End If
    txtNum.Text = Val(SubNull(!Numero))
    If Not IsNull(!Cep) Then
       mskCEP.Text = !Cep
    End If
    For x = 0 To cmbUF.ListCount - 1
        If Left$(cmbUF.List(x), 2) = !UF Then
            cmbUF.ListIndex = x
           Exit For
         End If
    Next
    For x = 0 To cmbCidade.ListCount - 1
        If cmbCidade.ItemData(x) = !CodCidade Then
           cmbCidade.ListIndex = x
           Exit For
         End If
    Next
    For x = 0 To cmbBairro.ListCount - 1
        If cmbBairro.ItemData(x) = !CodBairro Then
            cmbBairro.ListIndex = x
           Exit For
         End If
    Next
    
    txtFone.Text = SubNull(!telefone)
    txtEmail.Text = SubNull(!email)
    chkCarne.value = IIf(!RECEBECARNE, vbChecked, vbUnchecked)
    txtCRC.Text = SubNull(!CRC)
    If Not IsNull(!CPF) Then mskCPF.Text = Format(Trim(!CPF), "000\.000\.000-00")
    If Not IsNull(!Cnpj) Then mskCNPJ.Text = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
    If Not IsNull(!rg) Then txtRG.Text = !rg
    txtIM.Text = SubNull(!IM)
    txtCompl.Text = SubNull(!Complemento)
   .Close
End With

End Sub

Private Sub Limpa()

txtCod.Text = 0
txtNome.Text = ""
txtNomeLog.Text = ""
txtNum.Text = ""
cmbUF.ListIndex = -1
cmbCidade.ListIndex = -1
cmbBairro.ListIndex = -1
LimpaMascara mskCEP
LimpaMascara mskCNPJ
LimpaMascara mskCPF
txtCompl.Text = ""
txtIM.Text = ""
txtFone.Text = ""
txtEmail.Text = ""
chkCarne.value = vbUnchecked
txtCRC.Text = ""
End Sub

Private Sub Grava()
Dim MaxCod As Integer

Sql = "SELECT MAX(CODIGOESC) AS MAXIMO FROM ESCRITORIOCONTABIL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

If Evento = "Novo" Then
    Sql = "INSERT ESCRITORIOCONTABIL(CODIGOESC,NOMEESC,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,CODCIDADE,UF,TELEFONE,EMAIL,RECEBECARNE,CRC,CPF,CNPJ,RG,IM,COMPLEMENTO,CODLOGRADOURO) VALUES("
    Sql = Sql & MaxCod & ",'" & Mask(txtNome.Text) & "','" & Mask(txtNomeLog.Text) & "'," & Val(txtNum.Text) & ","
    Sql = Sql & cmbBairro.ItemData(cmbBairro.ListIndex) & ",'" & mskCEP.Text & "'," & cmbCidade.ItemData(cmbCidade.ListIndex) & ",'" & Left$(cmbUF.Text, 2) & "','" & Mask(txtFone.Text) & "','"
    Sql = Sql & txtEmail.Text & "'," & chkCarne.value & ",'" & Mask(txtCRC.Text) & "','" & IIf(mskCPF.ClipText = "", "", mskCPF.ClipText) & "','" & IIf(mskCNPJ.ClipText = "", "", mskCNPJ.ClipText) & "','"
    Sql = Sql & Mask(txtRG.Text) & "'," & Val(txtIM.Text) & ",'" & Mask(txtCompl.Text) & "'," & Val(txtNomeLog.Tag) & ")"
Else
    Sql = "UPDATE ESCRITORIOCONTABIL SET NOMEESC='" & Mask(txtNome.Text) & "',NOMELOGRADOURO='" & Mask(txtNomeLog.Text) & "',NUMERO=" & Val(txtNum.Text) & ","
    Sql = Sql & "CODBAIRRO=" & cmbBairro.ItemData(cmbBairro.ListIndex) & ",CEP='" & mskCEP.Text & "',CODCIDADE=" & cmbCidade.ItemData(cmbCidade.ListIndex) & ",UF='" & Left$(cmbUF.Text, 2) & "',TELEFONE='" & Mask(txtFone.Text) & "',"
    Sql = Sql & "EMAIL='" & txtEmail.Text & "',RECEBECARNE=" & chkCarne.value & ",CRC='" & Mask(txtCRC.Text) & "',CPF='" & IIf(mskCPF.ClipText = "", "", mskCPF.ClipText) & "',"
    Sql = Sql & "CNPJ='" & IIf(mskCNPJ.ClipText = "", "", mskCNPJ.ClipText) & "',RG='" & Mask(txtRG.Text) & "',IM=" & Val(txtIM.Text) & ",COMPLEMENTO='" & Mask(txtCompl.Text) & "',CODLOGRADOURO=" & Val(txtNomeLog.Tag) & " Where CODIGOESC = " & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   txtCod.Text = MaxCod
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtNome.Text
 ElseIf Evento = "Alterar" Then
   MaxCod = txtCod.Text
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.Text, "000") & " de " & sOldDesc & " para " & txtNome.Text
End If

bExec = True
CarregaLista
bExec = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If cmdNovo.Visible = True Then
      cmdNovo_Click
   Else
      cmdGravar_Click
   End If
End If
End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub lstCidade_Click()
If bExec Then Le
End Sub

Private Sub lstEsc_Click()
If lstEsc.ListIndex > -1 Then
    Limpa
    txtCod.Text = lstEsc.ItemData(lstEsc.ListIndex)
    Le
End If
End Sub


Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
    txtNomeLog.Text = lstNomeLog.Text
    txtNomeLog.Tag = CStr(lstNomeLog.ItemData(lstNomeLog.ListIndex))
    txtNum.SetFocus
End If

lstNomeLog.Visible = False

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
        txtNomeLog.Text = lstNomeLog.Text
        txtNomeLog.Tag = CStr(lstNomeLog.ItemData(lstNomeLog.ListIndex))
        txtNum.SetFocus
    End If
    lstNomeLog.Visible = False
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
End If

End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub txtIM_KeyPress(KeyAscii As Integer)
Tweak txtIM, KeyAscii, IntegerPositive
End Sub

Private Sub txtNomeLog_Change()
If Trim$(txtNomeLog) = "" Then
   txtNomeLog.Tag = "0"
End If

End Sub

Private Sub txtNomeLog_GotFocus()
txtNomeLog.SelStart = 0
txtNomeLog.SelLength = Len(txtNomeLog.Text)

End Sub

Private Sub txtNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLog.Text <> "" Then
      Sql = "select codlogradouro,endereco from logradouro where endereco like '%" & Trim$(txtNomeLog.Text) & "%'  order by endereco"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem !Endereco
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder 0
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
   txtNomeLog.Tag = "0"
End If

End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
Tweak txtNum, KeyAscii, IntegerPositive
End Sub
