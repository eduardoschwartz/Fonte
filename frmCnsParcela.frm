VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCnsParcela 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalhes da Parcela Nº "
   ClientHeight    =   5115
   ClientLeft      =   9600
   ClientTop       =   5340
   ClientWidth     =   10290
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   150
      Left            =   1755
      ScaleHeight     =   90
      ScaleWidth      =   405
      TabIndex        =   67
      Top             =   4950
      Visible         =   0   'False
      Width           =   465
   End
   Begin prjChameleon.chameleonButton cmdDup 
      Height          =   285
      Left            =   9810
      TabIndex        =   60
      ToolTipText     =   "Consulta Débitos Duplicados ou Restituidos"
      Top             =   4350
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCnsParcela.frx":0000
      PICN            =   "frmCnsParcela.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frDup 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consulta Débitos Duplicados e/ou Restituidos"
      Height          =   2805
      Left            =   2310
      TabIndex        =   55
      Top             =   540
      Width           =   6165
      Begin prjChameleon.chameleonButton cmdSairDup 
         Height          =   345
         Left            =   5655
         TabIndex        =   56
         ToolTipText     =   "Sair"
         Top             =   2310
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCnsParcela.frx":0176
         PICN            =   "frmCnsParcela.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid grdDup 
         Height          =   2445
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   4313
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedCols       =   0
         BackColor       =   12632256
         BackColorFixed  =   15658734
         BackColorSel    =   192
         FocusRect       =   0
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "^Pagamento   |^Recebimento   |>Valor Pago   |^Restituido       |^Doc               "
      End
   End
   Begin VB.Frame Panel1 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   705
      Index           =   5
      Left            =   30
      TabIndex        =   46
      Top             =   4350
      Width           =   10245
      Begin prjChameleon.chameleonButton cmdPrint 
         Height          =   315
         Left            =   8070
         TabIndex        =   47
         ToolTipText     =   "Imprimir Detalhe"
         Top             =   375
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
         MICON           =   "frmCnsParcela.frx":02EC
         PICN            =   "frmCnsParcela.frx":0308
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
         Left            =   9150
         TabIndex        =   48
         ToolTipText     =   "Sair da Tela"
         Top             =   375
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsParcela.frx":0462
         PICN            =   "frmCnsParcela.frx":047E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblContrib 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1230
         TabIndex        =   66
         Top             =   270
         Width           =   6735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Contribuin:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   45
         TabIndex        =   65
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "% Desconto:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   6090
         TabIndex        =   62
         Top             =   465
         Width           =   1155
      End
      Begin VB.Label lblDesconto 
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   7350
         TabIndex        =   61
         Top             =   465
         Width           =   435
      End
      Begin VB.Label lblIsentoMJ 
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   7380
         TabIndex        =   59
         Top             =   30
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Isento M/J:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   6120
         TabIndex        =   58
         Top             =   30
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Status....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   23
         Left            =   45
         TabIndex        =   54
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "DÉBITO PAGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1230
         TabIndex        =   53
         Top             =   495
         Width           =   3945
      End
      Begin VB.Label lblLanc 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1260
         TabIndex        =   52
         Top             =   15
         Width           =   6705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamento:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   51
         Top             =   15
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dupl/Restit:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   7980
         TabIndex        =   50
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblDup 
         BackStyle       =   0  'Transparent
         Caption         =   "Não"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   9360
         TabIndex        =   49
         Top             =   30
         Width           =   435
      End
   End
   Begin VB.Frame Panel1 
      BackColor       =   &H00EEEEEE&
      Height          =   1065
      Index           =   3
      Left            =   30
      TabIndex        =   34
      Top             =   3255
      Width           =   6855
      Begin VB.TextBox txtLivro 
         Alignment       =   2  'Center
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0000000"
         Top             =   330
         Width           =   765
      End
      Begin VB.TextBox txtPagina 
         Alignment       =   2  'Center
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   870
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "0000000"
         Top             =   600
         Width           =   765
      End
      Begin VB.TextBox txtCertidao 
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2940
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "00000000"
         Top             =   600
         Width           =   945
      End
      Begin esMaskEdit.esMaskedEdit mskIncricao 
         Height          =   255
         Left            =   2940
         TabIndex        =   35
         Top             =   330
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   450
         BackColor       =   15658734
         MouseIcon       =   "frmCnsParcela.frx":04EC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Locked          =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Livro..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   45
         TabIndex        =   45
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Página.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   30
         TabIndex        =   44
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Certidão..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   1770
         TabIndex        =   43
         Top             =   615
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   1770
         TabIndex        =   42
         Top             =   315
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuizamento.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   21
         Left            =   4230
         TabIndex        =   41
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label lblAjuizamento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   5670
         TabIndex        =   40
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000080&
         Caption         =   "  Dados da Divida Ativa e Ajuizamento"
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
         Left            =   -30
         TabIndex        =   39
         Top             =   0
         Width           =   6825
      End
   End
   Begin VB.Frame Panel1 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   2625
      Index           =   2
      Left            =   6930
      TabIndex        =   18
      Top             =   1665
      Width           =   3315
      Begin VB.TextBox txtValorPago 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   71
         Text            =   "0,00"
         Top             =   2070
         Width           =   1305
      End
      Begin VB.TextBox txtValorDiferenca 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1830
         Locked          =   -1  'True
         TabIndex        =   70
         Text            =   "0,00"
         Top             =   2310
         Width           =   1305
      End
      Begin VB.Label lblDataPagtoCalc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1740
         TabIndex        =   69
         Top             =   553
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Pagto.Calc.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   24
         Left            =   0
         TabIndex        =   68
         Top             =   553
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Pagamento..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   0
         TabIndex        =   33
         Top             =   300
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Receita....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   0
         TabIndex        =   32
         Top             =   806
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco Debitado..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   0
         TabIndex        =   31
         Top             =   1059
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   0
         TabIndex        =   30
         Top             =   2071
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Diferença.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   0
         TabIndex        =   29
         Top             =   2325
         Width           =   1785
      End
      Begin VB.Label lblDataPagto 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1740
         TabIndex        =   28
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label lblDataReceita 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1740
         TabIndex        =   27
         Top             =   806
         Width           =   1395
      End
      Begin VB.Label lblBanco 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1770
         TabIndex        =   26
         Top             =   1059
         Width           =   1365
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000080&
         Caption         =   "  Dados do Pagamento"
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
         Left            =   -60
         TabIndex        =   25
         Top             =   0
         Width           =   3405
      End
      Begin VB.Label lblAgencia 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1740
         TabIndex        =   24
         Top             =   1312
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agência Debitada:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   0
         TabIndex        =   23
         Top             =   1312
         Width           =   1785
      End
      Begin VB.Label lblValorTaxa 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1740
         TabIndex        =   22
         Top             =   1818
         Width           =   1395
      End
      Begin VB.Label lblNumDoc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0000000000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1740
         TabIndex        =   21
         Top             =   1565
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Taxa Exp.Doc....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   0
         TabIndex        =   20
         Top             =   1818
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Num. Documento..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   0
         TabIndex        =   19
         Top             =   1565
         Width           =   1785
      End
   End
   Begin VB.Frame Panel1 
      BackColor       =   &H00EEEEEE&
      Height          =   1635
      Index           =   1
      Left            =   6930
      TabIndex        =   8
      Top             =   30
      Width           =   3315
      Begin VB.Label lblDataVenctoCalc 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1815
         TabIndex        =   64
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vecto.Calc.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   22
         Left            =   45
         TabIndex        =   63
         Top             =   855
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Lançado...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   17
         Top             =   1110
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vencimento.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   16
         Top             =   570
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base.......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   15
         Top             =   300
         Width           =   1785
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Atual/pago:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   30
         TabIndex        =   14
         Top             =   1380
         Width           =   1785
      End
      Begin VB.Label lblValorLancado 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1800
         TabIndex        =   13
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label lblDataVencto 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1800
         TabIndex        =   12
         Top             =   570
         Width           =   1395
      End
      Begin VB.Label lblDataBase 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00/00/0000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1800
         TabIndex        =   11
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label lblValorAtualizado 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1800
         TabIndex        =   10
         Top             =   1380
         Width           =   1395
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000080&
         Caption         =   "  Valores Atualizados"
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
         Left            =   -30
         TabIndex        =   9
         Top             =   0
         Width           =   3345
      End
   End
   Begin VB.Frame Panel1 
      BackColor       =   &H00EEEEEE&
      Height          =   3180
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6855
      Begin MSFlexGridLib.MSFlexGrid grdTrib 
         Height          =   2565
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   4524
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColor       =   15658734
         ForeColor       =   12582912
         BackColorFixed  =   8388608
         ForeColorFixed  =   16774648
         BackColorSel    =   14737632
         ForeColorSel    =   0
         BackColorBkg    =   15658734
         GridColor       =   15658734
         GridColorFixed  =   15658734
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   "<Descrição                                           |>Valor Lanc.      |>Juros        |>Multa         |>Correção     "
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1950
         TabIndex        =   7
         Top             =   2880
         Width           =   615
      End
      Begin VB.Line Line2 
         X1              =   2760
         X2              =   6690
         Y1              =   2865
         Y2              =   2865
      End
      Begin VB.Label lblTotL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Height          =   225
         Left            =   2910
         TabIndex        =   6
         Top             =   2910
         Width           =   1065
      End
      Begin VB.Label lblTotJ 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Height          =   225
         Left            =   4140
         TabIndex        =   5
         Top             =   2910
         Width           =   675
      End
      Begin VB.Label lblTotM 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Height          =   225
         Left            =   5010
         TabIndex        =   4
         Top             =   2910
         Width           =   675
      End
      Begin VB.Label lblTotC 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
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
         Height          =   225
         Left            =   5940
         TabIndex        =   3
         Top             =   2910
         Width           =   675
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000080&
         Caption         =   "  Composição da Parcela"
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
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6825
      End
   End
End
Attribute VB_Name = "frmCnsParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nNumParc As Integer
Dim nAnoExer As Integer
Dim nCodLanc As Integer
Dim nCodSeq As Integer
Dim nCodComp As Integer
Dim nCodReduzido As Long
Dim nResp As Integer

Public Property Let nParcela(nNumeroParcela As Integer)
    nNumParc = nNumeroParcela
End Property

Public Property Let nResponsavel(nCodigoResponsavel As Integer)
    nResp = nCodigoResponsavel
End Property

Public Property Let nLancamento(nCodLancamento As Integer)
    nCodLanc = nCodLancamento
End Property

Public Property Let nSequencia(nCodSequencia As Integer)
    nCodSeq = nCodSequencia
End Property
Public Property Let nComplemento(nCodComplemento As Integer)
    nCodComp = nCodComplemento
End Property

Public Property Let nAno(nAnoExercicio As Integer)
    nAnoExer = nAnoExercicio
End Property

Public Property Let nCodRed(nCodigoReduzido As Long)
    nCodReduzido = nCodigoReduzido
End Property

Private Sub cmdDup_Click()
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim nValorTaxa As Double

If lblDup.Caption = "Não" Then
   Exit Sub
End If

For x = 0 To 5
    If x <> 4 Then
       Panel1(x).Enabled = False
   End If
Next

grdDup.Rows = 1

Sql = "SELECT CODREDUZIDO,VALORPAGO,DATAPAGAMENTO,DATARECEBIMENTO,CODBANCO,RESTITUIDO,NUMDOCUMENTO FROM DEBITOPAGO "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO = " & nAnoExer
Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nCodSeq
Sql = Sql & " AND CODCOMPLEMENTO=" & nCodComp
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      Do Until .EOF
            
           If IsDate(lblDataPagto.Caption) Then
              If Not IsNull(!NumDocumento) Then
                 'grdDup.AddItem Format(!DataPagamento, "dd/mm/yyyy") & Chr(9) & Format(!DATARECEBIMENTO, "dd/mm/yyyy") & Chr(9) & FormatNumber(!ValorPago + CDbl(lblValorTaxa.Caption), 2) & Chr(9) & IIf(IsDate(!RESTITUIDO), Format(!RESTITUIDO, "dd/mm/yyyy"), "") & Chr(9) & !NumDocumento & "-" & RetornaDVNumDoc(!NumDocumento)
                 grdDup.AddItem Format(!DataPagamento, "dd/mm/yyyy") & Chr(9) & Format(!datarecebimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(!ValorPago, 2) & Chr(9) & IIf(IsDate(!RESTITUIDO), Format(!RESTITUIDO, "dd/mm/yyyy"), "") & Chr(9) & !NumDocumento & "-" & RetornaDVNumDoc(!NumDocumento)
              End If
           Else
              Sql = "SELECT VALORTAXADOC FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & Val(SubNull(!NumDocumento))
              Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
              With RdoAux2
                  If .RowCount > 0 Then
                      nValorTaxa = !ValorTaxaDoc
                  End If
                 .Close
              End With
              If nValorTaxa = 0 Then
                'PODE SER QUE VEIO DA RUIM-APD ENTÃO TEMOS QUE VER SE A TAXA EXP ESTA NA PARCELA
                Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO "
                Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO = " & nAnoExer
                Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nCodSeq
                Sql = Sql & " AND CODCOMPLEMENTO=" & nCodComp & " AND CODTRIBUTO=3"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        nValorTaxa = !ValorTributo
                    End If
                   .Close
                End With
              End If
              If Not IsNull(!NumDocumento) Then
                 grdDup.AddItem Format(!DataPagamento, "dd/mm/yyyy") & Chr(9) & Format(!datarecebimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(!ValorPago + nValorTaxa, 2) & Chr(9) & IIf(IsDate(!RESTITUIDO), Format(!RESTITUIDO, "dd/mm/yyyy"), "") & Chr(9) & !NumDocumento & "-" & RetornaDVNumDoc(!NumDocumento)
              Else
                 grdDup.AddItem Format(!DataPagamento, "dd/mm/yyyy") & Chr(9) & Format(!datarecebimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(!ValorPago + nValorTaxa, 2) & Chr(9) & IIf(IsDate(!RESTITUIDO), Format(!RESTITUIDO, "dd/mm/yyyy"), "")
              End If
           End If
          .MoveNext
      Loop
     .Close
End With

frDup.Visible = True
End Sub

Private Sub cmdPrint_Click()
Dim Sql As String, x As Integer
'SaveFormImageToFile Me, Picture1, sPathBin & "\frmParcela.bmp"
'x = Shell("MSPAINT" & " " & sPathBin & "\frmParcela.bmp", vbNormalFocus)

'DELETA TEMPORARIO
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

With grdTrib
    For x = 2 To grdTrib.Rows - 1
        Sql = "INSERT DAM(COMPUTER,SEQ,FULLTRIB,PRINCIPAL,JUROS,MULTA,CORRECAO) VALUES('" & NomeDoUsuario & "'," & x & ",'" & .TextMatrix(x, 0) & "',"
        Sql = Sql & Virg2Ponto(RemovePonto(.TextMatrix(x, 1))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 2))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 3))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 4))) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

frmReport.ShowReport "PARCELA", frmMdi.hwnd, Me.hwnd

'DELETA TEMPORARIO
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSairDup_Click()
For x = 0 To 5
    If x <> 4 Then
    Panel1(x).Enabled = True
    End If
Next
frDup.Visible = False

End Sub

Private Sub Form_Activate()
Me.ZOrder 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyF9 Then
     KeyAscii = 0
     Unload Me
End If

End Sub

Private Sub Form_Load()
Dim nIndex As Long

Me.Caption = Me.Caption & Format(nNumParc, "00") & " - Exercício: " & sTr(nAnoExer)
frDup.Visible = False
Ocupado
CarregaParcela2
lblContrib.Caption = Format(frmDebitoImob.txtCod.Text, "000000") & "-" & frmDebitoImob.lblProp.Caption
Liberado

nIndex = frmDebitoImob.m_cMenuOpcoes.IndexForKey("mnuDA")
If frmDebitoImob.m_cMenuOpcoes.Enabled(nIndex) = True Then
    txtLivro.Locked = False
    txtPagina.Locked = False
    txtCertidao.Locked = False
    txtValorPago.Locked = False
    txtValorDiferenca.Locked = False
    mskIncricao.Locked = False
End If

End Sub

Private Sub CarregaParcela2()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery
Dim x As Integer, nSomaL As Double, nSomaJ As Double, nSomaM As Double, nSomaC As Double, bAchou As Boolean
Dim nValorCorrecao As Double, nValorTaxa As Double, nStatus As Integer, nValorTaxaFromLanc As Double, bDA As Boolean, nValorAtual As Double
Dim bPago As Boolean, bVeioDaSmar As Boolean, bJuros As Boolean, bMulta As Boolean, nValorMulta As Double, nValorJuros As Double

txtValorDiferenca.Text = "X": lblValorTaxa.Caption = "X"

'CARREGA O EXTRATO
Set qd.ActiveConnection = cn
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = nCodReduzido
qd(1) = nCodReduzido
qd(2) = nAnoExer
qd(3) = nAnoExer
qd(4) = nCodLanc
qd(5) = nCodLanc
qd(6) = nCodSeq
qd(7) = nCodSeq
qd(8) = nNumParc
qd(9) = nNumParc
qd(10) = nCodComp
qd(11) = nCodComp
qd(12) = 1
qd(13) = 99
qd(14) = Format(dDataAtualiza, "mm/dd/yyyy")
qd(15) = NomeDoUsuario
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
    lblStatus.Caption = Format(!statuslanc, "00") & " - DÉBITO " & !Situacao
    lblDataVencto.Caption = Format(!DataVencimento, "dd/mm/yyyy")
    lblDataVenctoCalc.Caption = Format(!DataVencimentoCalc, "dd/mm/yyyy")
    lblDataBase.Caption = Format(!DATADEBASE, "dd/mm/yyyy")
    txtLivro.Text = Format(!NUMLIVRO, "000000")
    txtPagina.Text = Format(!PAGINA, "000000")
    txtCertidao.Text = Format(!CERTIDAO, "000000")
    mskIncricao.Text = IIf(IsNull(!datainscricao), "00/00/0000", Format(!datainscricao, "dd/mm/yyyy"))
    nValorAtual = 0
    lblIsentoMJ.Caption = "Não"
    lblDesconto.Caption = Val(SubNull(!PERCDESCONTO)) & " %"
    lblAjuizamento.Caption = IIf(IsNull(!dataajuiza), "00/00/0000", Format(!dataajuiza, "dd/mm/yyyy"))
    Do Until .EOF
         bAchou = False
         For x = 1 To grdTrib.Rows - 1
              If !CodTributo = Val(Left(grdTrib.TextMatrix(x, 0), 3)) Then
                  bAchou = True
              End If
         Next
         If Not bAchou Then
            nValorJuros = !ValorJuros
            nValorMulta = !ValorMulta
            nValorCorrecao = !ValorCorrecao
            If Not IsNull(!NumDocumento) Then
                Sql = "SELECT NUMDOCUMENTO,ISENTOMJ,PERCISENCAO FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & !NumDocumento
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux2.RowCount > 0 Then
                    If Val(SubNull(RdoAux2!isentomj)) > 0 Then
                        lblIsentoMJ = "Sim"
'                        nValorJuros = 0
'                        nValorMulta = 0
                    End If
                    If Not IsNull(RdoAux2!percisencao) Then
                        lblDesconto.Caption = FormatNumber(RdoAux2!percisencao, 2)
                    Else
                        lblDesconto.Caption = "0,00%"
                    End If
                End If
                RdoAux2.Close
            End If
            grdTrib.AddItem Format(!CodTributo, "00") & " - " & !ABREVTRIBUTO & Chr(9) & FormatNumber(!ValorTributo, 2) & Chr(9) & _
            FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorMulta, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2)
            nValorAtual = nValorAtual + !ValorTotal
         End If
         
        .MoveNext
    Loop
   .Close
End With

nStatus = Val(Left(lblStatus.Caption, 2))
lblLanc.Caption = frmDebitoImob.grdExtrato.CellText(frmDebitoImob.grdExtrato.SelectedRow, 2)
nSomaL = 0
For x = 2 To grdTrib.Rows - 1
       nSomaL = nSomaL + grdTrib.TextMatrix(x, 1)
       nSomaJ = nSomaJ + grdTrib.TextMatrix(x, 2)
       nSomaM = nSomaM + grdTrib.TextMatrix(x, 3)
       nSomaC = nSomaC + grdTrib.TextMatrix(x, 4)
Next
lblTotJ.Caption = FormatNumber(nSomaJ, 2)
lblTotM.Caption = FormatNumber(nSomaM, 2)
lblTotC.Caption = FormatNumber(nSomaC, 2)
lblTotL.Caption = FormatNumber(nSomaL, 2)
lblValorLancado.Caption = FormatNumber(nSomaL, 2)


Sql = "SELECT MIN(SEQPAG) AS MINIMO From DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO=" & nAnoExer
Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nCodSeq & " AND CODCOMPLEMENTO=" & nCodComp & "  AND RESTITUIDO IS  NULL"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
      If Not IsNull(!MINIMO) Then
          Sql = "SELECT * From DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO=" & nAnoExer
          Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nCodSeq & " AND CODCOMPLEMENTO=" & nCodComp
          Sql = Sql & " AND SEQPAG=" & !MINIMO
          Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          With RdoAux3
               lblBanco.Caption = Format(!CodBanco, "000")
               lblNumDoc.Caption = SubNull(!NumDocumento)
               If Not IsNull(RdoAux3!VALORDIF) Then
                   txtValorDiferenca.Text = FormatNumber(RdoAux3!VALORDIF, 2)
               End If
               If Not IsNull(RdoAux3!ValorTarifa) Then
                   lblValorTaxa.Caption = FormatNumber(RdoAux3!ValorTarifa, 2)
               End If
               If !DataPagamento <> CDate("01/01/1900") And Not IsNull(!DataPagamento) Then
                  bPago = True
                  If nCodLanc = 5 Then
                     If Not IsNull(!valorpagoreal) And Val(SubNull(!valorpagoreal)) > 0 Then
                          txtValorPago.Text = FormatNumber(!valorpagoreal, 2)
                     Else
                          txtValorPago.Text = FormatNumber(!ValorPago, 2)
                     End If
                  Else
                      If Not IsNull(!valorpagoreal) And Val(SubNull(!valorpagoreal)) > 0 Then
                           txtValorPago.Text = FormatNumber(!valorpagoreal, 2)
                      Else
                           txtValorPago.Text = FormatNumber(!ValorPago, 2)
                      End If
                  End If
                  lblDataPagto.Caption = Format(!DataPagamento, "dd/mm/yyyy")
               Else
                  bPago = False
                  txtValorPago.Text = "0,00"
                  lblDataPagto.Caption = ""
               End If
               If Not IsNull(!DATAPAGAMENTOCALC) Then
                  lblDataPagtoCalc.Caption = Format(!DATAPAGAMENTOCALC, "dd/mm/yyyy")
               Else
                  lblDataPagtoCalc.Caption = lblDataPagto.Caption
               End If
               If !datarecebimento <> CDate("01/01/1900") Then
                  lblDataReceita.Caption = Format(!datarecebimento, "dd/mm/yyyy")
               Else
                  lblDataReceita.Caption = ""
               End If
              .Close
          End With
      Else
          Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
          Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
          Sql = Sql & "WHERE parceladocumento.codreduzido = " & nCodReduzido & " AND parceladocumento.anoexercicio = " & nAnoExer & " AND parceladocumento.codlancamento = " & nCodLanc & " AND "
          Sql = Sql & "parceladocumento.seqlancamento = " & nCodSeq & " AND parceladocumento.numparcela = " & nNumParc & " AND parceladocumento.codcomplemento = " & nCodComp & " AND "
          Sql = Sql & "numdocumento.valorpago > 0"
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          With RdoAux2
              If .RowCount > 0 Then
                  lblNumDoc.Caption = Val(SubNull(!NumDocumento))
                  txtValorPago.Text = FormatNumber(!ValorPago, 2)
              End If
             .Close
          End With
      End If
     .Close
End With

If nStatus = 1 Or nStatus = 2 Or nStatus = 7 Or nStatus = 9 Then
    Sql = "SELECT PARCELADOCUMENTO.NUMDOCUMENTO,NUMDOCUMENTO.CODBANCO,NUMDOCUMENTO.CODAGENCIA,NUMDOCUMENTO.VALORPAGO, NumDocumento.VALORTAXADOC "
    Sql = Sql & "FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NumDocumento = NumDocumento.NumDocumento "
    Sql = Sql & "WHERE PARCELADOCUMENTO.CODREDUZIDO = " & nCodReduzido & " AND PARCELADOCUMENTO.ANOEXERCICIO = " & nAnoExer & " AND "
    Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = " & nCodLanc & " AND PARCELADOCUMENTO.NUMPARCELA = " & nNumParc & " AND NUMDOCUMENTO.VALORPAGO > 0 AND "
    Sql = Sql & "PARCELADOCUMENTO.SEQLANCAMENTO = " & nCodSeq & " AND PARCELADOCUMENTO.CODCOMPLEMENTO = " & nCodComp
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
         If .RowCount > 0 Then
'                    lblBanco.Caption = Format(!CODBANCO, "000")
             lblAgencia.Caption = Format(!CodAgencia, "0000")
         End If
        .Close
    End With
Else
    lblAgencia.Caption = Format(0, "0000")
End If
Sql = "SELECT NOMEREDUZ FROM BANCO WHERE "
Sql = Sql & "CODBANCO=" & Val(lblBanco.Caption)
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux2
    If .RowCount > 0 Then
       lblBanco.Caption = lblBanco.Caption & "-" & !NOMEREDUZ
    Else
       lblBanco.Caption = lblBanco.Caption & "-" & "*****"
    End If
   .Close
End With

If nStatus = 4 Then
    lblValorAtualizado.Caption = FormatNumber(nValorAtual, 2)
ElseIf nStatus = 1 Or nStatus = 2 Or nStatus = 7 Then
    lblValorAtualizado.Caption = FormatNumber(txtValorPago.Text, 2)
Else
    lblValorAtualizado.Caption = FormatNumber(CDbl(lblTotL.Caption) + CDbl(lblTotM.Caption + CDbl(lblTotJ.Caption) + CDbl(lblTotC.Caption)), 2)
End If


If Val(txtValorPago.Text) > 0 And Val(lblNumDoc.Caption) = 0 Then
   Sql = "SELECT PARCELADOCUMENTO.NUMDOCUMENTO FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON "
   Sql = Sql & "PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO WHERE CODREDUZIDO = " & nCodReduzido & " AND "
   Sql = Sql & "ANOEXERCICIO=" & nAnoExer & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND "
   Sql = Sql & "SEQLANCAMENTO=" & nCodSeq & " AND CODCOMPLEMENTO=" & nCodComp & " AND VALORPAGO>0"
   Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux2
       If .RowCount > 0 Then
         nNumDoc = !NumDocumento
       Else
         nNumDoc = 0
       End If
      .Close
   End With

   Sql = "SELECT NUMDOCUMENTO,NUMDOCUMENTO.VALORTAXADOC FROM NUMDOCUMENTO "
   Sql = Sql & "WHERE NUMDOCUMENTO=" & Val(nNumDoc)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount = 0 Then
          nValorTaxa = 0
          bVeioDaSmar = True
       Else
          If !ValorTaxaDoc > 0 Then
             bVeioDaSmar = False
             nValorTaxa = FormatNumber(!ValorTaxaDoc, 2)
          Else
             bVeioDaSmar = True
             Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO = " & nCodReduzido & " AND ANOEXERCICIO = " & nAnoExer & " AND CODLANCAMENTO = " & nCodLanc & " AND "
             Sql = Sql & "SEQLANCAMENTO = " & nCodSeq & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCodComp & " AND CODTRIBUTO=3"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                 If .RowCount > 0 Then
                    nValorTaxa = !ValorTributo
                 Else
                    nValorTaxa = 0
                 End If
                .Close
             End With
          End If
       End If
      .Close
   End With
   lblNumDoc.Caption = Format(nNumDoc, "00000000") & "-" & RetornaDVNumDoc(CLng(nNumDoc))
   If nValorTaxa = 0 Then nValorTaxa = nValorTaxaFromLanc
   If Not bVeioDaSmar Then
      txtValorPago.Text = FormatNumber(CDbl(txtValorPago.Text) + nValorTaxa, 2)
   Else
      txtValorPago.Text = FormatNumber(CDbl(txtValorPago.Text), 2)
   End If
   If lblValorTaxa.Caption = "X" Then
      lblValorTaxa.Caption = FormatNumber(nValorTaxa, 2)
   End If
   If txtValorDiferenca.Text = "X" Then
      txtValorDiferenca.Text = FormatNumber(CDbl(txtValorPago.Text) - (CDbl(lblValorAtualizado.Caption) + (nValorTaxa)), 2)
   End If
Else
     Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO = " & nCodReduzido & " AND ANOEXERCICIO = " & nAnoExer & " AND CODLANCAMENTO = " & nCodLanc & " AND "
     Sql = Sql & "SEQLANCAMENTO = " & nCodSeq & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCodComp & " AND CODTRIBUTO=3"
     Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
     With RdoAux2
          If .RowCount > 0 Then
              nValorTaxa = !ValorTributo
          Else
              nValorTaxa = 0
          End If
         .Close
     End With
     If lblValorTaxa.Caption = "X" Then
        lblValorTaxa.Caption = FormatNumber(nValorTaxa, 2)
     End If
     If txtValorDiferenca.Text = "X" Then
        If CDbl(txtValorPago.Text) > 0 Then
            txtValorDiferenca.Text = FormatNumber(CDbl(txtValorPago.Text) - (CDbl(lblValorAtualizado.Caption) + (nValorTaxa)), 2)
        Else
            txtValorDiferenca.Text = "0,00"
        End If
     End If
End If

If nStatus > 2 And nStatus <> 7 Then
    lblValorTaxa.Caption = "0,00"
End If


Sql = "SELECT * FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO = " & nAnoExer
Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nCodSeq
Sql = Sql & " AND CODCOMPLEMENTO=" & nCodComp
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
      If nStatus = 1 And Val(txtValorPago.Text) > 0 And nNumParc > 0 Then
         lblDup.Caption = "Sim"
      Else
         If .RowCount > 0 Then
            If .RowCount > 1 Then
                lblDup.Caption = "Sim"
                cmdDup.Enabled = True
            Else
                If IsDate(lblDataPagto.Caption) Then
                    lblDup.Caption = "Não"
                    cmdDup.Enabled = False
                Else
                    lblDup.Caption = "Sim"
                    cmdDup.Enabled = True
                End If
            End If
         Else
            lblDup.Caption = "Não"
            cmdDup.Enabled = False
         End If
      End If
     .Close
End With

Sql = "SELECT debitoparcela.codreduzido, usuario.nomecompleto FROM debitoparcela LEFT OUTER JOIN usuario ON debitoparcela.userid = usuario.Id WHERE CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO = " & nAnoExer
Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=" & nNumParc & " AND SEQLANCAMENTO=" & nCodSeq
Sql = Sql & " AND CODCOMPLEMENTO=" & nCodComp
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If Not IsNull(!NomeCompleto) Then
        Me.Caption = Me.Caption & " (Gerado por: " & !NomeCompleto & ")"
    End If
   .Close
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
nResp = 0
End Sub

Private Sub mskIncricao_GotFocus()
mskIncricao.SetFocus
End Sub

Private Sub mskIncricao_LostFocus()
Grava
End Sub

Private Sub txtCertidao_GotFocus()
txtCertidao.SelStart = 0
txtCertidao.SelLength = Len(txtCertidao.Text)
End Sub

Private Sub txtCertidao_KeyPress(KeyAscii As Integer)
Tweak txtCertidao, KeyAscii, IntegerPositive
End Sub

Private Sub txtCertidao_LostFocus()
Grava
txtCertidao.Text = Format(txtCertidao.Text, "000000")
End Sub

Private Sub txtLivro_GotFocus()
txtLivro.SelStart = 0
txtLivro.SelLength = Len(txtLivro.Text)
End Sub

Private Sub txtLivro_KeyPress(KeyAscii As Integer)
Tweak txtLivro, KeyAscii, IntegerPositive
End Sub

Private Sub txtLivro_LostFocus()
Grava
txtLivro.Text = Format(txtLivro.Text, "000000")
End Sub

Private Sub txtPagina_GotFocus()
txtPagina.SelStart = 0
txtPagina.SelLength = Len(txtPagina.Text)
End Sub

Private Sub txtPagina_KeyPress(KeyAscii As Integer)
Tweak txtPagina, KeyAscii, IntegerPositive
End Sub

Private Sub Grava()
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer
Dim nComp As Integer


If Not IsNumeric(txtValorDiferenca.Text) Then
    txtValorDiferenca.Text = "0"
End If

If Not IsNumeric(txtValorPago.Text) Then
    txtValorPago.Text = "0"
End If


With frmDebitoImob.grdExtrato
    nLinha = .SelectedRow
    nAno = .CellText(nLinha, 1)
    nLanc = Left$(.CellText(nLinha, 2), 3)
    nSeq = .CellText(nLinha, 3)
    nParc = .CellText(nLinha, 4)
    nCompl = .CellText(nLinha, 5)
    
    Sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & Val(txtLivro.Text) & " ,PAGINALIVRO=" & Val(txtPagina.Text) & " ,NUMCERTIDAO=" & Val(txtCertidao.Text)
    Sql = Sql & " WHERE CODREDUZIDO = " & Val(frmDebitoImob.txtCod.Text) & " AND ANOEXERCICIO = " & nAno & " AND CODLANCAMENTO = " & nLanc & " AND "
    Sql = Sql & "SEQLANCAMENTO = " & nSeq & " AND NUMPARCELA = " & nParc & " AND CODCOMPLEMENTO = " & nCompl
    cn.Execute Sql, rdExecDirect
    
    Sql = "UPDATE DEBITOPAGO SET VALORDIF=" & Virg2Ponto(RemovePonto(txtValorDiferenca.Text)) & " ,VALORPAGOREAL=" & Virg2Ponto(RemovePonto(txtValorPago.Text))
    Sql = Sql & " WHERE CODREDUZIDO = " & Val(frmDebitoImob.txtCod.Text) & " AND ANOEXERCICIO = " & nAno & " AND CODLANCAMENTO = " & nLanc & " AND "
    Sql = Sql & "SEQLANCAMENTO = " & nSeq & " AND NUMPARCELA = " & nParc & " AND CODCOMPLEMENTO = " & nCompl
    cn.Execute Sql, rdExecDirect
    
    If IsDate(mskIncricao.Text) Then
        Sql = "UPDATE DEBITOPARCELA SET DATAINSCRICAO='" & Format(mskIncricao.Text, "mm/dd/yyyy")
        Sql = Sql & "' WHERE CODREDUZIDO = " & Val(frmDebitoImob.txtCod.Text) & " AND ANOEXERCICIO = " & nAno & " AND CODLANCAMENTO = " & nLanc & " AND "
        Sql = Sql & "SEQLANCAMENTO = " & nSeq & " AND NUMPARCELA = " & nParc & " AND CODCOMPLEMENTO = " & nCompl
        cn.Execute Sql, rdExecDirect
    End If
End With

End Sub

Private Sub txtPagina_LostFocus()
Grava
txtPagina.Text = Format(txtPagina.Text, "000000")
End Sub

Private Sub txtValorDiferenca_GotFocus()
txtValorDiferenca.SelStart = 0
txtValorDiferenca.SelLength = Len(txtValorDiferenca.Text)
End Sub

Private Sub txtValorDiferenca_KeyPress(KeyAscii As Integer)
Tweak txtValorDiferenca, KeyAscii, DecimalPositive, 2
End Sub

Private Sub txtValorDiferenca_LostFocus()
Grava
txtValorDiferenca.Text = FormatNumber(txtValorDiferenca.Text, 2)
End Sub

Private Sub txtValorPago_GotFocus()
txtValorPago.SelStart = 0
txtValorPago.SelLength = Len(txtValorPago.Text)
End Sub

Private Sub txtValorPago_KeyPress(KeyAscii As Integer)
Tweak txtValorPago, KeyAscii, DecimalPositive, 2
End Sub

Private Sub txtValorPago_LostFocus()
Grava
txtValorPago.Text = FormatNumber(txtValorPago.Text, 2)
End Sub
