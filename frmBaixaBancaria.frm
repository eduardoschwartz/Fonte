VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmBaixaBancaria 
   BackColor       =   &H00ECE7EE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa bancária"
   ClientHeight    =   5925
   ClientLeft      =   1590
   ClientTop       =   2220
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11220
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12780
      Top             =   8730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtErro 
      BackColor       =   &H00800000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   37
      Text            =   "frmBaixaBancaria.frx":0000
      Top             =   7665
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Tributacao.jcFrames pnlDetalhe 
      Height          =   2895
      Left            =   7770
      Top             =   1350
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   5106
      FrameColor      =   6974058
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Detalhes do Arquivo"
      TextBoxHeight   =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ComboBox cmbDataCredito 
         Height          =   315
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   270
         Width           =   1335
      End
      Begin VB.CommandButton cmdErro 
         Caption         =   "..."
         Height          =   255
         Left            =   2940
         TabIndex        =   36
         ToolTipText     =   "Exibir lista de erros"
         Top             =   2610
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblValorEfetivo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   44
         Top             =   2370
         Width           =   1245
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Efetivo...:"
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
         Left            =   150
         TabIndex        =   43
         Top             =   2370
         Width           =   1845
      End
      Begin VB.Label lblErro 
         BackStyle       =   0  'Transparent
         Caption         =   "999 ERRO(S) ENCONTRADO(S)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total.....:"
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
         Index           =   3
         Left            =   150
         TabIndex        =   30
         Top             =   2070
         Width           =   1845
      End
      Begin VB.Label lblValorTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   29
         Top             =   2070
         Width           =   1250
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Registros.:"
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
         Left            =   180
         TabIndex        =   28
         Top             =   1770
         Width           =   1815
      End
      Begin VB.Label lblNumReg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   27
         Top             =   1770
         Width           =   1250
      End
      Begin VB.Label lblAS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   24
         Top             =   870
         Width           =   1250
      End
      Begin VB.Label lblDA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   23
         Top             =   1470
         Width           =   1250
      End
      Begin VB.Label lblAC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   22
         Top             =   1170
         Width           =   1250
      End
      Begin VB.Label lblDB 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   600
         Width           =   1250
      End
      Begin VB.Label lblDC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "99/99/9999"
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
         Left            =   1920
         TabIndex        =   20
         Top             =   300
         Visible         =   0   'False
         Width           =   1250
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Arq.Déb.Autom...:"
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
         Index           =   9
         Left            =   150
         TabIndex        =   19
         Top             =   1470
         Width           =   1815
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Arq.de Cobrança.:"
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
         Index           =   8
         Left            =   150
         TabIndex        =   18
         Top             =   1170
         Width           =   1785
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Arq.Simples.....:"
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
         Index           =   7
         Left            =   150
         TabIndex        =   17
         Top             =   885
         Width           =   1725
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Baixa...:"
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
         Index           =   6
         Left            =   150
         TabIndex        =   16
         Top             =   585
         Width           =   1815
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Crédito.:"
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
         Index           =   5
         Left            =   150
         TabIndex        =   15
         Top             =   300
         Width           =   1875
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   1635
      Left            =   60
      Top             =   4260
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   2884
      FrameColor      =   6974058
      TextBoxColor    =   13302261
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Conteúdo do documento selecionado"
      TextBoxHeight   =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSFlexGridLib.MSFlexGrid grdParc 
         Height          =   1275
         Left            =   60
         TabIndex        =   26
         Top             =   270
         Width           =   10995
         _ExtentX        =   19394
         _ExtentY        =   2249
         _Version        =   393216
         Rows            =   8
         Cols            =   17
         FixedCols       =   0
         BackColor       =   15658734
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
         BackColorSel    =   192
         ForeColorSel    =   16777215
         BackColorBkg    =   15658734
         GridColor       =   8421504
         GridColorFixed  =   14737632
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmBaixaBancaria.frx":0007
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   2895
      Left            =   60
      Top             =   1350
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5106
      FrameColor      =   6974058
      TextBoxColor    =   13302261
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Documentos disponíveis no arquivo"
      TextBoxHeight   =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSFlexGridLib.MSFlexGrid grdReg 
         Height          =   2505
         Left            =   90
         TabIndex        =   25
         Top             =   330
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   4419
         _Version        =   393216
         Rows            =   10
         Cols            =   14
         FixedCols       =   0
         BackColor       =   15658734
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
         BackColorSel    =   192
         ForeColorSel    =   16777215
         BackColorBkg    =   15658734
         GridColor       =   8421504
         GridColorFixed  =   14737632
         FocusRect       =   0
         GridLinesFixed  =   1
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   $"frmBaixaBancaria.frx":00B7
      End
   End
   Begin Tributacao.jcFrames pnlCampo 
      Height          =   1275
      Left            =   5760
      Top             =   60
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   2249
      FrameColor      =   6974058
      TextBoxColor    =   13302261
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Dados do Arquivo"
      TextBoxHeight   =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin Tributacao.XP_ProgressBar PBar 
         Height          =   240
         Left            =   1440
         TabIndex        =   46
         Top             =   630
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   16777215
         Scrolling       =   1
         ShowText        =   -1  'True
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE7EE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   930
         Width           =   5145
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Progresso.:"
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
         Left            =   90
         TabIndex        =   31
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label lblBanco 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "."
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
         Left            =   1230
         TabIndex        =   14
         Top             =   300
         Width           =   4005
      End
      Begin VB.Label label1 
         BackColor       =   &H00E4C6BA&
         BackStyle       =   0  'Transparent
         Caption         =   "Banco.....:"
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
         Left            =   90
         TabIndex        =   13
         Top             =   300
         Width           =   1185
      End
   End
   Begin Tributacao.jcFrames pnlArquivo 
      Height          =   1275
      Left            =   2820
      Top             =   60
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   2249
      FrameColor      =   6974058
      TextBoxColor    =   13302261
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Arquivos Disponíveis"
      TextBoxHeight   =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ComboBox lstArq 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   390
         Width           =   2715
      End
      Begin prjChameleon.chameleonButton cmdLoad 
         Height          =   345
         Left            =   1530
         TabIndex        =   34
         ToolTipText     =   "Ler arquivo selecionado"
         Top             =   780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Carregar"
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
         MICON           =   "frmBaixaBancaria.frx":0196
         PICN            =   "frmBaixaBancaria.frx":01B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdOpcoes 
         Height          =   345
         Left            =   210
         TabIndex        =   38
         ToolTipText     =   "Ler arquivo selecionado"
         Top             =   780
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Opções"
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
         MICON           =   "frmBaixaBancaria.frx":02C4
         PICN            =   "frmBaixaBancaria.frx":02E0
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
   Begin Tributacao.jcFrames pnlManual 
      Height          =   4185
      Left            =   7770
      Top             =   60
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7382
      FrameColor      =   6974058
      TextBoxColor    =   13302261
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Dados para Baixa Manual"
      TextBoxHeight   =   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   16744576
      ColorTo         =   16744576
      Begin VB.TextBox txtNumDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         MaxLength       =   9
         TabIndex        =   1
         Top             =   510
         Width           =   1185
      End
      Begin VB.TextBox txtAgencia 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1890
         Width           =   1185
      End
      Begin VB.TextBox txtValorPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   2250
         Width           =   1185
      End
      Begin VB.TextBox txtBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H00ECE7EE&
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1530
         Width           =   1185
      End
      Begin esMaskEdit.esMaskedEdit mskDataPag 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   860
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         MouseIcon       =   "frmBaixaBancaria.frx":0380
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
      Begin esMaskEdit.esMaskedEdit mskDataCred 
         Height          =   285
         Left            =   1920
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         BackColor       =   15525870
         MouseIcon       =   "frmBaixaBancaria.frx":039C
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
         Locked          =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdAdd 
         Height          =   345
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Adicionar Débito ao Grid"
         Top             =   2910
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Adicionar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "frmBaixaBancaria.frx":03B8
         PICN            =   "frmBaixaBancaria.frx":03D4
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
         Left            =   1740
         TabIndex        =   39
         ToolTipText     =   "Retornar a tela de origem"
         Top             =   3360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Voltar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "frmBaixaBancaria.frx":052E
         PICN            =   "frmBaixaBancaria.frx":054A
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
         Left            =   1740
         TabIndex        =   40
         ToolTipText     =   "Remover Débito do Grid"
         Top             =   2910
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Remover"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "frmBaixaBancaria.frx":05B8
         PICN            =   "frmBaixaBancaria.frx":05D4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   1
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdBaixa 
         Height          =   345
         Left            =   240
         TabIndex        =   41
         ToolTipText     =   "Efetuar baixa nos documentos"
         Top             =   3360
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Executar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         MICON           =   "frmBaixaBancaria.frx":072E
         PICN            =   "frmBaixaBancaria.frx":074A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento...............:"
         Height          =   225
         Index           =   10
         Left            =   270
         TabIndex        =   12
         Top             =   570
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Pago................:"
         Height          =   225
         Index           =   5
         Left            =   270
         TabIndex        =   11
         Top             =   2280
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Agência....................:"
         Height          =   225
         Index           =   4
         Left            =   270
         TabIndex        =   10
         Top             =   1935
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Crédito.........:"
         Height          =   225
         Index           =   3
         Left            =   270
         TabIndex        =   9
         Top             =   1260
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco.......................:"
         Height          =   225
         Index           =   1
         Left            =   270
         TabIndex        =   8
         Top             =   1590
         Width           =   1605
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data de Pagamento..:"
         Height          =   225
         Index           =   0
         Left            =   270
         TabIndex        =   7
         Top             =   915
         Width           =   1605
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   2850
      Left            =   90
      TabIndex        =   45
      Top             =   5940
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   5027
      _Version        =   393216
      Rows            =   1
      Cols            =   21
      FixedCols       =   0
      BackColor       =   15658734
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   192
      ForeColorSel    =   16777215
      BackColorBkg    =   15658734
      GridColor       =   8421504
      GridColorFixed  =   14737632
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmBaixaBancaria.frx":0AEF
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      BackColor       =   &H00E4C6BA&
      BackStyle       =   0  'Transparent
      Caption         =   "BAIXA MANUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   3000
      TabIndex        =   42
      Top             =   330
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.Image imgLogo 
      Height          =   975
      Left            =   150
      Stretch         =   -1  'True
      Top             =   210
      Width           =   2505
   End
End
Attribute VB_Name = "frmBaixaBancaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenu As cPopupMenu
Attribute m_cMenu.VB_VarHelpID = -1

Private Type Registro
    nNumDoc As Long
    nSeq As Integer
    sDataDoc As String
    sDataPag As Date
    sDataCred As Date
    nValorPago As Double
    sAgencia As String
    nValorTarifa As Double
    sSitRetorno As String
    bExiste As Boolean
    bIsentoMJ As Boolean
    sCNPJ As String
    nAno As Integer
    nMes As Integer
    sDataVencto As Date
    nValorTarifaBancaria As Double
    nSomaTributo As Double
    sDataPagCalc As Date
    sConta As String
End Type

Private Type Documento
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    sDataVencto As String
    sSit As String
    nNumeroLivro As Integer
    nPaginaLivro As Integer
    bAjuizado As Boolean
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
    nValorTarifa As Double
    nValorDif As Double
    nValorCompensado As Double
    sBx As String
    sDp As String
    nSeqReg As Integer
    bExiste As Boolean
    sCNPJ As String
End Type

Private Type TRIBUTO
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTrib As Integer
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
    nValorTarifa As Double
    nValorCompensado As Double
    sAj As String
    sDA As String
    nFicha As Long
    nFichaJM As Long
    nFichaC As Long
End Type

Private Type TributoFicha
    nCodTrib As Integer
    sAbrevTrib As String
    Ficha As Long
    FichaJrMulta As Long
    FichaDivida As Long
    FichaDaJrMul As Long
    FichaDaEnca As Long
    FichaAjuiza As Long
    FichaAjJrMul As Long
    FichaAjEnca As Long
End Type

Private Type Ficha
    Ficha As Long
    Natureza As String
    Desc As String
    Vinculo As String
    Perc As Double
End Type

Private Type TributoProp
    nCodTrib As Integer
    nValorTrib As Double
    nPerc As Double
    nNovoValor As Double
End Type

Dim aRegistro() As Registro, aDoc() As Documento, aTrib() As TRIBUTO, aTribF() As TributoFicha, aFicha() As Ficha, sSeqArq As String

Private Sub cmbDataCredito_Click()
CarregaDataCredito
End Sub

Private Sub cmdAdd_Click()
Dim bAchou As Boolean, x As Integer, nNumDoc As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nValorTaxa As Double

If Val(txtNumDoc.Text) = 0 Then
   MsgBox "Digite o nº do documento.", vbExclamation, "Atenção"
   Exit Sub
End If
nNumDoc = Val(Left$(txtNumDoc.Text, Len(txtNumDoc.Text) - 1))
If Val(Right$(txtNumDoc.Text, 1)) <> RetornaDVNumDoc(nNumDoc) Then
    MsgBox "Digito Verificador Inválido", vbExclamation, "Atenção"
    Exit Sub
End If
If Not IsDate(mskDataPag.Text) Then
   MsgBox "Data de pagamento inválido.", vbExclamation, "Atenção"
   Exit Sub
End If
If Not IsDate(mskDataCred.Text) Then
   MsgBox "Data de crédito inválido.", vbExclamation, "Atenção"
   Exit Sub
End If
lblDC.Caption = mskDataCred.Text
If Val(txtValorPago.Text) = 0 Then
   MsgBox "Digite o valor pago.", vbExclamation, "Atenção"
   Exit Sub
End If
If CDate(mskDataCred.Text) < CDate(mskDataPag.Text) Then
   MsgBox "Data de crédito não pode ser menor que a data de pagamento.", vbExclamation, "Atenção"
   mskDataCred.SetFocus
   Exit Sub
End If
If Val(txtBanco.Text) = 0 Then
   MsgBox "Banco não selecionado.", vbExclamation, "Atenção"
   Exit Sub
End If

Sql = "SELECT numdocumento.* FROM parceladocumento INNER JOIN "
Sql = Sql & "numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
Sql = Sql & "Where NumDocumento.NumDocumento = " & nNumDoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Nº de documento não localizado no sistema.", vbExclamation, "Atenção"
        Exit Sub
    End If
End With

bAchou = False
For x = 1 To grdReg.Rows - 1
    If Val(grdReg.TextMatrix(x, 0)) = nNumDoc Then
        bAchou = True
        Exit For
    End If
Next
If bAchou Then
    MsgBox "Documento já inserido no grid.", vbInformation, "Atenção"
Else

    Sql = "SELECT NUMDOCUMENTO,VALORPAGO,VALORTAXADOC FROM NUMDOCUMENTO "
    Sql = Sql & " WHERE NUMDOCUMENTO=" & nNumDoc
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nValorTaxa = Val(SubNull(!ValorTaxaDoc))
        If !ValorPago > 0 Then
            If MsgBox("Já foi efetuado a baixa para este Documento." & vbCrLf & "A baixa será entendida como um pagamento em duplicidade." & vbCrLf & "Deseja Continuar ?", vbQuestion + vbYesNo, "Atenção") = vbNo Then
               Exit Sub
            End If
        End If
       .Close
    End With

    ReDim Preserve aRegistro(UBound(aRegistro) + 1)
    aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
    aRegistro(UBound(aRegistro)).nSeq = grdReg.Rows - 1
    aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
    aRegistro(UBound(aRegistro)).sDataPag = mskDataPag.Text
    aRegistro(UBound(aRegistro)).sDataCred = mskDataCred.Text
    aRegistro(UBound(aRegistro)).nValorPago = txtValorPago.Text
    aRegistro(UBound(aRegistro)).sAgencia = txtAgencia.Text
    aRegistro(UBound(aRegistro)).nValorTarifa = FormatNumber(nValorTaxa, 2)
    aRegistro(UBound(aRegistro)).sSitRetorno = "00-BAIXA NORMAL"
    aRegistro(UBound(aRegistro)).bExiste = True
    aRegistro(UBound(aRegistro)).bIsentoMJ = IIf(Val(SubNull(RdoAux!isentomj)) = 0, False, True)
    
    With aRegistro(UBound(aRegistro))
        grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & .sDataDoc & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
        Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
        IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & .nValorTarifa & Chr(9) & .nSeq
    End With
    
    CarregaParcela nNumDoc, grdReg.TextMatrix(grdReg.Rows - 1, 11), mskDataCred.Text
    CarregaDocumento grdReg.TextMatrix(grdReg.Row, 0), grdReg.TextMatrix(grdReg.Row, 11), grdReg.TextMatrix(grdReg.Row, 9)
    txtNumDoc.Text = "":  LimpaMascara mskDataPag: txtAgencia.Text = "": txtValorPago.Text = ""
    
End If
End Sub


Private Sub cmdBaixa_Click()
If lblErro.Visible Then
    MsgBox "Não é possivel efetuar baixa enquanto houver documentos com erro.", vbCritical, "Atenção"
    Exit Sub
End If

EfetuaBaixa
End Sub

Private Sub cmdErro_Click()
Dim x As Integer, bExec As Boolean
With txtErro
    If .Visible = False Then
        cmdOpcoes.Enabled = False: cmdLoad.Enabled = False: lstArq.Enabled = False
        .Top = 150: .Left = 6500: .Height = 4000: .Width = 4000
        .Visible = True
        .Text = "Nome do Arquivo: " & lstArq.Text & vbCrLf
        .Text = .Text & "Banco: " & lblBanco.Caption & vbCrLf
        .Text = .Text & "Data de crédito: " & lblDC.Caption & vbCrLf
        .Text = .Text & "Nº de erros: " & lblErro.Caption & vbCrLf
        .Text = .Text & "---------------------------------------------------------------------" & vbCrLf & vbCrLf
        bExec = True
        For x = 1 To UBound(aRegistro)
            If aRegistro(x).bExiste = False Then
                If lblAS.Caption = "N" Then
                    If bExec Then
                        .Text = .Text & "Documentos não Encontrados:" & vbCrLf
                        .Text = .Text & "--------------------------------------------------" & vbCrLf
                        bExec = False
                    End If
                    .Text = .Text & Format(aRegistro(x).nNumDoc, "000000000") & "-" & RetornaDVNumDoc(aRegistro(x).nNumDoc) & vbCrLf
                Else
                    If bExec Then
                        .Text = .Text & "CNPJ não Cadastrado:" & vbCrLf
                        .Text = .Text & "--------------------------------------------------" & vbCrLf
                        bExec = False
                    End If
                    .Text = .Text & Format(aRegistro(x).sCNPJ, "0#\.###\.###/####-##") & vbCrLf
                End If
            End If
        Next
        bExec = True
        For x = 1 To UBound(aDoc)
            If aDoc(x).bExiste = False Then
                If bExec Then
                    .Text = .Text & vbCrLf & "Documentos sem lançamentos:" & vbCrLf
                    .Text = .Text & "--------------------------------------------------" & vbCrLf
                    bExec = False
                End If
                .Text = .Text & Format(aDoc(x).nNumDoc, "000000000") & "-" & RetornaDVNumDoc(aDoc(x).nNumDoc) & vbCrLf
            End If
        Next
    Else
        .Visible = False
        cmdOpcoes.Enabled = True: cmdLoad.Enabled = True: lstArq.Enabled = True
    End If
End With
End Sub

Private Sub cmdLoad_Click()
If bLocal Then
    Exit Sub
End If


Dim Sql As String, RdoAux As rdoResultset
If lstArq.ListIndex > -1 Then
    cmdOpcoes.Enabled = False: cmdLoad.Enabled = False: lstArq.Enabled = False
    LimpaTela
    LeArquivo
   
    If grdReg.Rows = 1 Then
        
        MsgBox "Não existem registros para baixar, verifique o arquivo.", vbExclamation, "Atenção"
        GoTo FIM
    End If
   
    '*** VERIFICA BAIXA NO ARQUIVO ***
    If lblDC.Caption <> "Sem Baixa" Then
        Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & lstArq.Text & "' AND DATACREDITO='" & Format(lblDC.Caption, "mm/dd/yyyy") & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If IsNull(!DATABAIXA) Then
                    lblDB.Caption = "Sem Baixa"
                Else
                    lblDB.Caption = Format(!DATABAIXA, "dd/mm/yyyy")
                End If
            Else
                lblDB.Caption = "Sem Baixa"
            End If
           .Close
        End With
    Else
        lblDB.Caption = "Sem Baixa"
    End If
   
    If grdReg.Rows > 1 Then
         grdReg.ColSel = 10
         If grdReg.TextMatrix(1, 11) <> "" Then
            CarregaDocumento grdReg.TextMatrix(1, 0), grdReg.TextMatrix(1, 11), grdReg.TextMatrix(1, 9)
         End If
    End If
FIM:
    cmdOpcoes.Enabled = True: cmdLoad.Enabled = True: lstArq.Enabled = True
End If
End Sub

Private Sub cmdOpcoes_Click()

If bLocal Then
    Exit Sub
End If

lIndex = m_cMenu.ShowPopupMenu(cmdOpcoes.Left + pnlArquivo.Left, cmdOpcoes.Top + 400, cmdOpcoes.Left, cmdOpcoes.Top, Me.ScaleWidth - cmdOpcoes.Left - cmdOpcoes.Width, cmdOpcoes.Top + cmdOpcoes.Height, False)
End Sub

Private Sub cmdRemover_Click()
Dim nNumDoc As Long, nLinha As Integer, aDocTmp() As Documento, aRegistroTmp() As Registro
ReDim aDocTmp(0): ReDim aRegistroTmp(0)

If grdReg.Rows = 1 Then
    MsgBox "Não existem documentos.", vbCritical, "Atenção"
Else
    If grdReg.Row = 0 Then
        MsgBox "Selecione o documento que deseja remover.", vbCritical, "Atenção"
    Else
        nNumDoc = Val(grdReg.TextMatrix(grdReg.Row, 0))
        
        For nLinha = 1 To UBound(aRegistro)
            If aRegistro(nLinha).nNumDoc <> nNumDoc Then
                ReDim Preserve aRegistroTmp(UBound(aRegistroTmp) + 1)
                aRegistroTmp(UBound(aRegistroTmp)).nNumDoc = aRegistro(nLinha).nNumDoc
                aRegistroTmp(UBound(aRegistroTmp)).sDataDoc = aRegistro(nLinha).sDataDoc
                aRegistroTmp(UBound(aRegistroTmp)).sDataPag = aRegistro(nLinha).sDataPag
                aRegistroTmp(UBound(aRegistroTmp)).sDataCred = aRegistro(nLinha).sDataCred
                aRegistroTmp(UBound(aRegistroTmp)).nValorPago = aRegistro(nLinha).nValorPago
                aRegistroTmp(UBound(aRegistroTmp)).sAgencia = aRegistro(nLinha).sAgencia
                aRegistroTmp(UBound(aRegistroTmp)).nValorTarifa = aRegistro(nLinha).nValorTarifa
                aRegistroTmp(UBound(aRegistroTmp)).sSitRetorno = aRegistro(nLinha).sSitRetorno
                aRegistroTmp(UBound(aRegistroTmp)).bExiste = aRegistro(nLinha).bExiste
                aRegistroTmp(UBound(aRegistroTmp)).bIsentoMJ = aRegistro(nLinha).bIsentoMJ
            End If
        Next
        ReDim aRegistro(0)
        For nLinha = 1 To UBound(aRegistroTmp)
                ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                aRegistro(UBound(aRegistro)).nNumDoc = aRegistroTmp(nLinha).nNumDoc
                aRegistro(UBound(aRegistro)).sDataDoc = aRegistroTmp(nLinha).sDataDoc
                aRegistro(UBound(aRegistro)).sDataPag = aRegistroTmp(nLinha).sDataPag
                aRegistro(UBound(aRegistro)).sDataCred = aRegistroTmp(nLinha).sDataCred
                aRegistro(UBound(aRegistro)).nValorPago = aRegistroTmp(nLinha).nValorPago
                aRegistro(UBound(aRegistro)).sAgencia = aRegistroTmp(nLinha).sAgencia
                aRegistro(UBound(aRegistro)).nValorTarifa = aRegistroTmp(nLinha).nValorTarifa
                aRegistro(UBound(aRegistro)).sSitRetorno = aRegistroTmp(nLinha).sSitRetorno
                aRegistro(UBound(aRegistro)).bExiste = aRegistroTmp(nLinha).bExiste
                aRegistro(UBound(aRegistro)).bIsentoMJ = aRegistroTmp(nLinha).bIsentoMJ
        Next
        
        
        For nLinha = 1 To UBound(aDoc)
            If aDoc(nLinha).nNumDoc <> nNumDoc Then
                ReDim Preserve aDocTmp(UBound(aDocTmp) + 1)
                aDocTmp(UBound(aDocTmp)).nNumDoc = aDoc(nLinha).nNumDoc
                aDocTmp(UBound(aDocTmp)).nCodReduz = aDoc(nLinha).nCodReduz
                aDocTmp(UBound(aDocTmp)).nAno = aDoc(nLinha).nAno
                aDocTmp(UBound(aDocTmp)).nLanc = aDoc(nLinha).nLanc
                aDocTmp(UBound(aDocTmp)).nSeq = aDoc(nLinha).nSeq
                aDocTmp(UBound(aDocTmp)).nParc = aDoc(nLinha).nParc
                aDocTmp(UBound(aDocTmp)).nCompl = aDoc(nLinha).nCompl
                aDocTmp(UBound(aDocTmp)).sDataVencto = aDoc(nLinha).sDataVencto
                aDocTmp(UBound(aDocTmp)).sSit = aDoc(nLinha).sSit
                aDocTmp(UBound(aDocTmp)).nValorPrincipal = aDoc(nLinha).nValorPrincipal
                aDocTmp(UBound(aDocTmp)).nValorMulta = aDoc(nLinha).nValorMulta
                aDocTmp(UBound(aDocTmp)).nValorJuros = aDoc(nLinha).nValorJuros
                aDocTmp(UBound(aDocTmp)).nValorCorrecao = aDoc(nLinha).nValorCorrecao
                aDocTmp(UBound(aDocTmp)).nValorTotal = aDoc(nLinha).nValorTotal
                aDocTmp(UBound(aDocTmp)).nValorTarifa = aDoc(nLinha).nValorTarifa
                aDocTmp(UBound(aDocTmp)).nValorDif = aDoc(nLinha).nValorDif
                aDocTmp(UBound(aDocTmp)).nValorCompensado = aDoc(nLinha).nValorCompensado
                aDocTmp(UBound(aDocTmp)).sBx = aDoc(nLinha).sBx
                aDocTmp(UBound(aDocTmp)).sDp = aDoc(nLinha).sDp
                aDocTmp(UBound(aDocTmp)).nSeqReg = aDoc(nLinha).nSeqReg
                aDocTmp(UBound(aDocTmp)).bExiste = aDoc(nLinha).bExiste
            End If
        Next
        ReDim aDoc(0)
        For nLinha = 1 To UBound(aDocTmp)
                ReDim Preserve aDoc(UBound(aDoc) + 1)
                aDoc(UBound(aDoc)).nNumDoc = aDocTmp(nLinha).nNumDoc
                aDoc(UBound(aDoc)).nCodReduz = aDocTmp(nLinha).nCodReduz
                aDoc(UBound(aDoc)).nAno = aDocTmp(nLinha).nAno
                aDoc(UBound(aDoc)).nLanc = aDocTmp(nLinha).nLanc
                aDoc(UBound(aDoc)).nSeq = aDocTmp(nLinha).nSeq
                aDoc(UBound(aDoc)).nParc = aDocTmp(nLinha).nParc
                aDoc(UBound(aDoc)).nCompl = aDocTmp(nLinha).nCompl
                aDoc(UBound(aDoc)).sDataVencto = aDocTmp(nLinha).sDataVencto
                aDoc(UBound(aDoc)).sSit = aDocTmp(nLinha).sSit
                aDoc(UBound(aDoc)).nValorPrincipal = aDocTmp(nLinha).nValorPrincipal
                aDoc(UBound(aDoc)).nValorMulta = aDocTmp(nLinha).nValorMulta
                aDoc(UBound(aDoc)).nValorJuros = aDocTmp(nLinha).nValorJuros
                aDoc(UBound(aDoc)).nValorCorrecao = aDocTmp(nLinha).nValorCorrecao
                aDoc(UBound(aDoc)).nValorTotal = aDocTmp(nLinha).nValorTotal
                aDoc(UBound(aDoc)).nValorTarifa = aDocTmp(nLinha).nValorTarifa
                aDoc(UBound(aDoc)).nValorDif = aDocTmp(nLinha).nValorDif
                aDoc(UBound(aDoc)).nValorCompensado = aDocTmp(nLinha).nValorCompensado
                aDoc(UBound(aDoc)).sBx = aDocTmp(nLinha).sBx
                aDoc(UBound(aDoc)).sDp = aDocTmp(nLinha).sDp
                aDoc(UBound(aDoc)).nSeqReg = aDocTmp(nLinha).nSeqReg
                aDoc(UBound(aDoc)).bExiste = aDocTmp(nLinha).bExiste
        Next
    
        grdParc.Rows = 1
        If grdReg.Rows > 2 Then
            grdReg.RemoveItem (grdReg.Row)
            grdReg.Row = 1
            CarregaDocumento grdReg.TextMatrix(1, 0), grdReg.TextMatrix(1, 11), grdReg.TextMatrix(1, 9)
        Else
            grdReg.Rows = 1
        End If
    End If
End If

mskDataCred.Text = Format(frmPagAutomatico.Mv.value, "dd/mm/yyyy")
End Sub

Private Sub cmdSair_Click()
pnlCampo.Visible = True
pnlManual.Visible = False
pnlArquivo.Visible = True
pnlDetalhe.Visible = True
lblTit.Visible = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set m_cMenu = Nothing
End Sub

Private Sub grdReg_Click()
If grdReg.Rows = 1 Then Exit Sub
If grdReg.MouseRow > 0 Then Exit Sub

grdReg.Row = 1
grdReg.RowSel = grdReg.Rows - 1
grdReg.col = grdReg.MouseCol
grdReg.Sort = flexSortNumericAscending

End Sub

Private Sub grdReg_RowColChange()
On Error Resume Next
If grdReg.Row > 0 Then CarregaDocumento grdReg.TextMatrix(grdReg.Row, 0), grdReg.TextMatrix(grdReg.Row, 11), grdReg.TextMatrix(grdReg.Row, 9)
End Sub

Private Sub CarregaDocumento(nNumDoc As Long, nSeqDoc As Integer, sCNPJ As String)
Dim nLinha As Integer
grdParc.Rows = 1

For nLinha = 1 To UBound(aDoc)
    If lblAS.Caption = "N" Then
        With aDoc(nLinha)
            If .nNumDoc = nNumDoc And .nSeqDoc = nSeqDoc And .bExiste Then
                grdParc.AddItem .nCodReduz & Chr(9) & .nAno & Chr(9) & .nLanc & Chr(9) & .nSeq & Chr(9) & .nParc & Chr(9) & .nCompl & Chr(9) & .sDataVencto & Chr(9) & _
                Format(.sSit, "00") & Chr(9) & FormatNumber(.nValorPrincipal, 2) & Chr(9) & FormatNumber(.nValorMulta, 2) & Chr(9) & FormatNumber(.nValorJuros, 2) & Chr(9) & _
                FormatNumber(.nValorCorrecao, 2) & Chr(9) & FormatNumber(.nValorTotal, 2) & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & FormatNumber(.nValorDif, 2) & Chr(9) & _
                FormatNumber(.nValorCompensado, 2) & Chr(9) & .sDp
            End If
        End With
    Else
        With aDoc(nLinha)
            If .sCNPJ = RetornaNumero(sCNPJ) And .nSeqReg = nSeqDoc And .bExiste Then
                grdParc.AddItem .nCodReduz & Chr(9) & .nAno & Chr(9) & .nLanc & Chr(9) & .nSeq & Chr(9) & .nParc & Chr(9) & .nCompl & Chr(9) & .sDataVencto & Chr(9) & _
                Format(.sSit, "00") & Chr(9) & FormatNumber(.nValorPrincipal, 2) & Chr(9) & FormatNumber(.nValorMulta, 2) & Chr(9) & FormatNumber(.nValorJuros, 2) & Chr(9) & _
                FormatNumber(.nValorCorrecao, 2) & Chr(9) & FormatNumber(.nValorTotal, 2) & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & FormatNumber(.nValorDif, 2) & Chr(9) & _
                FormatNumber(.nValorCompensado, 2) & Chr(9) & .sDp
            End If
        End With
    End If
Next

End Sub


Private Sub Form_Load()
Dim x As Long
x = SetParent(Me.HWND, frmMdi.HWND)
MontaMenu
Centraliza Me
LimpaTela
Select Case Val(frmPagAutomatico.lblAux.Caption)
    Case 0
        lblBanco.Caption = "033-SANTANDER BANESPA"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(0).PictureNormal
    Case 1
        lblBanco.Caption = "001-BANCO DO BRASIL"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(1).PictureNormal
    Case 2
        lblBanco.Caption = "000-OUTROS BANCOS"
'        imgLogo.Picture = Null
    Case 3
        lblBanco.Caption = "237-BANCO BRADESCO"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(3).PictureNormal
    Case 4
        lblBanco.Caption = "104-CAIXA FEDERAL"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(4).PictureNormal
    Case 5
        lblBanco.Caption = "399-HSBC AMRO BANK"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(5).PictureNormal
    Case 6
        lblBanco.Caption = "341-BANCO ITAU"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(6).PictureNormal
    Case 7
        lblBanco.Caption = "151-NOSSA CAIXA"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(7).PictureNormal
    Case 8
        lblBanco.Caption = "409-UNIBANCO S/A"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(8).PictureNormal
End Select
sBanco = lblBanco.Caption
CarregaTributo
End Sub


Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Private Sub CarregaParcela(nNumDoc As Long, nNumSeqReg As Integer, sDataCredito As String)
Dim qd As New rdoQuery, RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, x As Integer, nFicha As Long, nFichaJM As Long, nFichaC As Long
Dim nLinha As Integer, nValorPago As Double, sDataPag As String, sDup As String, sBax As String, bNewDoc As Boolean, nSoma As Double
Dim nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double, nValorDif As Double
Dim nValorComp As Double, nValorTarifa As Double, nValorChecar As Double, bIsento As Boolean, nPercDesconto As Double, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset
Dim nValorTarifaGlobal As Double, nLast As Integer, nQtdeLanc As Integer, aDocTmp() As Documento, nSeqReg As Integer, sNumProc As String, dDataVencto As Date
Dim nCodReduz As Long, RdoEicon As rdoResultset

ReDim aDocTmp(0): bNewDoc = True: nSomaTrib = 0: ReDim aTrib(0)
grdParc.Rows = 1
If lblAS.Caption = "S" Then Exit Sub
Set qd.ActiveConnection = cn

For nLinha = 0 To UBound(aRegistro)
    If aRegistro(nLinha).nNumDoc = nNumDoc And aRegistro(nLinha).nSeq = nNumSeqReg Then
        nValorPago = aRegistro(nLinha).nValorPago
        sDataPag = aRegistro(nLinha).sDataPag
        nValorTarifaGlobal = 0
        bIsento = aRegistro(nLinha).bIsentoMJ
        Exit For
    End If
Next

'If nNumDoc = 15200191 Then MsgBox "teste"

nQtdeLanc = 0: nSeqReg = 0
Sql = "SELECT parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, debitoparcela.numerolivro, debitoparcela.paginalivro, debitoparcela.dataajuiza, "
Sql = Sql & "parceladocumento.seqlancamento, parceladocumento.numparcela, parceladocumento.codcomplemento,"
Sql = Sql & "parceladocumento.numdocumento,parceladocumento.plano, debitoparcela.datavencimento, debitoparcela.statuslanc, numdocumento.datadocumento,"
Sql = Sql & "NumDocumento.valortaxadoc,numdocumento.percisencao FROM parceladocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND "
Sql = Sql & "parceladocumento.anoexercicio = debitoparcela.anoexercicio AND parceladocumento.codlancamento = debitoparcela.codlancamento AND "
Sql = Sql & "parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND "
Sql = Sql & "parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
Sql = Sql & "WHERE PARCELADOCUMENTO.NumDocumento = " & nNumDoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        ReDim Preserve aDoc(UBound(aDoc) + 1)
        nLast = UBound(aDoc)
        aDoc(nLast).nNumDoc = nNumDoc
        aDoc(nLast).bExiste = False
        For nLinha = 0 To UBound(aRegistro)
            If aRegistro(nLinha).nNumDoc = nNumDoc And aRegistro(nLinha).nSeq = nNumSeqReg Then
                aRegistro(nLinha).sSitRetorno = "02-DOCUMENTO SEM LANÇAMENTOS"
                Exit For
            End If
        Next
        nPercDesconto = 0
        .Close: Exit Sub
    Else
        If Val(SubNull(!plano)) > 0 Then
            If !plano = 16 Then
                nPercDesconto = 100
            ElseIf !plano = 17 Then
                nPercDesconto = 80
            ElseIf !plano = 18 Then
                nPercDesconto = 60
            ElseIf !plano = 19 Then
                nPercDesconto = 50
            ElseIf !plano = 20 Then
                nPercDesconto = 50
            ElseIf !plano = 21 Then
                nPercDesconto = 40
            ElseIf !plano = 22 Then
                nPercDesconto = 30
            ElseIf !plano = 23 Then
                nPercDesconto = 100
            ElseIf !plano = 24 Then
                nPercDesconto = 60
            ElseIf !plano = 25 Then
                nPercDesconto = 40
            ElseIf !plano = 26 Then
                nPercDesconto = 100
            ElseIf !plano = 27 Then
                nPercDesconto = 90
            ElseIf !plano = 28 Then
                nPercDesconto = 80
            ElseIf !plano = 29 Then
                nPercDesconto = 50
            ElseIf !plano = 30 Then
                nPercDesconto = 40
            ElseIf !plano = 31 Then
                nPercDesconto = 30
            ElseIf !plano = 32 Then
                nPercDesconto = 100
            Else
                nPercDesconto = 0
            End If
        Else
            nPercDesconto = 0
        End If
        If nPercDesconto > 0 Then
            bIsento = True
        End If
            'sDataPag = Format(!DATADOCUMENTO, "dd/mm/yyyy")
        Do Until .EOF
            nCodReduz = !CODREDUZIDO
            If nCodReduz > 1000000 Then
                ConectaEicon2
                Sql = "select * from tb_inter_empresas_giss where num_cadastro=" & nCodReduz
                Set RdoEicon = cnEicon2.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If Not IsNull(RdoEicon!Inscricao_gti) Then
                    nCodReduz = RdoEicon!Inscricao_gti
                Else
                    MsgBox "teste"
                End If
                cnEicon2.Close
            End If
            'CARREGA AS PARCELAS DO DOCUMENTO
            On Error Resume Next
            RdoAux2.Close
            On Error GoTo 0
            nValorPrincipal = 0: nValorMulta = 0: nValorJuros = 0: nValorCorrecao = 0: nValorTotal = 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz: qd(1) = nCodReduz
            qd(2) = !AnoExercicio: qd(3) = !AnoExercicio
            qd(4) = !CodLancamento: qd(5) = !CodLancamento
            qd(6) = !SeqLancamento: qd(7) = !SeqLancamento
            qd(8) = !NumParcela: qd(9) = !NumParcela
            qd(10) = !CODCOMPLEMENTO: qd(11) = !CODCOMPLEMENTO
            qd(12) = 1: qd(13) = 99: qd(14) = Format(sDataPag, "mm/dd/yyyy")
            qd(15) = NomeDoUsuario
            Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
            Do Until RdoAux2.EOF
                dDataVencto = RdoAux2!DataVencimentoCalc
                nValorPrincipal = nValorPrincipal + RdoAux2!ValorTributo
                If Not bIsento Then
                    nValorMulta = RdoAux2!ValorMulta
                    nValorJuros = RdoAux2!ValorJuros
                Else
                    nValorMulta = 0
                    nValorJuros = 0
                End If
                nValorCorrecao = nValorCorrecao + CDbl(SubNull(RdoAux2!ValorCorrecao))
                
                If Not bIsento Then
                    nValorTotal = nValorTotal + (RdoAux2!ValorTributo + RdoAux2!ValorMulta + RdoAux2!ValorJuros + CDbl(SubNull(RdoAux2!ValorCorrecao)))
                Else
                    If nPercDesconto > 0 Then
                        nValorMulta = nValorMulta + ((100 - nPercDesconto) * RdoAux2!ValorMulta / 100)
                        nValorJuros = nValorJuros + ((100 - nPercDesconto) * RdoAux2!ValorJuros / 100)
                        nValorTotal = nValorTotal + (RdoAux2!ValorTributo + CDbl(SubNull(RdoAux2!ValorCorrecao) + nValorMulta + nValorJuros))
                   Else
                        nValorTotal = nValorTotal + (RdoAux2!ValorTributo + CDbl(SubNull(RdoAux2!ValorCorrecao)))
                   End If
                   
                End If
                'Carrega os tributos
                ReDim Preserve aTrib(UBound(aTrib) + 1)
                nLast = UBound(aTrib)
                aTrib(nLast).nNumDoc = nNumDoc
                aTrib(nLast).nSeqDoc = aRegistro(nLinha).nSeq
                aTrib(nLast).nCodReduz = nCodReduz
                aTrib(nLast).nAno = !AnoExercicio
                aTrib(nLast).nLanc = !CodLancamento
                aTrib(nLast).nSeq = !SeqLancamento
                aTrib(nLast).nParc = !NumParcela
                aTrib(nLast).nCompl = !CODCOMPLEMENTO
                aTrib(nLast).nCodTrib = RdoAux2!CodTributo
                aTrib(nLast).nValorPrincipal = RdoAux2!ValorTributo
                If Not bIsento Then
                    aTrib(nLast).nValorMulta = RdoAux2!ValorMulta
                    aTrib(nLast).nValorJuros = RdoAux2!ValorJuros
                Else
                    aTrib(nLast).nValorMulta = 0
                    aTrib(nLast).nValorJuros = 0
                End If
                aTrib(nLast).nValorCorrecao = RdoAux2!ValorCorrecao
                aTrib(nLast).nValorTarifa = 0
                bNewDoc = False
                aTrib(nLast).nValorTotal = aTrib(nLast).nValorPrincipal + aTrib(nLast).nValorJuros + aTrib(nLast).nValorMulta + aTrib(nLast).nValorCorrecao + aTrib(nLast).nValorTarifa
                aTrib(nLast).nValorCompensado = aTrib(nLast).nValorTotal
                
                '**************************
                If RdoAux2!CodLancamento = 20 Then
                    Sql = "SELECT NUMPROCESSO FROM DESTINOREPARC WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & RdoAux2!AnoExercicio & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & RdoAux2!CodLancamento & " AND NUMSEQUENCIA=" & RdoAux2!SeqLancamento & " AND NUMPARCELA=" & RdoAux2!NumParcela & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & RdoAux2!CODCOMPLEMENTO
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux3
                        If .RowCount = 0 Then
                            aTrib(nLast).sAj = "S"
                            aTrib(nLast).sDA = "S"
                            GoTo continua
                        End If
                        sNumProc = !numprocesso
                       .Close
                    End With
                    
                    Sql = "SELECT * FROM ORIGEMREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
                    Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux4
                        If RdoAux4.RowCount = 0 Then GoTo continua
                        Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                        Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !numsequencia & " AND NUMPARCELA=" & !NumParcela & " AND "
                        Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                        With RdoAux3
                            If Not IsNull(!datainscricao) Then
                               aTrib(nLast).sDA = "S"
                            Else
                               aTrib(nLast).sDA = "N"
                            End If
                            If Not IsNull(!dataajuiza) Then
                               aTrib(nLast).sAj = "S"
                            Else
                               aTrib(nLast).sAj = "N"
                            End If
                            .Close
                        End With
                        .Close
                    End With
                Else
                    aTrib(nLast).sAj = IIf(IsNull(RdoAux2!dataajuiza), "N", "S")
                    aTrib(nLast).sDA = IIf(IsNull(RdoAux2!datainscricao), "N", "S")
                End If
continua:
                '**************************
                
                For x = 1 To UBound(aTribF)
                    If aTribF(x).nCodTrib = aTrib(nLast).nCodTrib Then
                        Exit For
                    End If
                Next
                
                If aTrib(nLast).sAj = "N" Then
                    If aTrib(nLast).sDA = "N" Then
                        nFicha = aTribF(x).Ficha
                        nFichaJM = aTribF(x).FichaJrMulta
                        nFichaC = 0
                    Else
                        nFicha = aTribF(x).FichaDivida
                        nFichaJM = aTribF(x).FichaDaJrMul
                        nFichaC = aTribF(x).FichaDaEnca
                    End If
                Else
                    nFicha = aTribF(x).FichaAjuiza
                    nFichaJM = aTribF(x).FichaAjJrMul
                    nFichaC = aTribF(x).FichaAjEnca
                End If
                
                aTrib(nLast).nFicha = nFicha
                aTrib(nLast).nFichaJM = nFichaJM
                aTrib(nLast).nFichaC = nFichaC
                
                nSoma = nSoma + aTrib(nLast).nValorTotal
                
                With aTrib(nLast)
                    grdTrib.AddItem .nNumDoc & Chr(9) & .nSeqDoc & Chr(9) & .nCodReduz & Chr(9) & .nAno & Chr(9) & .nLanc & Chr(9) & .nSeq & Chr(9) & .nParc & Chr(9) & .nCompl & Chr(9) & .nCodTrib & Chr(9) & .nValorPrincipal & _
                    Chr(9) & .nValorMulta & Chr(9) & .nValorJuros & Chr(9) & .nValorCorrecao & Chr(9) & .nValorTarifa & Chr(9) & .nValorTotal & Chr(9) & .nValorCompensado & Chr(9) & .sAj & Chr(9) & .sDA & Chr(9) & nFicha & Chr(9) & _
                    nFichaJM & Chr(9) & nFichaC
                    
                    'grdTrib.AddItem .nNumDoc & Chr(9) & .nSeqDoc & Chr(9) & .nCodReduz & Chr(9) & .nAno & Chr(9) & .nLanc & Chr(9) & .nSeq & Chr(9) & .nParc & Chr(9) & .nCompl & Chr(9) & .nCodTrib & Chr(9) & Round(.nValorPrincipal, 2) & _
                    Chr(9) & Round(.nValorMulta, 2) & Chr(9) & Round(.nValorJuros, 2) & Chr(9) & Round(.nValorCorrecao, 2) & Chr(9) & Round(.nValorTarifa, 2) & Chr(9) & Round(.nValorTotal, 2) & Chr(9) & Round(.nValorCompensado, 2) & Chr(9) & .sAj & Chr(9) & .sDA & Chr(9) & nFicha & Chr(9) & _
                    nFichaJM & Chr(9) & nFichaC
                End With
                
                RdoAux2.MoveNext
            Loop
           
            nQtdeLanc = nQtdeLanc + 1
            ReDim Preserve aDoc(UBound(aDoc) + 1)
            nLast = UBound(aDoc)
            aDoc(nLast).nNumDoc = nNumDoc
            aDoc(nLast).nSeqDoc = aRegistro(nLinha).nSeq
            aDoc(nLast).nCodReduz = nCodReduz
            aDoc(nLast).nAno = !AnoExercicio
            aDoc(nLast).nLanc = !CodLancamento
            aDoc(nLast).nSeq = !SeqLancamento
            aDoc(nLast).nParc = !NumParcela
            aDoc(nLast).nCompl = !CODCOMPLEMENTO
            'aDoc(nLast).sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
            aDoc(nLast).sDataVencto = Format(dDataVencto, "dd/mm/yyyy")
            aDoc(nLast).sSit = !statuslanc
            aDoc(nLast).nNumeroLivro = Val(SubNull(!numerolivro))
            aDoc(nLast).nPaginaLivro = Val(SubNull(!paginalivro))
            aDoc(nLast).bAjuizado = IIf(IsNull(!dataajuiza), False, True)
            aDoc(nLast).nValorPrincipal = nValorPrincipal
            aDoc(nLast).nValorMulta = nValorMulta
            aDoc(nLast).nValorJuros = nValorJuros
            aDoc(nLast).nValorCorrecao = nValorCorrecao
            aDoc(nLast).nValorTotal = nValorTotal
            aDoc(nLast).nValorTarifa = 0
            aDoc(nLast).nValorDif = 0
            aDoc(nLast).nValorCompensado = 0
            aDoc(nLast).sBx = "X"
            aDoc(nLast).sDp = "X"
            aDoc(nLast).nSeqReg = nSeqReg
            aDoc(nLast).bExiste = True
            RdoAux2.Close
            nSeqReg = nSeqReg + 1
'            If Val(SubNull(!paginalivro)) > 0 Then MsgBox "tesdte"
           .MoveNext
        Loop
    End If
    .Close
End With

'DIVIDE A TAXA

nValorTarifa = FormatNumber(nValorTarifaGlobal / nQtdeLanc, 2)
If nValorTarifa > 4 Then
    nValorTarifa = 0
    nValorTarifaGlobal = 0
End If

For nLinha = 1 To UBound(aDoc)
    With aDoc(nLinha)
        If .nNumDoc = nNumDoc And .nSeqDoc = nNumSeqReg Then
           .nValorTarifa = nValorTarifa
            If aDocTmp(0).nNumDoc = 0 Or .nSeqReg > aDocTmp(0).nSeqReg Then
                aDocTmp(0).nNumDoc = .nNumDoc: aDocTmp(0).nSeqDoc = .nSeqDoc: aDocTmp(0).nCodReduz = .nCodReduz: aDocTmp(0).nAno = .nAno: aDocTmp(0).nLanc = .nLanc
                aDocTmp(0).nSeq = .nSeq: aDocTmp(0).nParc = .nParc: aDocTmp(0).nCompl = .nCompl: aDocTmp(0).nSeqReg = .nSeqReg
            End If
        End If
    End With
Next

'CARREGA MATRIZ TEMP COM ULTIMO LANCAMENTO DO DOCUMENTO
'REMOVE DIFERENÇA DA ÚLTIMA TAXA
If FormatNumber(nValorTarifa * nQtdeLanc, 2) > FormatNumber(nValorTarifaGlobal, 2) Then
    For nLinha = 1 To UBound(aDoc)
        With aDoc(nLinha)
            If .nNumDoc = aDocTmp(0).nNumDoc And .nSeqDoc = aDocTmp(0).nSeqDoc And .nCodReduz = aDocTmp(0).nCodReduz And .nAno = aDocTmp(0).nAno And .nLanc = aDocTmp(0).nLanc And _
               .nSeq = aDocTmp(0).nSeq And .nParc = aDocTmp(0).nParc And .nCompl = aDocTmp(0).nCompl And .nSeqReg = aDocTmp(0).nSeqReg Then
               aDoc(nLinha).nValorTarifa = nValorTarifa - ((nValorTarifa * nQtdeLanc) - nValorTarifaGlobal)
               'aDoc(nLinha).nValorTarifa = FormatNumber(nValorTarifa - ((nValorTarifa * nQtdeLanc) - nValorTarifaGlobal), 2)
               Exit For
            End If
        End With
    Next
ElseIf FormatNumber(nValorTarifa * nQtdeLanc, 2) < FormatNumber(nValorTarifaGlobal, 2) Then
    For nLinha = 1 To UBound(aDoc)
        With aDoc(nLinha)
            If .nNumDoc = aDocTmp(0).nNumDoc And .nSeqDoc = aDocTmp(0).nSeqDoc And .nCodReduz = aDocTmp(0).nCodReduz And .nAno = aDocTmp(0).nAno And .nLanc = aDocTmp(0).nLanc And _
               .nSeq = aDocTmp(0).nSeq And .nParc = aDocTmp(0).nParc And .nCompl = aDocTmp(0).nCompl And .nSeqReg = aDocTmp(0).nSeqReg Then
               'aDoc(nLinha).nValorTarifa = FormatNumber(nValorTarifa + (nValorTarifaGlobal - (nValorTarifa * nQtdeLanc)), 2)
               aDoc(nLinha).nValorTarifa = nValorTarifa + (nValorTarifaGlobal - (nValorTarifa * nQtdeLanc))
               Exit For
            End If
        End With
    Next
End If

'CALCULA AS DIFERENÇAS
nValorDif = 0
For nLinha = 1 To UBound(aDoc)
    With aDoc(nLinha)
        If .nNumDoc = nNumDoc And .nSeqDoc = nNumSeqReg Then
            'nValorDif = nValorDif + Round(.nValorTotal, 2) + Round(.nValorTarifa, 2)
            nValorDif = nValorDif + .nValorTotal
        End If
    End With
Next
For nLinha = 1 To UBound(aDoc)
    With aDoc(nLinha)
        If .nNumDoc = aDocTmp(0).nNumDoc And .nSeqDoc = aDocTmp(0).nSeqDoc And .nCodReduz = aDocTmp(0).nCodReduz And .nAno = aDocTmp(0).nAno And .nLanc = aDocTmp(0).nLanc And _
           .nSeq = aDocTmp(0).nSeq And .nParc = aDocTmp(0).nParc And .nCompl = aDocTmp(0).nCompl And .nSeqReg = aDocTmp(0).nSeqReg Then
           aDoc(nLinha).nValorDif = nValorPago - nValorDif
           'aDoc(nLinha).nValorDif = FormatNumber(nValorPago - nValorDif, 2)
           nValorDif = nValorPago - nValorDif
        End If
    End With
Next

'VALORCOMPENSADO
For nLinha = 1 To UBound(aDoc)
    With aDoc(nLinha)
        If .nNumDoc = nNumDoc And .nSeqDoc = nNumSeqReg Then
            .nValorCompensado = .nValorTotal + .nValorTarifa + .nValorDif
            '.nValorCompensado = FormatNumber(.nValorTotal + .nValorTarifa + .nValorDif, 2)
'            .nValorCompensado = FormatNumber(.nValorTotal + .nValorTarifa, 2)
        End If
    End With
Next


'DUPLICIDADE
For nLinha = 1 To UBound(aDoc)
    With aDoc(nLinha)
        If .nNumDoc = nNumDoc And .nSeqDoc = nNumSeqReg Then
            sDup = "N"
            Sql = "SELECT CODREDUZIDO,DATARECEBIMENTO,NUMDOCUMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & .nCodReduz & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND SEQLANCAMENTO=" & .nSeq & " AND "
            Sql = Sql & "NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl & " AND RESTITUIDO IS NULL"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                   If Format(!datarecebimento, "dd/mm/yyyy") <> Format(CDate(sDataCredito), "dd/MM/yyyy") Or Val(SubNull(!NumDocumento)) <> nNumDoc Then
                       sDup = "S"
                   End If
                Else
                   sDup = "N"
                End If
               .Close
            End With
            .sDp = sDup
        End If
    End With
Next

'DIFERENÇAS DE PAGAMENTO NOS TRIBUTOS
'nValorDif = Round(nValorPago - Round(nSoma, 2), 2)
nValorDif = nValorPago - nSoma
If nValorDif = 0 Then Exit Sub

If nValorDif < 0 Then 'O CONTRIBUINTE PAGOU MENOS DO QUE DEVIA
    nValorDif = Abs(nValorDif)
    For nLinha = 1 To UBound(aTrib)
        With aTrib(nLinha)
            If .nNumDoc = nNumDoc And .nSeqDoc = nNumSeqReg Then
                If nValorDif <= .nValorPrincipal Then  'SE DER PARA DESCONTAR TODA A DIF NESTA PARCELA, DESCONTA E SAI FORA
                'If nValorDif <= Round(.nValorPrincipal, 2) Then  'SE DER PARA DESCONTAR TODA A DIF NESTA PARCELA, DESCONTA E SAI FORA
                    aTrib(nLinha).nValorCompensado = aTrib(nLinha).nValorCompensado - nValorDif
'                    aTrib(nLinha).nValorPrincipal = aTrib(nLinha).nValorPrincipal - nValorDif
'                    aTrib(nLinha).nValorTotal = aTrib(nLinha).nValorTotal - nValorDif
                   'ATUALIZA O GRID DE TRIBUTOS
                    For x = 1 To grdTrib.Rows - 1
                        If grdTrib.TextMatrix(x, 0) = .nNumDoc And grdTrib.TextMatrix(x, 1) = .nSeqDoc And grdTrib.TextMatrix(x, 2) = .nCodReduz And _
                            grdTrib.TextMatrix(x, 3) = .nAno And grdTrib.TextMatrix(x, 4) = .nLanc And grdTrib.TextMatrix(x, 5) = .nSeq And _
                            grdTrib.TextMatrix(x, 6) = .nParc And grdTrib.TextMatrix(x, 7) = .nCompl And grdTrib.TextMatrix(x, 8) = .nCodTrib Then
                            grdTrib.TextMatrix(x, 9) = Format(.nValorPrincipal, "#0.00")
                            grdTrib.TextMatrix(x, 14) = Format(.nValorTotal, "#0.00")
                            Exit For
                        End If
                    Next
                    Exit For
                Else
'                    MsgBox "falta implementar 1"
                End If
            End If
        End With
    Next
ElseIf nValorDif > 0 Then 'O VALOR PAGO FOI MAIOR DO QUE DEVIA
    nValorDif = Abs(nValorDif)
    For nLinha = 1 To UBound(aTrib)
        With aTrib(nLinha)
            If .nNumDoc = nNumDoc And .nSeqDoc = nNumSeqReg Then
                If nValorDif <= Round(.nValorPrincipal, 2) Then  'SE DER PARA DESCONTAR TODA A DIF NESTA PARCELA, DESCONTA E SAI FORA
                    aTrib(nLinha).nValorCompensado = aTrib(nLinha).nValorCompensado + nValorDif
 '                   aTrib(nLinha).nValorPrincipal = aTrib(nLinha).nValorPrincipal + nValorDif
 '                   aTrib(nLinha).nValorTotal = aTrib(nLinha).nValorTotal + nValorDif
                   'ATUALIZA O GRID DE TRIBUTOS
                    For x = 1 To grdTrib.Rows - 1
                        If grdTrib.TextMatrix(x, 0) = .nNumDoc And grdTrib.TextMatrix(x, 1) = .nSeqDoc And grdTrib.TextMatrix(x, 2) = .nCodReduz And _
                            grdTrib.TextMatrix(x, 3) = .nAno And grdTrib.TextMatrix(x, 4) = .nLanc And grdTrib.TextMatrix(x, 5) = .nSeq And _
                            grdTrib.TextMatrix(x, 6) = .nParc And grdTrib.TextMatrix(x, 7) = .nCompl And grdTrib.TextMatrix(x, 8) = .nCodTrib Then
                            grdTrib.TextMatrix(x, 9) = Format(.nValorPrincipal, "#0.00")
                            grdTrib.TextMatrix(x, 14) = Format(.nValorTotal, "#0.00")
                            Exit For
                        End If
                    Next
                    Exit For
                Else
'                    MsgBox "falta implementar 1"
                End If
            End If
        End With
    Next
End If
'

DoEvents

End Sub

Private Sub LimpaTela()
    PBar.value = 0
    lblDC.Caption = ""
    lblDB.Caption = ""
    lblAS.Caption = "N"
    lblAC.Caption = "N"
    lblDA.Caption = "N"
    lblNumReg.Caption = "0,00"
    lblValorTotal.Caption = "0,00"
    lblValorEfetivo.Caption = "0,00"
    grdReg.Rows = 1
    grdParc.Rows = 1
    grdTrib.Rows = 1
    lblErro.Visible = False: cmdErro.Visible = False: txtErro.Visible = False
End Sub

Private Sub lstArq_Click()
LimpaTela
Select Case Val(frmPagAutomatico.lblAux.Caption)
    Case 0
        lblBanco.Caption = "033-SANTANDER BANESPA"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(0).PictureNormal
    Case 1
        lblBanco.Caption = "001-BANCO DO BRASIL"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(1).PictureNormal
    Case 2
       ' lblBanco.Caption = "641-BBV BANCO"
        lblBanco.Caption = "000-OUTROS BANCOS"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(2).PictureNormal
    Case 3
        lblBanco.Caption = "237-BANCO BRADESCO"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(3).PictureNormal
    Case 4
        lblBanco.Caption = "104-CAIXA FEDERAL"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(4).PictureNormal
    Case 5
        lblBanco.Caption = "399-HSBC AMRO BANK"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(5).PictureNormal
    Case 6
        lblBanco.Caption = "341-BANCO ITAU"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(6).PictureNormal
    Case 7
        lblBanco.Caption = "151-NOSSA CAIXA"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(7).PictureNormal
    Case 8
        lblBanco.Caption = "409-UNIBANCO S/A"
        imgLogo.Picture = frmPagAutomatico.cmdBanco(8).PictureNormal
End Select
sBanco = lblBanco.Caption
End Sub

Private Sub m_cMenu_Click(ItemNumber As Long)
Dim z As Variant, Sql As String, RdoAux As rdoResultset, bAchou As Boolean
Select Case m_cMenu.ItemKey(ItemNumber)
    Case "mnuVisualizar"
        If lstArq.ListCount > 0 Then
           x = Shell("NOTEPAD" & " " & txtPath.Text & lstArq.Text, vbNormalFocus)
        End If
    Case "mnuBaixa"
        If lblErro.Visible And lblAS.Caption = "N" Then
            MsgBox "Não é possivel efetuar baixa enquanto houver documentos com erro.", vbCritical, "Atenção"
            Exit Sub
        End If

        EfetuaBaixa
    Case "mnuReativar"
        ReativarArquivo
    Case "mnuFixDoc"
        bAchou = False
        z = InputBox("Digite o n° do documento sem o digito.", "Informação requerida")
        If Val(z) > 0 Then
            Sql = "SELECT * FROM PARCELADOCUMENTO WHERE NUMDOCUMENTO=" & Val(z)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                bAchou = True
            End If
            RdoAux.Close
            If Not bAchou Then
                MsgBox "Documento não cadastrado.", vbCritical, "Atenção"
                Exit Sub
            Else
                bAchou = False
                Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    bAchou = True
                End If
                RdoAux.Close
            End If
            If bAchou Then
                MsgBox "Documento já cadastrado.", vbExclamation, "Atenção"
            Else
                Sql = "INSERT NUMDOCUMENTO(NUMDOCUMENTO,DATADOCUMENTO,emissor) VALUES(" & Val(z) & ",'" & Format(Now, "mm/dd/yyyy") & "','" & NomeDeLogin & " (BAIXA BANCARIA)" & "')"
                cn.Execute Sql, rdExecDirect
                MsgBox "Documento cadastrado com sucesso.", vbExclamation, "Atenção"
            End If
        End If
    Case "mnuBaixaManual"
        pnlCampo.Visible = False
        pnlManual.Visible = True
        pnlArquivo.Visible = False
        pnlDetalhe.Visible = False
        lblTit.Visible = True
        txtBanco.Text = Left$(lblBanco.Caption, 3)
        LimpaTela
        ReDim aRegistro(0): ReDim aDoc(0)
        txtNumDoc.SetFocus
        mskDataCred.Text = Format(frmPagAutomatico.Mv.value, "dd/mm/yyyy")
    Case "mnuCBR724"
        
        With CommonDialog1
            .DialogTitle = "Selecione um arquivo CBR724"
            .CancelError = True
            .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
            .InitDir = "\\192.168.200.130\atualizagti"
            .Filter = "All Files (*.*)|*.*"
            .ShowOpen
            
            txtPath.Text = pathOfFile(.FileName)
            lstArq.AddItem .FileTitle
            lstArq.ListIndex = lstArq.ListCount - 1
            lstArq_Click
            
        End With
    Case "mnuOutro"
        MsgBox "Nº sequencial: " & sSeqArq, vbInformation, "OUTRAS INFORMAÇÕES"
    Case "mnuResumo"
        ResumoArquivo
    Case "mnuAnaliseR"
        If grdReg.Rows = 1 Then
            MsgBox "Arquivo não carregado.", vbExclamation, "Atenção"
        Else
            frmReport.ShowReport "ANALISE2", frmMdi.HWND, Me.HWND
        End If
    Case "mnuAnaliseD"
        If grdReg.Rows = 1 Then
            MsgBox "Arquivo não carregado.", vbExclamation, "Atenção"
        Else
            frmReport.ShowReport "ANALISE1", frmMdi.HWND, Me.HWND
        End If
End Select

End Sub

Private Sub EfetuaBaixa()
Dim nLinha As Integer, nLinha2 As Integer, Sql As String, aDocTmp() As Documento, nStatus As Integer, nSeqPag As Integer, sCNPJ As String, nNumDoc As Long
Dim RdoAux As rdoResultset, sNomeArq As String, nComplCP As Integer, RdoAux2 As rdoResultset, aTribProp() As TributoProp, k As Integer, nSomaTrib As Double, nSeq2 As Integer
Dim sNumProc As String, nNumproc As Long, nAnoproc As Integer, nSeq As Integer

If grdReg.Rows = 1 Then
    MsgBox "Arquivo não carregado.", vbExclamation, "Atenção"
    Exit Sub
End If

If lstArq.Visible = True Then
    If lblDB.Caption <> "Sem Baixa" Then
        MsgBox "Ja foi efetuado Baixa neste arquivo.", vbCritical, "Atenção"
        Exit Sub
    End If
End If

If MsgBox("Deseja efetuar a baixa deste(s) documento(s) ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   MsgBox "Operação cancelada pelo usuário. Nenhuma baixa foi efetuada.", vbInformation, "Atenção"
   Exit Sub
End If

If lstArq.Visible = True Then
    sNomeArq = lstArq.Text
Else
    sNomeArq = "BAIXA MANUAL"
End If

cmdOpcoes.Enabled = False: lstArq.Enabled = False: cmdLoad.Enabled = False
PBar.value = 0

If lblAS.Caption = "S" Then 'SOMENTE PARA SIMPLES NACIONAL
    '*** VALIDAÇÃO DE DOCUMENTOS***
    For nLinha = 1 To UBound(aRegistro)
        If aRegistro(nLinha).nNumDoc = 0 Then
            sCNPJ = aRegistro(nLinha).sCNPJ
            Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
            nNumDoc = RdoAux!maximo + 1
           'CRIA O DOCUMENTO
            Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,emissor) VALUES("
            Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & NomeDeLogin & " (BAIXA BANCÁRIA)" & "')"
            cn.Execute Sql, rdExecDirect
    
            aRegistro(nLinha).nNumDoc = nNumDoc
            For nLinha2 = 1 To UBound(aDoc)
                If aDoc(nLinha2).sCNPJ = aRegistro(nLinha).sCNPJ Then
                    aDoc(nLinha2).nNumDoc = nNumDoc
                   'CRIA PARCELA DOCUMENTO
                   On Error Resume Next
                   Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                    Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & aDoc(nLinha2).nCodReduz & ","
                    Sql = Sql & aDoc(nLinha2).nAno & "," & aDoc(nLinha2).nLanc & "," & aDoc(nLinha2).nSeq & "," & aDoc(nLinha2).nParc & ","
                    Sql = Sql & aDoc(nLinha2).nCompl & "," & nNumDoc & ")"
                    cn.Execute Sql, rdExecDirect
                    On Error GoTo 0
                End If
                
            Next
        End If
    Next
    
    '*** VALIDAÇÃO DE PARCELAS ***
    For nLinha2 = 1 To UBound(aDoc)
        With aDoc(nLinha2)
            Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
            If RdoAux.RowCount = 0 Then
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) "
 '               Sql = Sql & "VALUES(" & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ","
 '               Sql = Sql & 2 & ",'" & Format(.sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & 1 & ",'GTI')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) "
                Sql = Sql & "VALUES(" & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ","
                Sql = Sql & 2 & ",'" & Format(.sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & 1 & ",236)"
                cn.Execute Sql, rdExecDirect
                
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & ","
                Sql = Sql & .nParc & "," & .nCompl & "," & 13 & "," & Virg2Ponto(CStr(.nValorTotal)) & ")"
                cn.Execute Sql, rdExecDirect
                
            End If
            RdoAux.Close
        End With
    Next
    '******************************
End If

For nLinha = 1 To UBound(aRegistro)
    CallPb CLng(nLinha), CLng(UBound(aRegistro))
    If Val(Left(aRegistro(nLinha).sSitRetorno, 2)) > 0 Then GoTo proximo 'APENAS SIT.RETORNO NORMAL PODE SER BAIXADO

    If aRegistro(nLinha).sDataCred <> lblDC.Caption Then GoTo proximo
    ReDim aDocTmp(0)
    If aRegistro(nLinha).bExiste Then
        'CARREGA OS LANCAMENTOS DO DOCUMENTO E COPIA PARA UMA MATRIZ TEMPORÁRIA
        For nLinha2 = 1 To UBound(aDoc)
            If aDoc(nLinha2).nNumDoc = aRegistro(nLinha).nNumDoc And aDoc(nLinha2).nSeqDoc = aRegistro(nLinha).nSeq Then
                ReDim Preserve aDocTmp(UBound(aDocTmp) + 1)
                aDocTmp(UBound(aDocTmp)).nNumDoc = aDoc(nLinha2).nNumDoc
                aDocTmp(UBound(aDocTmp)).nSeqDoc = aDoc(nLinha2).nSeqDoc
                aDocTmp(UBound(aDocTmp)).nCodReduz = aDoc(nLinha2).nCodReduz
                aDocTmp(UBound(aDocTmp)).nAno = aDoc(nLinha2).nAno
                aDocTmp(UBound(aDocTmp)).nLanc = aDoc(nLinha2).nLanc
                aDocTmp(UBound(aDocTmp)).nSeq = aDoc(nLinha2).nSeq
                aDocTmp(UBound(aDocTmp)).nParc = aDoc(nLinha2).nParc
                aDocTmp(UBound(aDocTmp)).nCompl = aDoc(nLinha2).nCompl
                aDocTmp(UBound(aDocTmp)).sDataVencto = aDoc(nLinha2).sDataVencto
                aDocTmp(UBound(aDocTmp)).sSit = aDoc(nLinha2).sSit
                aDocTmp(UBound(aDocTmp)).nPaginaLivro = aDoc(nLinha2).nPaginaLivro
                aDocTmp(UBound(aDocTmp)).nNumeroLivro = aDoc(nLinha2).nNumeroLivro
                aDocTmp(UBound(aDocTmp)).bAjuizado = aDoc(nLinha2).bAjuizado
                aDocTmp(UBound(aDocTmp)).nValorPrincipal = aDoc(nLinha2).nValorPrincipal
                aDocTmp(UBound(aDocTmp)).nValorMulta = aDoc(nLinha2).nValorMulta
                aDocTmp(UBound(aDocTmp)).nValorJuros = aDoc(nLinha2).nValorJuros
                aDocTmp(UBound(aDocTmp)).nValorCorrecao = aDoc(nLinha2).nValorCorrecao
                aDocTmp(UBound(aDocTmp)).nValorTotal = aDoc(nLinha2).nValorTotal
                aDocTmp(UBound(aDocTmp)).nValorTarifa = aDoc(nLinha2).nValorTarifa
                aDocTmp(UBound(aDocTmp)).nValorDif = aDoc(nLinha2).nValorDif
                aDocTmp(UBound(aDocTmp)).nValorCompensado = aDoc(nLinha2).nValorCompensado
                aDocTmp(UBound(aDocTmp)).sBx = aDoc(nLinha2).sBx
                aDocTmp(UBound(aDocTmp)).sDp = aDoc(nLinha2).sDp
                aDocTmp(UBound(aDocTmp)).nSeqReg = aDoc(nLinha2).nSeqReg
                aDocTmp(UBound(aDocTmp)).bExiste = aDoc(nLinha2).bExiste
            End If
        Next
       'EFETUAMOS BAIXA EM CIMA DA MATRIZ TEMPORÁRIA
        For nLinha2 = 1 To UBound(aDocTmp)
            If aDocTmp(nLinha2).bExiste = True Then
               ' If aDocTmp(nLinha2).nNumDoc = 3070426 Then MsgBox "teste"
                
                Sql = "update debitopago set dataintegracao='" & Format(Now, "mm/dd/yyyy") & "' where codreduzido=" & aDocTmp(nLinha2).nCodReduz & " and anoexercicio=" & aDocTmp(nLinha2).nAno & " and "
                Sql = Sql & "codlancamento=" & aDocTmp(nLinha2).nLanc & " and seqlancamento=" & aDocTmp(nLinha2).nSeq & " and numparcela=" & aDocTmp(nLinha2).nParc & " and codcomplemento=" & aDocTmp(nLinha2).nCompl
                cn.Execute Sql, rdExecDirect
                
                '*** INTEGRATIVA ************
                If aDocTmp(nLinha2).nLanc = 20 Then
                    Sql = "SELECT CODREDUZIDO,NUMPROCESSO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA = " & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    sNumProc = SubNull(RdoAux2!numprocesso)
                    RdoAux2.Close
                    
                    If sNumProc <> "" Then
                        ConectaIntegrativa
                    
                        'nNumProc = Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1)
                        nNumproc = ExtraiNumeroProcesso(sNumProc)
                        'nAnoProc = Right$(sNumProc, 4)
                        nAnoproc = ExtraiAnoProcesso(sNumProc)
                                        
                        Sql = "select * from acordos where idacordo=" & nNumproc & " and anoacordo=" & nAnoproc
                        Set RdoAux2 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux2.RowCount > 0 Then
                            RdoAux2.Close
                            Sql = "select * from acordobaixas where idacordo=" & nNumproc & " and anoacordo=" & nAnoproc
                            Set RdoAux2 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            If RdoAux2.RowCount = 0 Then
                               'GRAVA NA TABELA ACORDOBAIXAS
                                Sql = "insert acordobaixas(idAcordo, anoAcordo, DtBaixa, TipoBaixa, NroParcela, VlrOriginal, VlrCorrecao, VlrJuros, VlrMulta, VlrTotal, DtGeracao) values("
                                'Sql = Sql & nNumProc & "," & nAnoProc & ",'" & Format(CDate(lblDC.Caption), "mm/dd/yyyy") & "','PAGAMENTO'," & aDocTmp(nLinha2).nParc & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorPrincipal)) & ","
                                Sql = Sql & nNumproc & "," & nAnoproc & ",'" & Format(CDate(aRegistro(nLinha).sDataCred), "mm/dd/yyyy") & "','PAGAMENTO'," & aDocTmp(nLinha2).nParc & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorPrincipal)) & ","
                                Sql = Sql & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorCorrecao)) & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorJuros)) & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorMulta)) & ","
                                Sql = Sql & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorTotal)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                                cnInt.Execute Sql, rdExecDirect
                            End If
                            RdoAux2.Close
                        Else
                            RdoAux2.Close
                        End If
                        
                        cnInt.Close
                    End If
                Else
                    If aDocTmp(nLinha2).nNumeroLivro > 0 Then
                        ConectaIntegrativa
                        Sql = "select * from pagamentos where iddevedor='" & aDocTmp(nLinha2).nCodReduz & "' and exercicio=" & aDocTmp(nLinha2).nAno & " and "
                        Sql = Sql & "lancamento=" & aDocTmp(nLinha2).nLanc & " and seq=" & aDocTmp(nLinha2).nSeq & " and nroParcela=" & aDocTmp(nLinha2).nParc & " and "
                        Sql = Sql & "complparcela=" & aDocTmp(nLinha2).nCompl
                        Set RdoAux2 = cnInt.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux2.RowCount = 0 Then
                           'GRAVA NA TABELA ACORDOBAIXAS
                            Sql = "insert pagamentos(dtPagamento,idDevedor,NroLivro,NroFolha,Seq,Lancamento,Exercicio,VlrOriginal,VlrJuros,VlrMulta,VlrCorrecao,VlrTotal,"
                            Sql = Sql & "nroParcela,ComplParcela,Ajuizado,DtGeracao) values('" & Format(Now, "mm/dd/yyyy") & "','" & aDocTmp(nLinha2).nCodReduz & "',"
                            Sql = Sql & aDocTmp(nLinha2).nNumeroLivro & "," & aDocTmp(nLinha2).nPaginaLivro & "," & aDocTmp(nLinha2).nSeq & "," & aDocTmp(nLinha2).nLanc & ","
                            Sql = Sql & aDocTmp(nLinha2).nAno & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorPrincipal)) & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorJuros)) & ","
                            Sql = Sql & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorMulta)) & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorCorrecao)) & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorTotal)) & ","
                            Sql = Sql & aDocTmp(nLinha2).nParc & "," & aDocTmp(nLinha2).nCompl & "," & IIf(aDocTmp(nLinha2).bAjuizado, 1, 0) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                            cnInt.Execute Sql, rdExecDirect
                        End If
                        RdoAux2.Close
                        cnInt.Close
                    End If
                End If
                
                '****************************
                
                If aDocTmp(nLinha2).nParc = 0 Then
                    nStatus = 1 'PAGO POR ÚNICA
                Else
                    nStatus = 2 'PAGO POR PARCELA
                End If
                'EFETUA BAIXA NA TABELA DEBITOPARCELA
                 Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                 Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                 Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                 cn.Execute Sql, rdExecDirect
                'SE FOR SIMPLES MARCA
                 If lblAS.Caption = "S" Then
                    Sql = "UPDATE DEBITOPARCELA SET SIMPLESNACIONAL=1 WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                    cn.Execute Sql, rdExecDirect
                 End If
                
                'SE FOR PARCELA UNICA COLOCA PAGO POR UNICA EM TODAS AS PARCELAS, SENÃO CANCELA A PARCELA ÚNICA
                 If nStatus = 1 Then
                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=1  WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA > 0 "
                    cn.Execute Sql, rdExecDirect
                    
                    
                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5  WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA = 0 and CODCOMPLEMENTO<>" & aDocTmp(nLinha2).nCompl & " And STATUSLANC = 3"
                    cn.Execute Sql, rdExecDirect
                    
                    
                 Else
                    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    'Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA = 0 AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA = 0 "
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux.RowCount > 0 Then
                        If RdoAux!statuslanc = 3 And lblTit.Visible = False Then
                            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                            Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                            Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA = 0 AND STATUSLANC=3"
                            cn.Execute Sql, rdExecDirect
                        End If
                    End If
                    RdoAux.Close
                 End If
                
                'EFETUA BAIXA NA TABELA DEBITOPAGO
                Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO = " & aDocTmp(nLinha2).nCompl
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If IsNull(!maximo) Then
                        nSeqPag = 0
                    Else
                        If .RowCount = 0 Then
                           nSeqPag = 0
                        Else
                           nSeqPag = !maximo + 1
                        End If
                    End If
                    .Close
                End With
                If lblDC.Caption = "" Then lblDC.Caption = grdReg.TextMatrix(1, 3)
                If aRegistro(nLinha).sDataPagCalc = "00:00:00" Or aRegistro(nLinha).sDataPagCalc = "01/01/1900" Then
                    aRegistro(nLinha).sDataPagCalc = grdReg.TextMatrix(1, 2)
                End If
                Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL,VALORTARIFA,ARQUIVOBANCO,VALORDIF,DATAPAGAMENTOCALC,CONTACORRENTE) VALUES(" & aDocTmp(nLinha2).nCodReduz & ","
                Sql = Sql & aDocTmp(nLinha2).nAno & "," & aDocTmp(nLinha2).nLanc & "," & aDocTmp(nLinha2).nSeq & "," & aDocTmp(nLinha2).nParc & "," & aDocTmp(nLinha2).nCompl & "," & nSeqPag & ",'"
                'Sql = Sql & Format(aRegistro(nLinha).sDataPag, "mm/dd/yyyy") & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorCompensado + Abs(aDocTmp(nLinha2).nValorDif))) & ","
                Sql = Sql & Format(aRegistro(nLinha).sDataPag, "mm/dd/yyyy") & "','" & Format(aRegistro(nLinha).sDataCred, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorCompensado + Abs(aDocTmp(nLinha2).nValorDif))) & ","
                Sql = Sql & Val(Left(lblBanco.Caption, 3)) & ",'" & Val(RetornaNumero(aRegistro(nLinha).sAgencia)) & "'," & aDocTmp(nLinha2).nNumDoc & "," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorCompensado)) & ","
                Sql = Sql & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorTarifa)) & ",'" & Left(sNomeArq, 50) & "'," & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorDif)) & ",'" & Format(aRegistro(nLinha).sDataPagCalc, "mm/dd/yyyy") & "','" & aRegistro(nLinha).sConta & "')"
                cn.Execute Sql, rdExecDirect
                
                'SE FOR SIMPLES GRAVA NA TABELA DE COMPLEMENTO
                If lblAS.Caption = "S" Then
                    On Error Resume Next
                    Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR,CNPJ,ANO,MES) VALUES(" & aDocTmp(nLinha2).nCodReduz & ","
                    'Sql = Sql & aDocTmp(nLinha2).nAno & "," & aDocTmp(nLinha2).nLanc & "," & aDocTmp(nLinha2).nSeq & "," & aDocTmp(nLinha2).nParc & "," & aDocTmp(nLinha2).nCompl & ",'" & sNomeArq & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "',"
                    Sql = Sql & aDocTmp(nLinha2).nAno & "," & aDocTmp(nLinha2).nLanc & "," & aDocTmp(nLinha2).nSeq & "," & aDocTmp(nLinha2).nParc & "," & aDocTmp(nLinha2).nCompl & ",'" & sNomeArq & "','" & Format(aRegistro(nLinha).sDataCred, "mm/dd/yyyy") & "',"
                    Sql = Sql & Virg2Ponto(CStr(aDocTmp(nLinha2).nValorCompensado)) & ",'" & aRegistro(nLinha).sCNPJ & "'," & aRegistro(nLinha).nAno & "," & aRegistro(nLinha).nMes & ")"
                    cn.Execute Sql, rdExecDirect
                    On Error GoTo 0
                End If
                
                'SE TIVER DIFERENÇA ACIMA DE R$2,00 GERA COMPLEMENTO
                With aDocTmp(nLinha2)
                    If .nValorDif < -2 Then
                         ReDim aTribProp(0): nSomaTrib = 0
                        'PRIMEIRO VERIFICA SE JA EXISTE COMPLEMENTO, SE EXISTIR NÃO CRIA NOVAMENTE
                        Sql = "SELECT * FROM COMPLEMENTOPAGTO WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl & " AND ARQUIVOBANCO='" & Left(sNomeArq, 30) & "'"
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux.RowCount = 0 Then
                            'SE NÃO EXISTIR CRIA O COMPLEMENTO
                            
                            'VAMOS ENCONTRAR O MAXIMO COMPL DA PARCELA
                            Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
                            Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            nComplCP = RdoAux2!maximo + 1
                            RdoAux2.Close
                            
                            'numero do processo se tiver
                            Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
                            Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            sNumProc = SubNull(RdoAux2!numprocesso)
                            RdoAux2.Close
                            
                            'GERA O COMPLEMENTO EM DEBITOPARCELA E DEBITOTRIBUTO
'                            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'                            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMPROCESSO,USUARIO) "
'                            Sql = Sql & "VALUES(" & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & nComplCP & ","
'                            Sql = Sql & 25 & ",'" & Format(lblDC.Caption, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "','GTI')"
                            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMPROCESSO,USERID) "
                            Sql = Sql & "VALUES(" & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & nComplCP & ","
                            Sql = Sql & 25 & ",'" & Format(lblDC.Caption, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "',236)"
                            cn.Execute Sql, rdExecDirect
                            
                            'GERA MATRIZ PROPORCIONAL DE TRIBUTOS
                            Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
                            Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl & " AND CODTRIBUTO<>3"
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                Do Until .EOF
                                    ReDim Preserve aTribProp(UBound(aTribProp) + 1)
                                    aTribProp(UBound(aTribProp)).nCodTrib = !CodTributo
                                    aTribProp(UBound(aTribProp)).nValorTrib = !ValorTributo
                                    nSomaTrib = nSomaTrib + !ValorTributo
                                   .MoveNext
                                Loop
                               .Close
                            End With
                                                                                                                                        
                            For k = 1 To UBound(aTribProp)
                                aTribProp(k).nPerc = aTribProp(k).nValorTrib * 100 / nSomaTrib
                            Next
                            For k = 1 To UBound(aTribProp)
                                aTribProp(k).nNovoValor = Round(Abs(.nValorDif) * aTribProp(k).nPerc / 100, 2)
                            Next
                                                        
                            For k = 1 To UBound(aTribProp)
                                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                                Sql = Sql & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & ","
                                Sql = Sql & .nParc & "," & nComplCP & "," & aTribProp(k).nCodTrib & "," & Virg2Ponto(CStr(aTribProp(k).nNovoValor)) & ")"
                                cn.Execute Sql, rdExecDirect
                            Next
                            
                            'GERA O COMPLEMENTO EM COMPLEMENTOPAGTO
                            Sql = "INSERT COMPLEMENTOPAGTO(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODCOMPLEMENTOCP,ARQUIVOBANCO,DATACREDITO,VALOR) VALUES(" & aDocTmp(nLinha2).nCodReduz & ","
                            Sql = Sql & aDocTmp(nLinha2).nAno & "," & aDocTmp(nLinha2).nLanc & "," & aDocTmp(nLinha2).nSeq & "," & aDocTmp(nLinha2).nParc & "," & aDocTmp(nLinha2).nCompl & "," & nComplCP & ",'" & sNomeArq & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "',"
                            Sql = Sql & Virg2Ponto(CStr(Abs(aDocTmp(nLinha2).nValorDif))) & ")"
                            cn.Execute Sql, rdExecDirect
                            
                            'GRAVA OBS PARCELA
                            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & .nCodReduz & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND "
                            Sql = Sql & "SEQLANCAMENTO=" & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & nComplCP
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                If IsNull(!maximo) Then
                                    nSeq2 = 1
                                Else
                                    nSeq2 = !maximo + 1
                                End If
                               .Close
                            End With
                            sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
                            sObs = "Complemento gerado automaticamente pela diferença de pagamento da parcela." & vbCrLf & "Valores devidos: Vl.Prin.="
                            sObs = sObs & FormatNumber(.nValorPrincipal, 2) & " Vl.Jur.=" & FormatNumber(.nValorJuros, 2) & " Vl.Mul.=" & FormatNumber(.nValorMulta, 2) & " Vl.Cor.=" & FormatNumber(.nValorCorrecao, 2)
                            sObs = sObs & " Vl.Tarifa=" & FormatNumber(.nValorTarifa, 2) & " Vl.Total=" & FormatNumber(.nValorTotal + .nValorTarifa, 2) & " Valor Pago=" & FormatNumber(.nValorCompensado, 2) & " Vl.Dif.=" & FormatNumber(Abs(.nValorDif), 2)
'                            Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & .nCodReduz & "," & .nAno & ","
'                            Sql = Sql & .nLanc & "," & .nSeq & "," & .nParc & "," & nComplCP & "," & nSeq2 & ",'" & sObs & "','GTI AUTOMÁTICO','" & Format(sData, "mm/dd/yyyy") & "')"
                            Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & .nCodReduz & "," & .nAno & ","
                            Sql = Sql & .nLanc & "," & .nSeq & "," & .nParc & "," & nComplCP & "," & nSeq2 & ",'" & sObs & "',236,'" & Format(sData, "mm/dd/yyyy") & "')"
                            cn.Execute Sql, rdExecDirect
                                                       
                            
                        End If
                        RdoAux.Close
                    End If
                End With
            Else
                'DOCUMENTO SEM LANÇAMENTOS
                GoTo DocumentoErro
            End If
            
           'EFETUA BAIXA NO DOCUMENTO
            Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=" & Val(Left(lblBanco.Caption, 3)) & " ,CODAGENCIA ='" & aRegistro(nLinha).sAgencia & "' , VALORPAGO=" & Virg2Ponto(CStr(aRegistro(nLinha).nValorPago))
            Sql = Sql & " WHERE NUMDOCUMENTO=" & aRegistro(nLinha).nNumDoc
            cn.Execute Sql, rdExecDirect

        Next
    Else
DocumentoErro:
        'DOCUMENTO NÃO ENCONTRADO
        Sql = "INSERT RECEITACLASSIFICAR (NOMEARQ,DATARECEITA,CODBANCO,NUMDOCUMENTO,VALORTOTAL) VALUES('"
        Sql = Sql & sNomeArq & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "'," & Val(Left(lblBanco.Caption, 3)) & ","
        Sql = Sql & aRegistro(nLinha).nNumDoc & "," & Virg2Ponto(CStr(aRegistro(nLinha).nValorPago)) & ")"
        cn.Execute Sql, rdExecDirect
    End If
proximo:
Next

PBar.value = 0
cmdOpcoes.Enabled = True: lstArq.Enabled = True: cmdLoad.Enabled = True

lblDB.Caption = Format(Now, "dd/mm/yyyy")

'Sql = "SELECT * FROM ARQUIVOBANCO WHERE NOMEARQ='" & sNomeArq & "' AND DATACREDITO='" & Format(lblDC.Caption, "mm/dd/yyyy") & "'"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'If RdoAux.RowCount = 0 Then
'    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(lblDC.Caption, "mm/dd/yyyy") & "'"
'    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    If IsNull(RdoAux2!maximo) Then
'        nSeq = 1
'    Else
'        nSeq = RdoAux2!maximo + 1
'    End If

'    Sql = "INSERT ARQUIVOBANCO(DATACREDITO,SEQ,CODBANCO,CODAGENCIA,DATAINCLUSAO,NOMEARQ,DA) VALUES('"
'    Sql = Sql & Format(CDate(lblDC.Caption), "mm/dd/yyyy") & "'," & nSeq & "," & Val(Left(lblBanco.Caption, 3)) & "," & "0" & ",'"
'    Sql = Sql & Format(Now, "mm/dd/yyyy") & "','" & sNomeArq & "'," & 0 & ")"
'Else
    Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA='" & Format(Now, "mm/dd/yyyy") & "' WHERE NOMEARQ='" & sNomeArq & "' AND DATACREDITO='" & Format(lblDC.Caption, "mm/dd/yyyy") & "'"
'End If
cn.Execute Sql, rdExecDirect

'GravaAnalise

MsgBox "Baixa efetuada com sucesso.", vbInformation, "Informação"
PBar.value = 0

If lstArq.Visible = False Then
    ReDim aReg(0): ReDim aDoc(0)
    grdReg.Rows = 1
    grdParc.Rows = 1
End If

End Sub

Private Sub ReativarArquivo()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nNumDoc As Long
Dim Sql As String, RdoAux As rdoResultset, aDocTmp() As Documento, RdoAux2 As rdoResultset, sNomeArq As String

If lblAS.Caption = "S" Then
    MsgBox "Não é possível reativar arquivo do Simple Nacional.", vbCritical, "ERRO"
    Exit Sub
End If

If grdReg.Rows = 1 Then
    MsgBox "Carregue um arquivo válido para reativar.", vbExclamation, "Atenção"
    Exit Sub
End If

If lblDB.Caption = "Sem Baixa" Then
    If MsgBox("Não foi efetuado Baixa neste arquivo." & vbCrLf & "Deseja forçar a reativação assim mesmo?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
       Exit Sub
    End If
End If

If MsgBox("Deseja REATIVAR as baixas deste(s) documento(s) ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   MsgBox "Operação cancelada pelo usuário. Nenhuma baixa foi reativada.", vbInformation, "Atenção"
   Exit Sub
End If
sNomeArq = lstArq.Text
cmdOpcoes.Enabled = False: lstArq.Enabled = False: cmdLoad.Enabled = False

'APAGA DA TABELA RECEITACLASSIFICAR
Sql = "DELETE FROM RECEITACLASSIFICAR WHERE DATARECEITA='" & Format(cmbDataCredito.Text, "mm/dd/yyyy") & "' AND NOMEARQ='" & lstArq.Text & "'"
cn.Execute Sql, rdExecDirect

PBar.value = 0
For nLinha = 1 To UBound(aRegistro)
    CallPb CLng(nLinha), CLng(UBound(aRegistro))
    ReDim aDocTmp(0)
    If aRegistro(nLinha).bExiste Then
        'CARREGA OS LANCAMENTOS DO DOCUMENTO E COPIA PARA UMA MATRIZ TEMPORÁRIA
        For nLinha2 = 1 To UBound(aDoc)
            If aDoc(nLinha2).nNumDoc = aRegistro(nLinha).nNumDoc Then
                ReDim Preserve aDocTmp(UBound(aDocTmp) + 1)
                aDocTmp(UBound(aDocTmp)).nNumDoc = aDoc(nLinha2).nNumDoc
                aDocTmp(UBound(aDocTmp)).nCodReduz = aDoc(nLinha2).nCodReduz
                aDocTmp(UBound(aDocTmp)).nAno = aDoc(nLinha2).nAno
                aDocTmp(UBound(aDocTmp)).nLanc = aDoc(nLinha2).nLanc
                aDocTmp(UBound(aDocTmp)).nSeq = aDoc(nLinha2).nSeq
                aDocTmp(UBound(aDocTmp)).nParc = aDoc(nLinha2).nParc
                aDocTmp(UBound(aDocTmp)).nCompl = aDoc(nLinha2).nCompl
                aDocTmp(UBound(aDocTmp)).sDataVencto = aDoc(nLinha2).sDataVencto
                aDocTmp(UBound(aDocTmp)).sSit = aDoc(nLinha2).sSit
                aDocTmp(UBound(aDocTmp)).nValorPrincipal = aDoc(nLinha2).nValorPrincipal
                aDocTmp(UBound(aDocTmp)).nValorMulta = aDoc(nLinha2).nValorMulta
                aDocTmp(UBound(aDocTmp)).nValorJuros = aDoc(nLinha2).nValorJuros
                aDocTmp(UBound(aDocTmp)).nValorCorrecao = aDoc(nLinha2).nValorCorrecao
                aDocTmp(UBound(aDocTmp)).nValorTotal = aDoc(nLinha2).nValorTotal
                aDocTmp(UBound(aDocTmp)).nValorTarifa = aDoc(nLinha2).nValorTarifa
                aDocTmp(UBound(aDocTmp)).nValorDif = aDoc(nLinha2).nValorDif
                aDocTmp(UBound(aDocTmp)).nValorCompensado = aDoc(nLinha2).nValorCompensado
                aDocTmp(UBound(aDocTmp)).sBx = aDoc(nLinha2).sBx
                aDocTmp(UBound(aDocTmp)).sDp = aDoc(nLinha2).sDp
                aDocTmp(UBound(aDocTmp)).nSeqReg = aDoc(nLinha2).nSeqReg
                aDocTmp(UBound(aDocTmp)).bExiste = aDoc(nLinha2).bExiste
            End If
        Next
       'EFETUAMOS A REATIVAÇÃO EM CIMA DA MATRIZ TEMPORÁRIA
        For nLinha2 = 1 To UBound(aDocTmp)
            If aDocTmp(nLinha2).bExiste = True Then
                If aDocTmp(nLinha2).nParc = 0 Then
                    nStatus = 1 'PAGO POR ÚNICA
                Else
                    nStatus = 2 'PAGO POR PARCELA
                End If
                'EFETUA BAIXA NA TABELA DEBITOPARCELA
                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3 WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                cn.Execute Sql, rdExecDirect
                'SE FOR PARCELA UNICA REATIVA TODAS AS PARCELAS
                If nStatus = 1 Then
                   Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3  WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                   Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                   Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                   cn.Execute Sql, rdExecDirect
                End If
                'SE FOR SIMPLES NACIONAL REMOVEMOS O COMPLEMENTO
                If lblAS.Caption = "S" Then
                    Sql = "DELETE FROM COMPLEMENTOSIMPLES WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                    cn.Execute Sql, rdExecDirect
                    'ATUALIZA DOCUMENTO
                    Sql = "SELECT NUMDOCUMENTO FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    If RdoAux.RowCount > 0 Then
                        Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=0,CODAGENCIA=0,VALORPAGO=0 "
                        Sql = Sql & "WHERE NUMDOCUMENTO = " & RdoAux!NumDocumento
                        cn.Execute Sql, rdExecDirect
                    End If
                    
                End If
                
                'SE TIVER COMPLEMENTO DE DIF DE PAGTO
                Sql = "SELECT * FROM COMPLEMENTOPAGTO WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl & " AND ARQUIVOBANCO='" & Left(sNomeArq, 30) & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    nCompl = RdoAux!CODCOMPLEMENTOCP
                    'CANCELA O COMPLEMENTO DE DEBITOPARCELA SOMENTE SE NAO ESTIVER PAGO
                    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                    Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & nCompl
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux2!statuslanc = 3 Then
                        'APAGA O COMPELEMENTO DA TABELA COMPLEMENTO
                        Sql = "DELETE FROM COMPLEMENTOPAGTO WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & aDocTmp(nLinha2).nCompl & " AND ARQUIVOBANCO='" & Left(sNomeArq, 30) & "'"
                        cn.Execute Sql, rdExecDirect
                        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & aDocTmp(nLinha2).nCodReduz & " AND ANOEXERCICIO=" & aDocTmp(nLinha2).nAno & " AND CODLANCAMENTO=" & aDocTmp(nLinha2).nLanc & " AND "
                        Sql = Sql & "SEQLANCAMENTO=" & aDocTmp(nLinha2).nSeq & " AND NUMPARCELA=" & aDocTmp(nLinha2).nParc & " AND CODCOMPLEMENTO=" & nCompl
                        cn.Execute Sql, rdExecDirect
                    End If
                    RdoAux2.Close
                End If
                RdoAux.Close
            Else
                'DOCUMENTO SEM LANÇAMENTOS
                GoTo DocumentoErro
            End If
        Next
    Else
DocumentoErro:
        'DOCUMENTO NÃO ENCONTRADO
    End If
    
    'ATUALIZA A TABELA NUMDOCUMENTO
    Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=0,CODAGENCIA=0,VALORPAGO=0 "
    Sql = Sql & "WHERE NUMDOCUMENTO = " & aRegistro(nLinha).nNumDoc
    cn.Execute Sql, rdExecDirect
    
Next
Ocupado
'REATIVAMOS NA TABELA DEBITOPAGO (O NOME DO ARQUIVO NÃO ERA GRAVADO ANTES, ENTÃO TEMOS QUE DIFERENCIAR)
Sql = "SELECT * FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(cmbDataCredito.Text, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL AND ARQUIVOBANCO='" & lstArq.Text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    Sql = "DELETE FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(cmbDataCredito.Text, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL  AND ARQUIVOBANCO='" & lstArq.Text & "'"
    cn.Execute Sql, rdExecDirect
Else
    RdoAux.Close
    Sql = "SELECT * FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(cmbDataCredito.Text, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL AND ARQUIVOBANCO='" & lstArq.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        Sql = "DELETE FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(cmbDataCredito.Text, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL AND ARQUIVOBANCO='" & lstArq.Text & "'"
        cn.Execute Sql, rdExecDirect
    End If
End If

PBar.value = 0

Sql = "DELETE FROM ANALISE2 WHERE USUARIO='" & NomeDeLogin & "' AND DATARECEITA='" & Format(lblDC.Caption, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left(lblBanco.Caption, 3)) & " AND ARQUIVO='" & lstArq.Text & "'"
'cn.Execute Sql, rdExecDirect

'FIM DA ROTINA
lblDB.Caption = "Sem Baixa"
Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA=NULL WHERE NOMEARQ='" & lstArq.Text & "' AND DATACREDITO='" & Format(lblDC.Caption, "mm/dd/yyyy") & "'"
cn.Execute Sql, rdExecDirect
cmdOpcoes.Enabled = True: lstArq.Enabled = True: cmdLoad.Enabled = True
Liberado
LimpaTela
MsgBox "Todos os lançamentos descriminados e seus documentos foram reativados.", vbInformation, "INFORMAÇÃO"
End Sub

Private Sub ResumoArquivo()
Dim Sql As String, x As Integer, sDataBaixa As String

If grdReg.Rows = 1 Then
    MsgBox "Arquivo não carregado.", vbExclamation, "Atenção"
    Exit Sub
End If

'Sql = "DELETE FROM RESUMOBANCO3 WHERE COMPUTER='" & NomeDeLogin & "';"
Sql = "DELETE FROM RESUMOBANCO2 WHERE COMPUTER='" & NomeDeLogin & "';"
Sql = Sql & "DELETE FROM RESUMOBANCO1 WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

sDataBaixa = lblDB.Caption

Sql = "INSERT RESUMOBANCO1(COMPUTER,ARQUIVO,BANCO,DATACREDITO,DATABAIXA,ARQS,ARQC,ARQDA,NUMREGISTRO,VALORTOTAL) VALUES('"
Sql = Sql & NomeDeLogin & "','" & lstArq.Text & "','" & lblBanco.Caption & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "','" & sDataBaixa & "','"
Sql = Sql & lblAS.Caption & "','" & lblAC.Caption & "','" & lblDA.Caption & "'," & Val(lblNumReg.Caption) & "," & Virg2Ponto(RemovePonto(lblValorTotal.Caption)) & ")"
cn.Execute Sql, rdExecDirect

For x = 1 To UBound(aRegistro)
    With aRegistro(x)
        Sql = "INSERT RESUMOBANCO2 (COMPUTER,ARQUIVO,NUMDOCUMENTO,DATAPAGAMENTO,VALORPAGO,VALORTARIFA,ISENTOMJ,SITUACAO) VALUES('"
        Sql = Sql & NomeDeLogin & "','" & lstArq.Text & "','" & Format(.nNumDoc, "000000000") & "','" & Format(.sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(.nValorPago)) & ","
        Sql = Sql & Virg2Ponto(CStr(Round(.nValorTarifa, 2))) & ",'" & IIf(.bIsentoMJ, "S", "N") & "','" & .sSitRetorno & "')"
        cn.Execute Sql, rdExecDirect
    End With
Next

For x = 1 To UBound(aDoc)
    With aDoc(x)
        Sql = "INSERT RESUMOBANCO3 (COMPUTER,ARQUIVO,NUMDOCUMENTO,CODREDUZIDO,ANO,LANC,SEQ,PARC,COMPL,DATAVENCTO,PRINCIPAL,MULTA,JUROS,CORRECAO,TOTAL,TARIFA,DIF,COMPENSADO,DUP) VALUES('"
        Sql = Sql & NomeDeLogin & "','" & lstArq.Text & "','" & Format(.nNumDoc, "000000000") & "','" & Format(.nCodReduz, "000000") & "','" & .nAno & "','" & Format(.nLanc, "000") & "','" & Format(.nSeq, "000") & "','" & Format(.nParc, "000") & "','"
        Sql = Sql & Format(.nCompl, "00") & "','" & Format(.sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(.nValorPrincipal)) & "," & Virg2Ponto(CStr(.nValorMulta)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorJuros)) & "," & Virg2Ponto(CStr(.nValorCorrecao)) & "," & Virg2Ponto(CStr(.nValorTotal)) & "," & Virg2Ponto(CStr(.nValorTarifa)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorDif)) & "," & Virg2Ponto(CStr(.nValorCompensado)) & ",'" & .sDp & "')"
        cn.Execute Sql, rdExecDirect
    End With
Next

frmReport.ShowReport "RESUMOBANCO", frmMdi.HWND, Me.HWND

Sql = "DELETE FROM RESUMOBANCO3 WHERE COMPUTER='" & NomeDeLogin & "';"
Sql = Sql & "DELETE FROM RESUMOBANCO2 WHERE COMPUTER='" & NomeDeLogin & "';"
Sql = Sql & "DELETE FROM RESUMOBANCO1 WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mskDataCred_GotFocus()
mskDataCred.SetFocus
End Sub

Private Sub mskDataPag_GotFocus()
mskDataPag.SetFocus
End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
Tweak txtAgencia, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
Tweak txtNumDoc, KeyAscii, IntegerPositive
End Sub

Private Sub txtValorPago_KeyPress(KeyAscii As Integer)
Tweak txtValorPago, KeyAscii, DecimalPositive
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Sub LeArquivo()
Dim sFullPath As String, sReg As String, nPos As Long, nTot As Long, FF1 As Integer, nCodBanco As Integer, bExec As Boolean, sTipoArq As String
Dim nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorGuia As Double, nNumDoc As Long, nErro As Integer, RdoAux4 As rdoResultset
Dim sAno As String, sMes As String, sAgencia As String, bLayoutNovo As Boolean, RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, RdoAux3 As rdoResultset
Dim nNumParc As Integer, bAchou As Boolean, nSeq As Integer, nCompl As Integer, nCodReduz As Long, sDataVencto As String, nRetorno As Integer, sRetorno As String
Dim nSeqReg As Integer, nValorTaxa As Double, R As Integer, sDataGeracao As String, sLinhaT As String, sLinhaU As String
Dim nData As Integer, bData As Boolean

ReDim aRegistro(0): ReDim aDoc(0): ReDim aTrib(0)
nSeqReg = 1: grdTrib.Rows = 1

'*** VERIFICA EXISTENCIA DO ARQUIVO

If txtPath.Text = "" Then Exit Sub
sFullPath = Replace(sFullPath, "/", "\")
sFullPath = txtPath.Text & lstArq.Text

If Dir$(sFullPath) = "" Then
    MsgBox "Não localizado o arquivo em " & sFullPath, vbCritical, "ERRO FATAL !!!"
    Exit Sub
End If

nPos = 0: nTot = 0: PBar.value = 0: nValorEfetivo = 0
Ocupado
lblAS.Caption = "N": lblAC.Caption = "N": lblDA.Caption = "N"

sTipoArq = ""
'**********************************
'****** CABEÇALHO DO ARQUIVO ******
'**********************************
On Error Resume Next
FF1 = FreeFile(): nPos = 0: sTipoArq = ""
Open sFullPath For Binary Access Read Write As FF1
    Do While Not EOF(FF1)
        On Error GoTo CloseFile0
        If sReg = "48" Then Exit Do
        Input #FF1, sReg
        
        If Left(sReg, 1) = "A" And InStr(1, sReg, "DEBITO AUT") = 0 And sTipoArq <> "CBR724" Then 'ARQUIVO NORMAL
            sTipoArq = "NORMAL"
        ElseIf Left(sReg, 15) = "100000001DAF607" Then 'ARQUIVO SIMPLES
            sTipoArq = "SIMPLES"
        ElseIf Left(sReg, 9) = "02RETORNO" Then  'ARQUIVO COBRANÇA LAYOUT ANTIGO
            sTipoArq = "COBRANCAA"
        ElseIf Left(sReg, 8) = "03300000" Then  'ARQUIVO COBRANÇA LAYOUT NOVO
            sTipoArq = "COBRANCAN"
        ElseIf Left(sReg, 8) = "00100000" Then  'ARQUIVO COBRANÇA BANCO DO BRASIL
            sTipoArq = "COBRANCABB"
        ElseIf InStr(1, sReg, "DEBITO AUT") > 0 Then 'ARQUIVO DEBITO AUTOMATICO
            sTipoArq = "DEBAUT"
        ElseIf Mid(sReg, 7, 6) = "CBR724" Then
          '  Close #FF1
            sTipoArq = "CBR724"
            If Left(sReg, 2) = "48" Then
              Exit Do
              Close #FF1
            End If
        End If
        nTot = nTot + 1
    Loop
CloseFile0:
Close #FF1

On Error GoTo 0
If sTipoArq = "" Then
    Liberado
    MsgBox "ARQUIVO DESCONHECIDO.", vbCritical, "ERRO CRÍTICO"
    Exit Sub
End If

sReg = ""

If sTipoArq = "NORMAL" Then
    GoTo LEARQNORMAL
ElseIf sTipoArq = "SIMPLES" Then
    GoTo LEARQSIMPLES
ElseIf sTipoArq = "COBRANCAA" Then
    bLayoutNovo = False
    GoTo LEARQCOBRANCA
ElseIf sTipoArq = "COBRANCAN" Then
    bLayoutNovo = True
    nTot = nTot / 2
    GoTo LEARQCOBRANCA
ElseIf sTipoArq = "COBRANCABB" Then
    nTot = nTot / 2
    GoTo LEARQCOBRANCABB
ElseIf sTipoArq = "DEBAUT" Then
    GoTo LEDEBAUT
ElseIf sTipoArq = "CBR724" Then
    GoTo LECBR724
End If


'***********************
'******* CBR724 ********
'***********************
LECBR724:
'FF1 = FreeFile(): nPos = 0: sTipoArq = ""
'Open sFullPath For Binary Access Read Write As FF1
'    Do While Not EOF(FF1)
'        On Error GoTo CloseFile0
'        Input #FF1, sReg
'        If nPos Mod 50 = 0 Then CallPb nPos, nTot
'        If Mid(sReg, 4, 7) = "2873532" Then
'            nNumDoc = Mid(sReg, 22, 8)
'            Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
'            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux.RowCount > 0 Then
'                ReDim Preserve aRegistro(UBound(aRegistro) + 1)
'                aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
'                aRegistro(UBound(aRegistro)).nSeq = nSeqReg
'                aRegistro(UBound(aRegistro)).sConta = "740004"
'                aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
'                aRegistro(UBound(aRegistro)).sDataCred = Format(Now, "dd/mm/yyyy")
'                aRegistro(UBound(aRegistro)).sDataPag = Format(Now, "dd/mm/yyyy")
'                aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 90, 17))
'                aRegistro(UBound(aRegistro)).bExiste = True
'                aRegistro(UBound(aRegistro)).sSitRetorno = "00-BAIXA NORMAL"
'                aRegistro(UBound(aRegistro)).bIsentoMJ = False
'                aRegistro(UBound(aRegistro)).sDataPagCalc = Format(Now, "dd/mm/yyyy")
'                CarregaParcela nNumDoc, nSeqReg, Format(Now, "dd/mm/yyyy")
'                nSeqReg = nSeqReg + 1
'            End If
'        End If
'
'        nPos = nPos + 1
'    Loop
'Close #FF1
'Liberado
'nErro = 0
'
'
'For nPos = 1 To UBound(aRegistro)
'
'    With aRegistro(nPos)
'        If aRegistro(nPos).bExiste = True Then
'            grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & .sDataDoc & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
'            Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
'            IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq & Chr(9) & Format(CDate(.sDataPagCalc), "dd/mm/yyyy")
'        Else
'            nErro = nErro + 1
'            grdReg.AddItem Format(.nNumDoc, "000000000")
'        End If
'        If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
'            nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
'        End If
'    End With
'Next
'lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)
'For nPos = 1 To UBound(aDoc)
'    If aDoc(nPos).bExiste = False Then
'        nErro = nErro + 1
'    End If
'Next
'
'If nErro > 0 Then
'    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
'    lblErro.Visible = True: cmdErro.Visible = True
'End If
'
'If grdReg.Rows > 1 Then
'    For x = 1 To grdReg.Rows - 1
'        If grdReg.TextMatrix(x, 3) <> "" Then
'            lblDC.Caption = grdReg.TextMatrix(x, 3)
'            Exit For
'        End If
'    Next
'
'Else
'    lblDC.Caption = "Sem Registros"
'End If
'
'Exit Sub


Exit Sub
'****************************
'****** ARQUIVO NORMAL ******
'****************************
LEARQNORMAL:
ConectaEicon
grdReg.TextMatrix(0, 2) = "Data Pagam."
Open sFullPath For Binary Access Read Write As FF1
    While Not EOF(FF1)
        If nPos Mod 10 = 0 Then CallPb nPos, nTot
        On Error GoTo CloseFile1
        If Left(sReg, 1) = "Z" Then GoTo CloseFile1
        bExec = False
        Input #FF1, sReg
        If Left(sReg, 1) = "A" Then
            sSeqArq = Mid(sReg, 74, 6)
        ElseIf Left(sReg, 1) = "G" Then
           'LE OS REGISTROS
            With grdReg
                nNumDoc = Val(Mid(sReg, 65, 9))
                GoTo Test1
Reduz1:
                nNumDoc = Val(Mid(sReg, 65, 8))
Aumenta1:
                nNumDoc = Val(Mid(sReg, 68, 8))
Test1:
                Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                
                    'If nNumDoc = 13896501 Then MsgBox "teste"
                    If IsNull(RdoAux!ValorTaxaDoc) Or RdoAux!ValorTaxaDoc = 0 Then
                       Sql = "SELECT debitotributo.valortributo FROM debitotributo INNER JOIN parceladocumento ON debitotributo.codreduzido = parceladocumento.codreduzido AND "
                       Sql = Sql & "debitotributo.anoexercicio = parceladocumento.anoexercicio AND debitotributo.codlancamento = parceladocumento.codlancamento AND "
                       Sql = Sql & "debitotributo.seqlancamento = parceladocumento.seqlancamento AND debitotributo.numparcela = parceladocumento.numparcela AND debitotributo.CODCOMPLEMENTO = parceladocumento.CODCOMPLEMENTO "
                       Sql = Sql & "Where (parceladocumento.NumDocumento = " & nNumDoc & ") And (debitotributo.CodTributo = 3)"
                       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                       With RdoAux2
                           If .RowCount > 0 Then
                              nValorTaxa = FormatNumber(!ValorTributo, 2)
                           Else
                              nValorTaxa = "0,00"
                           End If
                          .Close
                       End With
                    Else
                       nValorTaxa = FormatNumber(RdoAux!ValorTaxaDoc, 2)
                       If nValorTaxa = 4 Then nValorTaxa = 0
                    End If
                
                
                    ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                    aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                    aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                    aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
                    aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 22, 8))
                    aRegistro(UBound(aRegistro)).sDataCred = ConvDataSerial(Mid(sReg, 30, 8))
                    aRegistro(UBound(aRegistro)).nValorPago = (CDbl(Mid(sReg, 83, 11)) / 100)
                    aRegistro(UBound(aRegistro)).nValorTarifaBancaria = (CDbl(Mid(sReg, 94, 7)) / 100)
                    aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 109, 4)
                    aRegistro(UBound(aRegistro)).nValorTarifa = nValorTaxa
                    aRegistro(UBound(aRegistro)).sSitRetorno = "00-BAIXA NORMAL"
                    aRegistro(UBound(aRegistro)).bExiste = True
                    aRegistro(UBound(aRegistro)).bIsentoMJ = IIf(Val(SubNull(RdoAux!isentomj)) = 0, False, True)
                    aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                    CarregaParcela nNumDoc, nSeqReg, ConvDataSerial(Mid(sReg, 30, 8))
                
                    
                    If (nNumDoc > 2000000 And nNumDoc < 3000000) Or nNumDoc < 2000 Then
                        '***** GRAVA BAIXA NA GISS ***************
                        Sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
                        Sql = Sql & 2177 & "," & Val(Left(lblBanco.Caption, 3)) & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
                        Sql = Sql & lstArq.Text & "','" & Format(aRegistro(UBound(aRegistro)).sDataCred, "mm/dd/yyyy") & "')"
                        cnEicon.Execute Sql, rdExecDirect
                        
                        For x = 1 To UBound(aDoc)
                            If aDoc(x).nNumDoc = nNumDoc Then
                                Exit For
                            End If
                        Next
                        Sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
                        Sql = Sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & Val(Left(lblBanco.Caption, 3)) & "," & 0 & "," & nNumDoc & "," & nPos & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
                        Sql = Sql & Virg2Ponto(CStr(aDoc(x).nValorTotal)) & "," & Virg2Ponto(CStr(aRegistro(UBound(aRegistro)).nValorPago)) & ",'" & Format(aRegistro(UBound(aRegistro)).sDataPag, "mm/dd/yyyy") & "'," & 0 & ",'"
                        Sql = Sql & "" & "','" & "" & "')"
                        cnEicon.Execute Sql, rdExecDirect
                    
                    '*****************************************
                    End If
                
                
                Else
                    If Not bExec And (Val(Mid(sReg, 65, 1)) > 0 Or Val(Mid(sReg, 65, 2)) <> "01") And Len(CStr(nNumDoc)) > 7 Then
                        bExec = True
                        GoTo Reduz1
                    ElseIf Not bExec And (Val(Mid(sReg, 65, 1)) > 0 Or Val(Mid(sReg, 65, 2)) <> "01") And Len(CStr(nNumDoc)) > 5 And nNumDoc < 2000000 Then
                        bExec = True
                        GoTo Aumenta1
                    Else
                        ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                        aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                        aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                        aRegistro(UBound(aRegistro)).sDataDoc = ""
                        aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 22, 8))
                        aRegistro(UBound(aRegistro)).sDataCred = ConvDataSerial(Mid(sReg, 30, 8))
                        aRegistro(UBound(aRegistro)).nValorPago = (CDbl(Mid(sReg, 83, 11)) / 100)
                        aRegistro(UBound(aRegistro)).nValorTarifaBancaria = (CDbl(Mid(sReg, 94, 7)) / 100)
                        aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 109, 4)
                        aRegistro(UBound(aRegistro)).nValorTarifa = 0
                        aRegistro(UBound(aRegistro)).sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                        aRegistro(UBound(aRegistro)).bExiste = False
                        aRegistro(UBound(aRegistro)).bIsentoMJ = False
                        aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                    End If
                End If
                DoEvents
                nSeqReg = nSeqReg + 1
                RdoAux.Close
            End With
        ElseIf Left(sReg, 1) = "Z" Then
           'LE O RODAPÉ DO ARQUIVO
            lblNumReg.Caption = Format(grdReg.Rows - 1, "000000")
            'lblNumReg.Caption = Format(Val(Mid(sReg, 2, 6)) - 2, "000000")
            lblValorTotal.Caption = FormatNumber(CDbl(Mid(sReg, 8, 17) / 100), 2)
        End If
        nPos = nPos + 1
    Wend
CloseFile1:
Close #FF1
Liberado
PBar.Color = vbWhite
PBar.value = 0
nErro = 0
cmbDataCredito.Clear

For nPos = 1 To UBound(aRegistro)
    
    bData = False
    For nData = 0 To cmbDataCredito.ListCount - 1
        If aRegistro(nPos).sDataCred = cmbDataCredito.List(nData) Then
            bData = True
        End If
    Next
    If bData = False Then
        cmbDataCredito.AddItem aRegistro(nPos).sDataCred
    End If
    
    
'    With aRegistro(nPos)
'        If aRegistro(nPos).bExiste = True Then
'            grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & .sDataDoc & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
'            Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
'            IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq & Chr(9) & Format(CDate(.sDataPagCalc), "dd/mm/yyyy")
'        Else
'            nErro = nErro + 1
'            grdReg.AddItem Format(.nNumDoc, "000000000")
'        End If
'        If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
'            nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
'        End If
'    End With
Next
cmbDataCredito.ListIndex = 0

'lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)
'For nPos = 1 To UBound(aDoc)
'    If aDoc(nPos).bExiste = False Then
'        nErro = nErro + 1
'    End If
'Next

'If nErro > 0 Then
'    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
'    lblErro.Visible = True: cmdErro.Visible = True
'End If

'If grdReg.Rows > 1 Then
'    For x = 1 To grdReg.Rows - 1
'        If grdReg.TextMatrix(x, 3) <> "" Then
'            lblDC.Caption = grdReg.TextMatrix(x, 3)
'            Exit For
'        End If
'    Next
    
'Else
'    lblDC.Caption = "Sem Registros"
'End If

Exit Sub

'*****************************************
'****** ARQUIVO DO SIMPLES NACIONAL ******
'*****************************************
LEARQSIMPLES:
lblAS.Caption = "S"
grdReg.TextMatrix(0, 2) = "Competên."
'** TROCA OS BANCOS PELOS BANCOS VIRTUAIS **
nCodBanco = Val(Left(lblBanco.Caption, 3))
If nCodBanco = 0 Then
    nCodBanco = 90: sNomeBanco = "SN-OUTROS BANCOS"
ElseIf nCodBanco = 1 Then
    nCodBanco = 91: sNomeBanco = "SN-BANCO DO BRASIL"
ElseIf nCodBanco = 33 Then
    nCodBanco = 92: sNomeBanco = "SN-BANESPA"
ElseIf nCodBanco = 237 Then
    nCodBanco = 93: sNomeBanco = "SN-BRADESCO"
ElseIf nCodBanco = 341 Then
    nCodBanco = 94: sNomeBanco = "SN-ITAU"
ElseIf nCodBanco = 409 Then
    nCodBanco = 95: sNomeBanco = "SN-UNIBANCO"
ElseIf nCodBanco = 151 Then
    nCodBanco = 96: sNomeBanco = "SN-NOSSA CAIXA"
ElseIf nCodBanco = 104 Then
    nCodBanco = 97: sNomeBanco = "SN-CAIXA FEDERAL"
ElseIf nCodBanco = 399 Then
    nCodBanco = 98: sNomeBanco = "SN-HSBC AMRO BANK"
ElseIf nCodBanco > 90 And nCodBanco < 99 Then
    GoTo HERE
Else
    nCodBanco = 91: sNomeBanco = "SN-BANCO DO BRASIL"
End If
lblBanco.Caption = Format(nCodBanco, "000") & " - " & sNomeBanco
HERE:

Open sFullPath For Binary Access Read Write As FF1
    While Not EOF(FF1)
        On Error GoTo CloseFile2
        bExec = False
        If Left(sReg, 1) = "9" Then GoTo CloseFile2
        Input #FF1, sReg
        If Left(sReg, 1) = "1" Then
            sSeqArq = Mid(sReg, 2, 8)
        ElseIf Left(sReg, 1) = "2" Then
           'LE OS REGISTROS
            With grdReg
                nValorPrincipal = CDbl(Mid(sReg, 107, 17)) / 100
                nValorJuros = CDbl(Mid(sReg, 124, 17)) / 100
                nValorMulta = CDbl(Mid(sReg, 141, 17)) / 100
                nValorGuia = nValorPrincipal + nValorJuros + nValorMulta
                sAno = Mid(sReg, 101, 4)
                sMes = Mid(sReg, 105, 2)
                sAgencia = Mid(sReg, 223, 4)
                sCNPJ = Mid(sReg, 75, 14)
                
                nSeq = 0
                bAchou = False
                For R = 1 To UBound(aRegistro)
                    If aRegistro(R).sCNPJ = Mid(sReg, 75, 14) Then
                        bAchou = True
                        nSeq = nSeq + 1
                    End If
                Next
                
                ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                aRegistro(UBound(aRegistro)).sDataVencto = ConvDataSerial(Mid(sReg, 18, 8))
                aRegistro(UBound(aRegistro)).sDataCred = ConvDataSerial(Mid(sReg, 10, 8))
                aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 10, 8))
                aRegistro(UBound(aRegistro)).nValorPago = nValorGuia
                aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 223, 4)
                aRegistro(UBound(aRegistro)).sCNPJ = Mid(sReg, 75, 14)
                aRegistro(UBound(aRegistro)).nAno = Val(Mid(sReg, 101, 4))
                aRegistro(UBound(aRegistro)).nMes = Val(Mid(sReg, 105, 2))
                aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
                aRegistro(UBound(aRegistro)).sSitRetorno = "CNPJ: " & Format(aRegistro(UBound(aRegistro)).sCNPJ, "0#\.###\.###/####-##")
                aRegistro(UBound(aRegistro)).bExiste = True
                aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                If Not bAchou Then
                    aRegistro(UBound(aRegistro)).nSeq = 0
                Else
                    aRegistro(UBound(aRegistro)).nSeq = nSeq
                End If
                
                'PROCURA SE O DEBITO JA FOI BAIXADO
                Sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE ARQUIVOBANCO='" & lstArq.Text & "' AND DATACREDITO='" & Format(ConvDataSerial(Mid(sReg, 10, 8)), "mm/dd/yyyy") & "' AND "
                Sql = Sql & "CNPJ='" & Mid(sReg, 75, 14) & "' AND ANO=" & Val(Mid(sReg, 101, 4)) & " AND MES=" & Val(Mid(sReg, 105, 2))
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        'CARREGA PARCELA GRAVADA
                        ReDim Preserve aDoc(UBound(aDoc) + 1)
                        aDoc(UBound(aDoc)).sCNPJ = aRegistro(UBound(aRegistro)).sCNPJ
                        aDoc(UBound(aDoc)).nCodReduz = !CODREDUZIDO
                        aDoc(UBound(aDoc)).nAno = !AnoExercicio
                        aDoc(UBound(aDoc)).nLanc = !CodLancamento
                        aDoc(UBound(aDoc)).nSeq = !SeqLancamento
                        aDoc(UBound(aDoc)).nParc = !NumParcela
                        aDoc(UBound(aDoc)).nCompl = !CODCOMPLEMENTO
                        aDoc(UBound(aDoc)).sDataVencto = aRegistro(UBound(aRegistro)).sDataVencto
                        aDoc(UBound(aDoc)).sSit = 2
                        aDoc(UBound(aDoc)).nValorPrincipal = nValorPrincipal
                        aDoc(UBound(aDoc)).nValorMulta = nValorMulta
                        aDoc(UBound(aDoc)).nValorJuros = nValorJuros
                        aDoc(UBound(aDoc)).nValorCorrecao = 0
                        aDoc(UBound(aDoc)).nValorTotal = nValorGuia
                        aDoc(UBound(aDoc)).nValorTarifa = 0
                        aDoc(UBound(aDoc)).nValorDif = 0
                        aDoc(UBound(aDoc)).nValorCompensado = nValorGuia
                        aDoc(UBound(aDoc)).sBx = "S"
                        aDoc(UBound(aDoc)).sDp = "N"
                        aDoc(UBound(aDoc)).bExiste = True
                        aDoc(UBound(aDoc)).nSeqReg = aRegistro(UBound(aRegistro)).nSeq
                    Else
                        'DEFINIR NOVA PARCELA
                        'BUSCA CÓDIGO
                        Sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE DATAENCERRAMENTO IS NULL and CONVERT(BIGINT, cnpj) = " & Val(aRegistro(UBound(aRegistro)).sCNPJ)
                        Sql = Sql & " OR CNPJ='" & Format(aRegistro(UBound(aRegistro)).sCNPJ, "00\.000\.000/0000-00") & "' AND DATAENCERRAMENTO IS NULL ORDER BY CODIGOMOB DESC"
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                nCodReduz = !codigomob
                                .Close
                            Else
                                .Close
                                Sql = "SELECT CODCIDADAO,CNPJ FROM CIDADAO WHERE CNPJ = '" & RetornaNumero(aRegistro(UBound(aRegistro)).sCNPJ) & "' OR "
                                Sql = Sql & "CNPJ='" & Format(aRegistro(UBound(aRegistro)).sCNPJ, "00\.000\.000/0000-00") & "' ORDER BY CODCIDADAO DESC"
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    If .RowCount > 0 Then
                                        nCodReduz = !CodCidadao
                                    Else
                                        'CNPJ NÃO LOCALIZADO
                                        aRegistro(UBound(aRegistro)).bExiste = False
                                        Sql = "SELECT * FROM SIMPLESCNPJ WHERE CNPJ='" & aRegistro(UBound(aRegistro)).sCNPJ & "' AND ANOCOMP=" & aRegistro(UBound(aRegistro)).nAno & " AND MESCOMP=" & aRegistro(UBound(aRegistro)).nMes
                                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                        If RdoAux3.RowCount = 0 Then
                                            Sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
                                            Sql = Sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(aRegistro(UBound(aRegistro)).sCNPJ) & "','" & lstArq.Text & "'," & Val(Left(lblBanco.Caption, 3)) & ",'" & Format(aRegistro(UBound(aRegistro)).sDataCred, "mm/dd/yyyy") & "','"
                                            Sql = Sql & Format(aRegistro(UBound(aRegistro)).sDataVencto, "mm/dd/yyyy") & "'," & aRegistro(UBound(aRegistro)).nAno & "," & aRegistro(UBound(aRegistro)).nMes & "," & Virg2Ponto(CStr(aRegistro(UBound(aRegistro)).nValorPago)) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & ",'"
                                            Sql = Sql & aRegistro(UBound(aRegistro)).sAgencia & "'," & 0 & ")"
                                            cn.Execute Sql, rdExecDirect
                                        End If
                                        RdoAux3.Close
                                        GoTo CONTSN
                                    End If
                                End With
                            End If
                           
                        End With
                                
                        'BUSCA LANCAMENTO
                         Sql = "SELECT debitoparcela.codreduzido,debitoparcela.anoexercicio, debitoparcela.codlancamento,DEBITOPARCELA.SEQLANCAMENTO,debitoparcela.numparcela,DEBITOPARCELA.CODCOMPLEMENTO,debitoparcela.datavencimento, debitoparcela.statuslanc, debitotributo.valortributo "
                         Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                         Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
                         Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.codlancamento = 5) AND (MONTH(debitoparcela.datavencimento) = " & Month(CDate(aRegistro(UBound(aRegistro)).sDataVencto)) & ") AND "
                         Sql = Sql & "(YEAR(debitoparcela.datavencimento) = " & Year(CDate(aRegistro(UBound(aRegistro)).sDataVencto)) & ") AND (debitotributo.codtributo = 13) and debitotributo.valortributo =" & Virg2Ponto(CStr(nValorGuia)) & " AND statuslanc<>6"
                         Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                         With RdoAux3
                            'EXISTE LANCAMENTO NESTE MÊS/ANO?
                             If .RowCount > 0 Then 'SIM
                                 nNumParc = !NumParcela 'CAPTURA A PARCELA
                                 bAchou = False
                                'TEM ALGUMA QUE NÃO ESTA PAGA?
                                 Do Until .EOF
                                     If !statuslanc = 3 Then
                                         bAchou = True
                                         Exit Do
                                     End If
                                    .MoveNext
                                 Loop
                                 
                                'SE ACHOU PEGA A PARCELA
                                 If bAchou Then
                                     nSeq = !SeqLancamento
                                     nCompl = !CODCOMPLEMENTO '---------------> PARCELA PRONTA PARA USO
                                 Else
                                    'SE NÃO ACHAR
                                    .MoveFirst
                                     nCompl = 0
                                    'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
                                     Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (MONTH(datavencimento) = " & Month(dDataVencto) & ") AND "
                                     Sql = Sql & "(YEAR(datavencimento) = " & Year(dDataVencto) & ")"
                                     Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                     With RdoAux4
                                         If IsNull(!maximo) Then
                                             nSeq = 0
                                         Else
                                             nSeq = !maximo + 1
                                         End If
                                        .Close
                                     End With
                                 End If
                             
                             Else
                                'NÃO ACHOU LANCAMENTOS NESTE MÊS/ANO
                                'AUMENTA O LANCAMENTO
                                 Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & Val(sAno) & ")"
                                 Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux4
                                     If IsNull(!maximo) Then
                                         nSeq = 1
                                     Else
                                         nSeq = !maximo + 1
                                     End If
                                    .Close
                                 End With
                                 'VERIFICA SE A SEQ JA NÃO EXISTE NA MATRIZ
                                 For R = 1 To UBound(aDoc)
                                    If aDoc(R).nCodReduz = nCodReduz And aDoc(R).nAno = Val(sAno) Then
                                        nSeq = aDoc(R).nSeq + 1
                                    End If
                                 Next
                                 
                                 nCompl = 0
                                 nNumParc = 1
                             End If
                             ReDim Preserve aDoc(UBound(aDoc) + 1)
                             aDoc(UBound(aDoc)).sCNPJ = aRegistro(UBound(aRegistro)).sCNPJ
                             aDoc(UBound(aDoc)).nCodReduz = nCodReduz
                             aDoc(UBound(aDoc)).nAno = Val(sAno)
                             aDoc(UBound(aDoc)).nLanc = 5
                             aDoc(UBound(aDoc)).nSeq = nSeq
                             aDoc(UBound(aDoc)).nParc = nNumParc
                             aDoc(UBound(aDoc)).nCompl = nCompl
                             aDoc(UBound(aDoc)).sDataVencto = aRegistro(UBound(aRegistro)).sDataVencto
                             aDoc(UBound(aDoc)).sSit = 3
                             aDoc(UBound(aDoc)).nValorPrincipal = nValorPrincipal
                             aDoc(UBound(aDoc)).nValorMulta = nValorMulta
                             aDoc(UBound(aDoc)).nValorJuros = nValorJuros
                             aDoc(UBound(aDoc)).nValorCorrecao = 0
                             aDoc(UBound(aDoc)).nValorTotal = nValorGuia
                             aDoc(UBound(aDoc)).nValorTarifa = 0
                             aDoc(UBound(aDoc)).nValorDif = 0
                             aDoc(UBound(aDoc)).nValorCompensado = nValorGuia
                             aDoc(UBound(aDoc)).sBx = ""
                             aDoc(UBound(aDoc)).sDp = ""
                             aDoc(UBound(aDoc)).bExiste = True
                             aDoc(UBound(aDoc)).nSeqReg = aRegistro(UBound(aRegistro)).nSeq
                            .Close
                         End With
CONTSN:
'**********************************
                    End If
                   .Close
                End With
                
            End With
        ElseIf Left(sReg, 1) = "9" Then
           'LE O RODAPÉ DO ARQUIVO
            lblNumReg.Caption = Format(Val(Mid(sReg, 10, 6)) - 2, "000000")
            lblValorTotal.Caption = FormatNumber(CDbl(Mid(sReg, 16, 17) / 100), 2)
        End If
        nPos = nPos + 1
    Wend
CloseFile2:
Close #FF1
PBar.value = 0
Liberado
nErro = 0

cmbDataCredito.Clear

For nPos = 1 To UBound(aRegistro)
    bData = False
    For nData = 0 To cmbDataCredito.ListCount - 1
        If aRegistro(nPos).sDataCred = cmbDataCredito.List(nData) Then
            bData = True
        End If
    Next
    If bData = False Then
        cmbDataCredito.AddItem aRegistro(nPos).sDataCred
    End If
'    With aRegistro(nPos)
'        If aRegistro(nPos).bExiste = True Then
'            grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & Chr(9) & Format(.nMes, "00") & "/" & CStr(.nAno) & Chr(9) & _
'            Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
'            IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq
'        ElseIf aRegistro(nPos).bExiste = False And lblAS.Caption = "S" Then
'            grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & Chr(9) & Format(.nMes, "00") & "/" & CStr(.nAno) & Chr(9) & _
'            Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
'            IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq
'            nErro = nErro + 1
'        Else
'            nErro = nErro + 1
'            grdReg.AddItem Format(.nNumDoc, "000000000")
'        End If
'        If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
'            nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
'        End If
'    End With
Next


lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)
For nPos = 1 To UBound(aDoc)
    If aDoc(nPos).bExiste = False Then
        nErro = nErro + 1
    End If
Next
cmbDataCredito.ListIndex = 0

'If nErro > 0 Then
'    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
'    lblErro.Visible = True: cmdErro.Visible = True
'End If

'If grdReg.Rows > 1 Then
'    lblDC.Caption = grdReg.TextMatrix(1, 3)
'Else
'    lblDC.Caption = "Sem Registros"
'End If


Exit Sub

'*********************************
'****** ARQUIVO DE COBRANÇA ******
'*********************************
LEARQCOBRANCA:
grdReg.TextMatrix(0, 2) = "Data Pagam."
lblAC.Caption = "S"
nValorGuia = 0
Open sFullPath For Binary Access Read Write As FF1
    If bLayoutNovo Then
        '*** LAYOUT NOVO ***
        While Not EOF(FF1)
            On Error GoTo CloseFile3
            If nPos Mod 25 = 0 Then CallPb nPos, nTot
            If Mid(sReg, 9, 1) = " " And (Mid(sReg, 1, 8) <> "03300000") Then GoTo CloseFile3
            bExec = False
            Input #FF1, sReg
            If Mid(sReg, 1, 8) = "03300000" Then
                sSeqArq = Mid(sReg, 158, 6)
                sAgencia = Mid(sReg, 33, 5)
            ElseIf Mid(sReg, 14, 1) = "T" Then
               'LE OS REGISTROS TIPO T
                With grdReg
                    nNumDoc = Val(Mid(sReg, 41, 13))
                    GoTo Test2
Reduz2:
                    nNumDoc = Val(Mid(sReg, 41, 12))
Test2:
                    Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux.RowCount > 0 Then
                        ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                        aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                        If nNumDoc = 14294407 Then MsgBox "teste"
                        aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                        aRegistro(UBound(aRegistro)).nValorTarifaBancaria = (CDbl(Mid(sReg, 194, 15)) / 100)
                       'LOGO EM SEGUIDA LE OS REGISTROS TIPO U
                        Input #FF1, sReg
                        aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
                        aRegistro(UBound(aRegistro)).sDataPag = Left$(Mid(sReg, 138, 8), 2) & "/" & Mid$(Mid(sReg, 138, 8), 3, 2) & "/" & Right$(Mid(sReg, 138, 8), 4)
                        aRegistro(UBound(aRegistro)).sDataCred = Left$(Mid(sReg, 146, 8), 2) & "/" & Mid$(Mid(sReg, 146, 8), 3, 2) & "/" & Right$(Mid(sReg, 146, 8), 4)
                        aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 78, 15) / 100)
                        aRegistro(UBound(aRegistro)).sAgencia = sAgencia
                        aRegistro(UBound(aRegistro)).sConta = Mid(sReg, 30, 7)
                        If Not IsNull(RdoAux!ValorTaxaDoc) Then
                            aRegistro(UBound(aRegistro)).nValorTarifa = RdoAux!ValorTaxaDoc
                        Else
                            aRegistro(UBound(aRegistro)).nValorTarifa = 0
                        End If
                        aRegistro(UBound(aRegistro)).sSitRetorno = "00-BAIXA NORMAL"
                        aRegistro(UBound(aRegistro)).bExiste = True
                        aRegistro(UBound(aRegistro)).bIsentoMJ = IIf(Val(SubNull(RdoAux!isentomj)) = 0, False, True)
                        aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPagCalc))
                        CarregaParcela nNumDoc, nSeqReg, Left$(Mid(sReg, 138, 8), 2) & "/" & Mid$(Mid(sReg, 138, 8), 3, 2) & "/" & Right$(Mid(sReg, 138, 8), 4)
                    Else
                        If Not bExec And (Val(Mid(sReg, 41, 5)) > 0 Or Val(Mid(sReg, 41, 6)) <> "000001") And Len(CStr(nNumDoc)) > 7 Then
                            bExec = True
                            GoTo Reduz2
                        Else
                            Input #FF1, sReg
                            ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                            aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                            aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                            aRegistro(UBound(aRegistro)).sDataDoc = "01/01/1900"
                            aRegistro(UBound(aRegistro)).sDataPag = Left$(Mid(sReg, 138, 8), 2) & "/" & Mid$(Mid(sReg, 138, 8), 3, 2) & "/" & Right$(Mid(sReg, 138, 8), 4)
                            aRegistro(UBound(aRegistro)).sDataCred = Left$(Mid(sReg, 146, 8), 2) & "/" & Mid$(Mid(sReg, 146, 8), 3, 2) & "/" & Right$(Mid(sReg, 146, 8), 4)
                            aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 78, 15) / 100)
                            aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 109, 4)
                            aRegistro(UBound(aRegistro)).sConta = Mid(sReg, 30, 7)
                            aRegistro(UBound(aRegistro)).nValorTarifa = 0
                            aRegistro(UBound(aRegistro)).sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                            aRegistro(UBound(aRegistro)).bExiste = False
                            aRegistro(UBound(aRegistro)).bIsentoMJ = False
                            aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPagCalc))
                        End If
                    End If
                    nSeqReg = nSeqReg + 1
                    RdoAux.Close
                    nValorGuia = nValorGuia + aRegistro(UBound(aRegistro)).nValorPago
                End With
            End If
            nPos = nPos + 1
        Wend
        lblNumReg.Caption = Format(nPos, "000000")
        lblValorTotal.Caption = FormatNumber(nValorGuia, 2)
    Else
        '*** LAYOUT ANTIGO ***
        While Not EOF(FF1)
            On Error GoTo CloseFile3
            If nPos Mod 10 = 0 Then CallPb nPos, nTot
            If Left(sReg, 1) = "9" Then GoTo CloseFile3
            Input #FF1, sReg
            If Left(sReg, 1) = "0" Then
                lblDC.Caption = ConvDataSerial(Mid(sReg, 67, 6))
            ElseIf Left(sReg, 1) = "1" Then
               'LE OS REGISTROS
                With grdReg
                    nNumDoc = Val(Mid(sReg, 16, 7))
                    GoTo Test3
Reduz3:
                    nNumDoc = Val(Mid(sReg, 16, 6))
Test3:
                    Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux.RowCount > 0 Then
                        ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                        aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                        aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                        aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
                        aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 26, 6))
                        aRegistro(UBound(aRegistro)).sDataCred = lblDC.Caption
                        aRegistro(UBound(aRegistro)).nValorPago = (CDbl(Mid(sReg, 74, 15)) / 100) + (CDbl(Mid(sReg, 89, 13)) / 100)
                        aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 13, 3)
                        aRegistro(UBound(aRegistro)).nValorTarifa = RdoAux!ValorTaxaDoc
                        aRegistro(UBound(aRegistro)).sSitRetorno = "00-BAIXA NORMAL"
                        aRegistro(UBound(aRegistro)).bExiste = True
                        aRegistro(UBound(aRegistro)).bIsentoMJ = IIf(Val(SubNull(RdoAux!isentomj)) = 0, False, True)
                        aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                        CarregaParcela nNumDoc, nSeqReg, lblDC.Caption
                    Else
                        If Not bExec And (Val(Mid(sReg, 16, 1)) > 0 Or Val(Mid(sReg, 16, 2)) <> "01") And Len(CStr(nNumDoc)) > 7 Then
                            bExec = True
                            GoTo Reduz3
                        Else
                            ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                            aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                            aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                            aRegistro(UBound(aRegistro)).sDataDoc = "01/01/1900"
                            aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 26, 6))
                            aRegistro(UBound(aRegistro)).sDataCred = lblDC.Caption
                            aRegistro(UBound(aRegistro)).nValorPago = (CDbl(Mid(sReg, 72, 17)) / 100) + (CDbl(Mid(sReg, 89, 13)) / 100)
                            aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 13, 3)
                            aRegistro(UBound(aRegistro)).nValorTarifa = 0
                            aRegistro(UBound(aRegistro)).sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                            aRegistro(UBound(aRegistro)).bExiste = False
                            aRegistro(UBound(aRegistro)).bIsentoMJ = False
                            aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                        End If
                    End If
                    nSeqReg = nSeqReg + 1
                    RdoAux.Close
                End With
            ElseIf Left(sReg, 1) = "9" Then
               'LE O RODAPÉ DO ARQUIVO
                lblNumReg.Caption = Format(Val(Mid(sReg, 11, 6)) - 2, "000000")
                lblValorTotal.Caption = FormatNumber(CDbl(Mid(sReg, 17, 14) / 100), 2)
            End If
            nPos = nPos + 1
        Wend
    End If
CloseFile3:
Close #FF1
nErro = 0
cmbDataCredito.Clear

For nPos = 1 To UBound(aRegistro)

    bData = False
    For nData = 0 To cmbDataCredito.ListCount - 1
        If aRegistro(nPos).sDataCred = cmbDataCredito.List(nData) Then
            bData = True
        End If
    Next
    If bData = False Then
        cmbDataCredito.AddItem aRegistro(nPos).sDataCred
    End If

'    With aRegistro(nPos)
'        If aRegistro(nPos).bExiste = False Then
'            nErro = nErro + 1
'        End If
'        If .sDataDoc = "" Then .sDataDoc = Format(Now, "dd/mm/yyyy")
'        grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & Format(CDate(.sDataDoc), "dd/mm/yyyy") & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
'        Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & sConta
'        If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
'            nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
'        End If
'    End With
Next
cmbDataCredito.ListIndex = 0

'lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)
'For nPos = 1 To UBound(aDoc)
'    If aDoc(nPos).bExiste = False Then
'        nErro = nErro + 1
'    End If
'Next
'If nErro > 0 Then
'    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
'    lblErro.Visible = True: cmdErro.Visible = True
'End If

'If bLayoutNovo Then
'    lblNumReg.Caption = Format(grdReg.Rows - 1, "000000")
'    lblValorTotal.Caption = FormatNumber(nValorGuia, 2)
'End If

'If grdReg.Rows > 1 Then
'    lblDC.Caption = grdReg.TextMatrix(1, 3)
'Else
'    lblDC.Caption = "Sem Registros"
'End If

PBar.value = 0
Liberado

Exit Sub

'*********************************
'****** ARQUIVO DE COBRANÇA BB******
'*********************************
LEARQCOBRANCABB:
grdReg.TextMatrix(0, 2) = "Data Pagam."
lblAC.Caption = "S"
nValorGuia = 0
ConectaEicon
Open sFullPath For Binary Access Read Write As FF1

'*** LAYOUT NOVO ***
While Not EOF(FF1)
    On Error GoTo CloseFile3bb
    If nPos Mod 15 = 0 Then CallPb nPos, nTot
    If Mid(sReg, 9, 1) = " " And (Mid(sReg, 1, 8) <> "00100000") Then GoTo CloseFile3bb
    bExec = False
    Input #FF1, sReg
    If Mid(sReg, 1, 8) = "00100000" Then
        sSeqArq = Mid(sReg, 158, 6)
        sDataGeracao = Left$(Mid(sReg, 144, 8), 2) & "/" & Mid$(Mid(sReg, 144, 8), 3, 2) & "/" & Right$(Mid(sReg, 144, 8), 4)
        sAgencia = Mid(sReg, 33, 5)
    ElseIf Mid(sReg, 14, 1) = "T" Then
       'LE OS REGISTROS TIPO T
        sLinhaT = sReg
        With grdReg
            If Val(Mid(sReg, 16, 2)) = 6 Or Val(Mid(sReg, 16, 2)) = 17 Then
                'nNumDoc = Val(Mid(sReg, 45, 9))
                
                
                If Mid(sReg, 45, 2) = "00" Then
                    nNumDoc = Val(Mid(sReg, 47, 8))
                ElseIf Mid(sReg, 45, 1) <> "0" Then
                    nNumDoc = Val(Mid(sReg, 45, 8))
                Else
                    nNumDoc = Val(Mid(sReg, 46, 8))
                End If
                'If nNumDoc = 2103745 Then MsgBox "teste"
               ' nNumDoc = Val(Mid(sReg, 106, 8))
               ' If nNumDoc = 0 Then
               '     nNumDoc = Val(Mid(sReg, 45, 9))
               ' End If
            Else
                GoTo proximoBB
            '    nNumDoc = Val(Mid(sReg, 45, 10))
            End If
'            If nNumDoc = 2058074 Then MsgBox "teste"
            If nNumDoc > 200000 And nNumDoc < 300000 Then GoTo DocEicon
            GoTo Test2bb
DocEicon:
           nNumDoc = Val(Mid(sReg, 45, 10))
            GoTo Test2bb
Reduz2bb:
            nNumDoc = Val(Mid(sReg, 45, 9))
            
Test2bb:
            Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
'                If nNumDoc = 2058074 Then MsgBox "teste"
                aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                aRegistro(UBound(aRegistro)).nValorTarifaBancaria = (CDbl(Mid(sReg, 194, 15)) / 100)
                 aRegistro(UBound(aRegistro)).sConta = Mid(sReg, 30, 7)
'                aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 78, 19) / 100)
               'LOGO EM SEGUIDA LE OS REGISTROS TIPO U
                Input #FF1, sReg
                sLinhaU = sReg
                aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
                aRegistro(UBound(aRegistro)).sDataPag = Left$(Mid(sReg, 138, 8), 2) & "/" & Mid$(Mid(sReg, 138, 8), 3, 2) & "/" & Right$(Mid(sReg, 138, 8), 4)
                aRegistro(UBound(aRegistro)).sDataCred = Left$(Mid(sReg, 146, 8), 2) & "/" & Mid$(Mid(sReg, 146, 8), 3, 2) & "/" & Right$(Mid(sReg, 146, 8), 4)
                aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 78, 15) / 100)
             '   aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 93, 15) / 100)
                aRegistro(UBound(aRegistro)).sAgencia = sAgencia
               
                If Not IsNull(RdoAux!ValorTaxaDoc) Then
                    aRegistro(UBound(aRegistro)).nValorTarifa = RdoAux!ValorTaxaDoc
                Else
                    aRegistro(UBound(aRegistro)).nValorTarifa = 0
                End If
                aRegistro(UBound(aRegistro)).sSitRetorno = "00-BAIXA NORMAL"
                aRegistro(UBound(aRegistro)).bExiste = True
                aRegistro(UBound(aRegistro)).bIsentoMJ = IIf(Val(SubNull(RdoAux!isentomj)) = 0, False, True)
                aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                CarregaParcela nNumDoc, nSeqReg, CStr(aRegistro(UBound(aRegistro)).sDataPag)
            
                If (nNumDoc > 2000000 And nNumDoc < 3000000) Or nNumDoc < 3000 Then
                    '***** GRAVA BAIXA NA GISS ***************
                    Sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
                    Sql = Sql & 2177 & "," & 1 & "," & Val(sSeqArq) & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(sDataGeracao, "mm/dd/yyyy") & "','"
                    Sql = Sql & lstArq.Text & "','" & Format(aRegistro(UBound(aRegistro)).sDataCred, "mm/dd/yyyy") & "')"
                    cnEicon.Execute Sql, rdExecDirect
                    
                    For x = 1 To UBound(aDoc)
                        If aDoc(x).nNumDoc = nNumDoc Then
                            Exit For
                        End If
                    Next
                    Sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
                    Sql = Sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & 1 & "," & Val(sSeqArq) & "," & nNumDoc & "," & nPos & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
                    Sql = Sql & Virg2Ponto(CStr(aDoc(x).nValorTotal)) & "," & Virg2Ponto(CStr(aRegistro(UBound(aRegistro)).nValorPago)) & ",'" & Format(aRegistro(UBound(aRegistro)).sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(aDoc(x).nValorJuros + aDoc(x).nValorMulta + aDoc(x).nValorCorrecao)) & ",'"
                    Sql = Sql & sLinhaT & "','" & sLinhaU & "')"
                    cnEicon.Execute Sql, rdExecDirect
                    
                    '*****************************************
                End If
            Else
                If Not bExec And (Val(Mid(sReg, 41, 5)) > 0 Or Val(Mid(sReg, 41, 6)) <> "000001") And Len(CStr(nNumDoc)) > 7 Then
                    bExec = True
                    GoTo Reduz2bb
                Else
                    Input #FF1, sReg
                    ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                    aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                    aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                    aRegistro(UBound(aRegistro)).sDataDoc = "01/01/1900"
                    aRegistro(UBound(aRegistro)).sDataPag = Left$(Mid(sReg, 138, 8), 2) & "/" & Mid$(Mid(sReg, 138, 8), 3, 2) & "/" & Right$(Mid(sReg, 138, 8), 4)
                    aRegistro(UBound(aRegistro)).sDataCred = Left$(Mid(sReg, 146, 8), 2) & "/" & Mid$(Mid(sReg, 146, 8), 3, 2) & "/" & Right$(Mid(sReg, 146, 8), 4)
                    aRegistro(UBound(aRegistro)).nValorPago = CDbl(Mid(sReg, 78, 15) / 100)
                    aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 109, 4)
                    aRegistro(UBound(aRegistro)).nValorTarifa = 0
                    aRegistro(UBound(aRegistro)).sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                    aRegistro(UBound(aRegistro)).bExiste = False
                    aRegistro(UBound(aRegistro)).bIsentoMJ = False
                    aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPagCalc))
                End If
            End If
            nSeqReg = nSeqReg + 1
            RdoAux.Close
            nValorGuia = nValorGuia + aRegistro(UBound(aRegistro)).nValorPago
        End With
    End If
proximoBB:
    nPos = nPos + 1
    
    DoEvents
Wend
cnEicon.Close
CloseFile3bb:

'********
'Sql = "truncate table t2"
'cn.Execute Sql, rdExecDirect

'For nPos = 1 To UBound(aRegistro)
'    Sql = "insert t2(numdoc,valordoc) values(" & aRegistro(nPos).nNumDoc & "," & Virg2Ponto(CStr(aRegistro(nPos).nValorPago)) & ")"
'    cn.Execute Sql, rdExecDirect
'Next

'********
Dim sDataDoc As String
lblNumReg.Caption = Format(nPos, "000000")
lblValorTotal.Caption = FormatNumber(nValorGuia, 2)
Close #FF1
nErro = 0

cmbDataCredito.Clear

For nPos = 1 To UBound(aRegistro)
'If nPos = 327 Then MsgBox "teste"
    bData = False
    For nData = 0 To cmbDataCredito.ListCount - 1
        If aRegistro(nPos).sDataCred = cmbDataCredito.List(nData) Then
            bData = True
        End If
    Next
    If bData = False Then
        cmbDataCredito.AddItem aRegistro(nPos).sDataCred
    End If
    
'    With aRegistro(nPos)
'        If aRegistro(nPos).bExiste = False Then
'            nErro = nErro + 1
'        End If
'        If .sDataDoc <> "" Then
'           sDataDoc = Format(CDate(.sDataDoc), "dd/mm/yyyy")
'        Else
'            sDataDoc = Format(CDate(.sDataPag))
'       End If
'        grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & sDataDoc & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
'        Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & .sConta
'        If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
'            nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
'        End If
'    End With
Next

cmbDataCredito.ListIndex = 0

'lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)
'For nPos = 1 To UBound(aDoc)
'    If aDoc(nPos).bExiste = False Then
'        nErro = nErro + 1
'    End If
'Next
'If nErro > 0 Then
'    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
'    lblErro.Visible = True: cmdErro.Visible = True
'End If

'If bLayoutNovo Then
'    lblNumReg.Caption = Format(grdReg.Rows - 1, "000000")
'    lblValorTotal.Caption = FormatNumber(nValorGuia, 2)
'End If

'If grdReg.Rows > 1 Then
'    lblDC.Caption = grdReg.TextMatrix(1, 3)
'Else
'    lblDC.Caption = "Sem Registros"
'End If

PBar.value = 0
Liberado

Exit Sub

'***************************************
'****** ARQUIVO DÉBITO AUTOMÁTICO ******
'***************************************
LEDEBAUT:

grdReg.TextMatrix(0, 2) = "Data Pagam."
Open sFullPath For Binary Access Read Write As FF1
    While Not EOF(FF1)
        If nPos Mod 10 = 0 Then CallPb nPos, nTot
        On Error GoTo CloseFile4
        If Left(sReg, 1) = "Z" Then GoTo CloseFile4
        bExec = False
        Input #FF1, sReg
        If Left(sReg, 1) = "A" Then
            sSeqArq = Mid(sReg, 74, 6)
        ElseIf Left(sReg, 1) = "X" Then
        ElseIf Left(sReg, 1) = "F" Then
           'LE OS REGISTROS
            With grdReg
                nCodReduz = Val(Mid(sReg, 11, 6))
                'nCodReduz = Val(Mid(sReg, 71, 6))
'                Sql = "select codigopref from debitoautomatico where codreduz=" & nCodReduz
'                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
 '               If RdoAux.RowCount > 0 Then
  '                  nCodReduz = Val(SubNull(RdoAux!codigopref))
   '             End If
    '            RdoAux.Close
'                If nCodReduz = 26967 Then MsgBox "teste"
                
                
                If sDataVencto = "" Then
                    sDataVencto = ConvDataSerial(Mid(sReg, 45, 8))
                Else
                    sDataVencto = sDataVencto
                End If
                If IsDate(lblDC.Caption) Then
                    sDataVencto = lblDC.Caption
                End If
                nRetorno = Val(Mid(sReg, 68, 2))
                
                Select Case nRetorno
                        Case "00"
                                sRetorno = Format(nRetorno, "00") & " - " & "Débito Efetuado"
                        Case "01"
                                sRetorno = Format(nRetorno, "00") & " - " & "Insuficiência de Fundos"
                        Case "02"
                                sRetorno = Format(nRetorno, "00") & " - " & "Conta Corrente não Cadastrada"
                        Case "04"
                                sRetorno = Format(nRetorno, "00") & " - " & "Outras Restrições"
                        Case "10"
                                sRetorno = Format(nRetorno, "00") & " - " & "Agência em Regime de Encerramento"
                        Case "12"
                                sRetorno = Format(nRetorno, "00") & " - " & "Valor Inválido"
                        Case "13"
                                sRetorno = Format(nRetorno, "00") & " - " & "Data de Lançamento inválida"
                        Case "14"
                                sRetorno = Format(nRetorno, "00") & " - " & "Agência Inválida"
                        Case "15"
                                sRetorno = Format(nRetorno, "00") & " - " & "DAC da conta corrente inválido"
                        Case "18"
                                sRetorno = Format(nRetorno, "00") & " - " & "Data do Débito anterior ao do processamento"
                        Case "30"
                                sRetorno = Format(nRetorno, "00") & " - " & "Sem contrato de débito automático"
                        Case "96"
                                sRetorno = Format(nRetorno, "00") & " - " & "Manutenção do Cadastro"
                        Case "97"
                                sRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - Não Encontrado"
                        Case "98"
                                sRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - não efetuado, fora de tempo habil"
                        Case "99"
                                sRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - cancelado conforme solicitado"
                        Case Else
                               sRetorno = Format(nRetorno, "00") & " - " & "Erro Indefinido"
                End Select
                
                Sql = "SELECT lancamento.descreduz, debitoparcela.statuslanc, situacaolancamento.descsituacao, debitoparcela.datavencimento, debitoparcela.datadebase,"
                Sql = Sql & "debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
                Sql = Sql & "debitoparcela.CODCOMPLEMENTO , parceladocumento.NumDocumento FROM lancamento INNER JOIN debitoparcela ON lancamento.codlancamento = debitoparcela.codlancamento INNER JOIN "
                Sql = Sql & "situacaolancamento ON debitoparcela.statuslanc = situacaolancamento.codsituacao INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND "
                Sql = Sql & "debitoparcela.anoexercicio = parceladocumento.anoexercicio AND debitoparcela.codlancamento = parceladocumento.codlancamento AND "
                Sql = Sql & "debitoparcela.seqlancamento = parceladocumento.seqlancamento AND debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.CODCOMPLEMENTO = parceladocumento.CODCOMPLEMENTO "
                'Sql = Sql & "WHERE (DEBITOPARCELA.SEQLANCAMENTO=0) AND (DEBITOPARCELA.STATUSLANC=3) AND (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.CODLANCAMENTO = 1) AND "
                Sql = Sql & "WHERE (DEBITOPARCELA.SEQLANCAMENTO=0) AND (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.CODLANCAMENTO = 1) AND "
                Sql = Sql & "(DEBITOPARCELA.NUMPARCELA > 0) AND (DEBITOPARCELA.DATAVENCIMENTO = '" & Format(sDataVencto, "mm/dd/yyyy") & "')"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount = 0 Then
                        MsgBox "Código não encontrado nos optantes de débito automático (" & nCodReduz & ")"
                        nNumDoc = 0
                    Else
                        nNumDoc = !NumDocumento
                    End If
                   .Close
                End With
                
                Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 And nNumDoc > 0 Then
                    ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                    aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                    aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                    aRegistro(UBound(aRegistro)).sDataDoc = Format(RdoAux!DATADOCUMENTO, "dd/mm/yyyy")
                    aRegistro(UBound(aRegistro)).sDataPag = sDataVencto
                    aRegistro(UBound(aRegistro)).sDataCred = sDataVencto
                    aRegistro(UBound(aRegistro)).nValorPago = (CDbl(Mid(sReg, 53, 15)) / 100)
                    aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 27, 4)
                    If IsNull(RdoAux!ValorTaxaDoc) Then
                        aRegistro(UBound(aRegistro)).nValorTarifa = 0
                    Else
                        aRegistro(UBound(aRegistro)).nValorTarifa = RdoAux!ValorTaxaDoc
                    End If
                    aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
                    aRegistro(UBound(aRegistro)).sSitRetorno = sRetorno
                    aRegistro(UBound(aRegistro)).bExiste = True
                    aRegistro(UBound(aRegistro)).bIsentoMJ = False
                    aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                    CarregaParcela nNumDoc, nSeqReg, sDataVencto
                Else
                    ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                    aRegistro(UBound(aRegistro)).nNumDoc = nNumDoc
                    aRegistro(UBound(aRegistro)).nSeq = nSeqReg
                    aRegistro(UBound(aRegistro)).sDataDoc = ""
                    aRegistro(UBound(aRegistro)).sDataPag = sDataVencto
                    aRegistro(UBound(aRegistro)).sDataCred = sDataVencto
                    aRegistro(UBound(aRegistro)).nValorPago = (CDbl(Mid(sReg, 53, 15)) / 100)
                    aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 27, 4)
                    aRegistro(UBound(aRegistro)).nValorTarifa = 0
                    aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
                    aRegistro(UBound(aRegistro)).sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                    aRegistro(UBound(aRegistro)).bExiste = False
                    aRegistro(UBound(aRegistro)).bIsentoMJ = False
                    aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                End If
                RdoAux.Close
            End With
            nSeqReg = nSeqReg + 1
        ElseIf Left(sReg, 1) = "Z" Then
           'LE O RODAPÉ DO ARQUIVO
            lblNumReg.Caption = Format(Val(Mid(sReg, 2, 6)) - 2, "000000")
            lblValorTotal.Caption = FormatNumber(CDbl(Mid(sReg, 8, 17) / 100), 2)
            lblDA.Caption = "S"
        End If
        nPos = nPos + 1
    Wend
CloseFile4:
Close #FF1
Liberado
PBar.value = 0

nErro = 0
cmbDataCredito.Clear


For nPos = 1 To UBound(aRegistro)
    
    bData = False
    For nData = 0 To cmbDataCredito.ListCount - 1
        If aRegistro(nPos).sDataCred = cmbDataCredito.List(nData) Then
            bData = True
        End If
    Next
    If bData = False Then
        cmbDataCredito.AddItem aRegistro(nPos).sDataCred
    End If
    
    
'    With aRegistro(nPos)
'        If aRegistro(nPos).bExiste = True Then
'            grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & .sDataDoc & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
'            Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
'            IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 0) & Chr(9) & .nSeq & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy")
'        Else
'            nErro = nErro + 1
'            grdReg.AddItem Format(.nNumDoc, "000000000")
'        End If
'        If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
'            nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
 '       End If
 '   End With
Next
If cmbDataCredito.ListCount > 0 Then
    cmbDataCredito.ListIndex = 0
End If
'lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)

'For nPos = 1 To UBound(aDoc)
'    If aDoc(nPos).bExiste = False Then
'        nErro = nErro + 1
'    End If
'Next

'If nErro > 0 Then
'    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
'    lblErro.Visible = True: cmdErro.Visible = True
'End If

'If grdReg.Rows > 1 Then
'    lblDC.Caption = grdReg.TextMatrix(1, 3)
'Else
'    lblDC.Caption = "Sem Registros"
'End If

'Exit Sub

End Sub

Private Sub CarregaTributo()
Dim Sql As String, RdoAux As rdoResultset, nLast As Integer

ReDim aTribF(0)
Sql = "SELECT * FROM TRIBUTO ORDER BY CODTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aTribF(UBound(aTribF) + 1)
        nLast = UBound(aTribF)
        aTribF(nLast).nCodTrib = !CodTributo
        aTribF(nLast).sAbrevTrib = !ABREVTRIBUTO
        aTribF(nLast).Ficha = !Ficha
        aTribF(nLast).FichaJrMulta = !FichaJrMulta
        aTribF(nLast).FichaDivida = !FichaDivida
        aTribF(nLast).FichaDaJrMul = !FichaDaJrMul
        aTribF(nLast).FichaDaEnca = !FichaDaEnca
        aTribF(nLast).FichaAjuiza = !FichaAjuiza
        aTribF(nLast).FichaAjJrMul = !FichaAjJrMul
        aTribF(nLast).FichaAjEnca = !FichaAjEnca
       .MoveNext
    Loop
   .Close
End With

ReDim aFicha(1)
aFicha(1).Ficha = 0
aFicha(1).Natureza = "Não Cadastrado"
aFicha(1).Desc = "Não Cadastrado"
aFicha(1).Vinculo = "Não Cadastrado"
aFicha(1).Perc = 0
Sql = "SELECT * FROM FICHACONTABIL ORDER BY FICHA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aFicha(UBound(aFicha) + 1)
        nLast = UBound(aFicha)
        aFicha(nLast).Ficha = !Ficha
        aFicha(nLast).Natureza = !Natureza
        aFicha(nLast).Desc = SubNull(!descta)
        aFicha(nLast).Vinculo = !Vinculo
        aFicha(nLast).Perc = !Perc
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub GravaAnalise()
Dim x As Integer, Sql As String, Y As Integer, z As Integer
 
For x = 1 To UBound(aTrib)
    With aTrib(x)
        For Y = 1 To UBound(aTribF)
            If aTribF(Y).nCodTrib = .nCodTrib Then Exit For
        Next
        
        If .nFicha = 0 And .nValorTotal > 0 Then
            .nFicha = 50416
        End If
        If .nFichaJM = 0 And (.nValorMulta + .nValorJuros) > 0 Then
            .nFichaJM = 50416
        End If
        If .nFichaC = 0 And .nValorCorrecao > 0 Then
            .nFichaC = 50416
        End If
        
        For z = 1 To UBound(aFicha)
            If aFicha(z).Ficha = .nFicha Then Exit For
        Next
        If .nFicha > 0 Then
            Sql = "INSERT ANALISE2(USUARIO,DATARECEITA,CODBANCO,ARQUIVO,NUMDOCUMENTO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
            Sql = Sql & "CODTRIBUTO,DESCTRIBUTO,VALORTOTAL,NUMFICHA,DESCFICHA,NATUREZA,VINCULO,PERC) "
            Sql = Sql & "VALUES('" & NomeDeLogin & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "'," & Val(Left(lblBanco.Caption, 3)) & ",'" & lstArq.Text & "',"
            Sql = Sql & .nNumDoc & "," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & "," & .nCodTrib & ",'" & aTribF(Y).sAbrevTrib & "',"
            'Sql = Sql & Virg2Ponto(Format(.nValorpTotal - .nValorMulta - .nValorJuros - .nValorCorrecao, "#0.00")) & "," & .nFicha & ",'" & aFicha(z).Desc & "','" & aFicha(z).Natureza & "','" & aFicha(z).Vinculo & "'," & aFicha(z).Perc & ")"
            Sql = Sql & Virg2Ponto(Format(.nValorPrincipal + .nValorTarifa, "#0.00")) & "," & .nFicha & ",'" & aFicha(z).Desc & "','" & aFicha(z).Natureza & "','" & aFicha(z).Vinculo & "'," & aFicha(z).Perc & ")"
            cn.Execute Sql, rdExecDirect
        End If
        For z = 1 To UBound(aFicha)
            If aFicha(z).Ficha = .nFichaJM Then Exit For
        Next
        If .nFichaJM > 0 Then
            Sql = "INSERT ANALISE2(USUARIO,DATARECEITA,CODBANCO,ARQUIVO,NUMDOCUMENTO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
            Sql = Sql & "CODTRIBUTO,DESCTRIBUTO,VALORTOTAL,NUMFICHA,DESCFICHA,NATUREZA,VINCULO,PERC) "
            Sql = Sql & "VALUES('" & NomeDeLogin & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "'," & Val(Left(lblBanco.Caption, 3)) & ",'" & lstArq.Text & "',"
            Sql = Sql & .nNumDoc & "," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & "," & .nCodTrib & ",'" & aTribF(Y).sAbrevTrib & "',"
            Sql = Sql & Virg2Ponto(Format(.nValorJuros + .nValorMulta, "#0.00")) & "," & .nFichaJM & ",'" & aFicha(z).Desc & "','" & aFicha(z).Natureza & "','" & aFicha(z).Vinculo & "'," & aFicha(z).Perc & ")"
            cn.Execute Sql, rdExecDirect
        End If
        
        If .nFichaC > 0 Then
            For z = 1 To UBound(aFicha)
                If aFicha(z).Ficha = .nFichaC Then Exit For
            Next
            Sql = "INSERT ANALISE2(USUARIO,DATARECEITA,CODBANCO,ARQUIVO,NUMDOCUMENTO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
            Sql = Sql & "CODTRIBUTO,DESCTRIBUTO,VALORTOTAL,NUMFICHA,DESCFICHA,NATUREZA,VINCULO,PERC) "
            Sql = Sql & "VALUES('" & NomeDeLogin & "','" & Format(lblDC.Caption, "mm/dd/yyyy") & "'," & Val(Left(lblBanco.Caption, 3)) & ",'" & lstArq.Text & "',"
            Sql = Sql & .nNumDoc & "," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & "," & .nCodTrib & ",'" & aTribF(Y).sAbrevTrib & "',"
            Sql = Sql & Virg2Ponto(Format(.nValorCorrecao, "#0.00")) & "," & .nFichaC & ",'" & aFicha(z).Desc & "','" & aFicha(z).Natureza & "','" & aFicha(z).Vinculo & "'," & aFicha(z).Perc & ")"
            cn.Execute Sql, rdExecDirect
        End If
    End With
Next

End Sub

Private Sub MontaMenu()

   Set m_cMenu = New cPopupMenu
   With m_cMenu
      .hwndOwner = Me.HWND
      .GradientHighlight = True
      
      i = .AddItem("Visualizar arquivo texto", "", 1, , , , , "mnuVisualizar")
      .OwnerDraw(i) = True
      i = .AddItem("Efetuar baixa no arquivo", "", 1, , , , , "mnuBaixa")
      .OwnerDraw(i) = True
      i = .AddItem("Reativar os lançamentos", "", 1, , , , , "mnuReativar")
      .OwnerDraw(i) = True
      i = .AddItem("Baixa manual", "", 1, , , , , "mnuBaixaManual")
      .OwnerDraw(i) = True
      i = .AddItem("Arquivo CBR724", "", 1, , , , , "mnuCBR724")
      .OwnerDraw(i) = True
      i = .AddItem("Corrigir Documento", "", 1, , , , , "mnuFixDoc")
      .OwnerDraw(i) = True
      i = .AddItem("Outras informações", "", 1, , , , , "mnuOutro")
      .OwnerDraw(i) = True
      i = .AddItem("Relatórios", "", 1, , , , , "mnuRelatório")
      .OwnerDraw(i) = True
      .AddItem "Resumo do Arquivo", "", 1, i, , , , "mnuResumo"
      h = .AddItem("Gerar análise do arquivo", "", 1, i, , , , "mnuAnalise")
      .AddItem "Análise Resumida", "", 1, h, , , , "mnuAnaliseR"
      .AddItem "Análise Detalhada", "", 1, h, , , , "mnuAnaliseD"
   End With

End Sub

Private Sub CarregaDataCredito()
Dim nValorEfetivo As Double
grdReg.Rows = 1
For nPos = 1 To UBound(aRegistro)
    With aRegistro(nPos)
        If aRegistro(nPos).sDataCred = cmbDataCredito.Text Then
            If aRegistro(nPos).bExiste = True Then
                grdReg.AddItem Format(.nNumDoc, "000000000") & Chr(9) & .sDataDoc & Chr(9) & Format(CDate(.sDataPag), "dd/mm/yyyy") & Chr(9) & _
                Format(CDate(.sDataCred), "dd/mm/yyyy") & Chr(9) & FormatNumber(.nValorPago, 2) & Chr(9) & .sAgencia & Chr(9) & FormatNumber(.nValorTarifa, 2) & Chr(9) & _
                IIf(.bExiste, "N", "S") & Chr(9) & IIf(.bIsentoMJ, "S", "N") & Chr(9) & .sSitRetorno & Chr(9) & FormatNumber(.nValorTarifaBancaria, 2) & Chr(9) & .nSeq & Chr(9) & Format(CDate(.sDataPagCalc), "dd/mm/yyyy")
            Else
                nErro = nErro + 1
                grdReg.AddItem Format(.nNumDoc, "000000000")
            End If
            If Val(Left(aRegistro(nPos).sSitRetorno, 2)) = 0 Then
                nValorEfetivo = nValorEfetivo + CDbl(aRegistro(nPos).nValorPago)
            End If
        End If
    End With
Next

lblValorEfetivo.Caption = FormatNumber(nValorEfetivo, 2)
For nPos = 1 To UBound(aDoc)
    If aDoc(nPos).bExiste = False Then
        nErro = nErro + 1
    End If
Next

If nErro > 0 Then
    lblErro.Caption = nErro & " ERRO(S) ENCONTRADO(S)"
    lblErro.Visible = True: cmdErro.Visible = True
End If

If grdReg.Rows > 1 Then
    For x = 1 To grdReg.Rows - 1
        If grdReg.TextMatrix(x, 3) <> "" Then
            lblDC.Caption = grdReg.TextMatrix(x, 3)
            Exit For
        End If
    Next
    
Else
    lblDC.Caption = "Sem Registros"
End If
lblNumReg.Caption = Format(grdReg.Rows - 1, "000000")

If IsDate(lblDC.Caption) Then
    Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & lstArq.Text & "' AND DATACREDITO='" & Format(lblDC.Caption, "mm/dd/yyyy") & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If IsNull(!DATABAIXA) Then
                lblDB.Caption = "Sem Baixa"
            Else
                lblDB.Caption = Format(!DATABAIXA, "dd/mm/yyyy")
            End If
        Else
            lblDB.Caption = "Sem Baixa"
        End If
       .Close
    End With
Else
    lblDB.Caption = "Sem Baixa"
End If


End Sub
