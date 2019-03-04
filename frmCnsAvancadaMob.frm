VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmCnsAvancadaMob 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta avançada de empresas e gerador de correspondência"
   ClientHeight    =   6435
   ClientLeft      =   14190
   ClientTop       =   4815
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   11400
   Begin RichTextLib.RichTextBox RtbTmp 
      Height          =   915
      Left            =   45
      TabIndex        =   76
      Top             =   6885
      Visible         =   0   'False
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   1614
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"frmCnsAvancadaMob.frx":0000
   End
   Begin Tributacao.jcFrames frModelo 
      Height          =   3570
      Left            =   7470
      Top             =   1320
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   6297
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Modelos Disponíveis"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   4
      ColorFrom       =   12648384
      ColorTo         =   8454016
      Begin VB.ListBox lstModelo 
         Appearance      =   0  'Flat
         Height          =   2760
         Left            =   0
         TabIndex        =   70
         Top             =   360
         Width           =   3480
      End
      Begin prjChameleon.chameleonButton cmdSelectModel 
         Height          =   315
         Left            =   495
         TabIndex        =   71
         ToolTipText     =   "Voltar a tela anterior"
         Top             =   3195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Selecionar"
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
         MICON           =   "frmCnsAvancadaMob.frx":0082
         PICN            =   "frmCnsAvancadaMob.frx":009E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCloseModel 
         Height          =   315
         Left            =   1845
         TabIndex        =   72
         ToolTipText     =   "Retorna Cidadão Selecionado"
         Top             =   3195
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Fechar"
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
         MICON           =   "frmCnsAvancadaMob.frx":02B2
         PICN            =   "frmCnsAvancadaMob.frx":02CE
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
   Begin Tributacao.jcFrames tb2 
      Height          =   465
      Left            =   45
      Top             =   5940
      Visible         =   0   'False
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   820
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
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
      Begin VB.TextBox txtSep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3015
         MaxLength       =   1
         TabIndex        =   33
         Text            =   ","
         Top             =   105
         Width           =   510
      End
      Begin prjChameleon.chameleonButton cmdVoltar 
         Height          =   315
         Left            =   9990
         TabIndex        =   34
         ToolTipText     =   "Voltar a tela anterior"
         Top             =   90
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
         MICON           =   "frmCnsAvancadaMob.frx":033C
         PICN            =   "frmCnsAvancadaMob.frx":0358
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
         Left            =   135
         TabIndex        =   107
         ToolTipText     =   "Opções de impressão"
         Top             =   90
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Imprimir para ..."
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
         MICON           =   "frmCnsAvancadaMob.frx":04B2
         PICN            =   "frmCnsAvancadaMob.frx":04CE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDebitos 
         Height          =   315
         Left            =   3990
         TabIndex        =   126
         ToolTipText     =   "Voltar a tela anterior"
         Top             =   90
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Débitos"
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
         MCOL            =   16711935
         MPTR            =   1
         MICON           =   "frmCnsAvancadaMob.frx":0628
         PICN            =   "frmCnsAvancadaMob.frx":0644
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Separador:"
         Height          =   195
         Index           =   0
         Left            =   2160
         TabIndex        =   30
         Top             =   135
         Width           =   825
      End
   End
   Begin Tributacao.jcFrames tb3 
      Height          =   465
      Left            =   45
      Top             =   5940
      Visible         =   0   'False
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   820
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
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
      Begin VB.TextBox txtAutoInc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6075
         MaxLength       =   5
         TabIndex        =   130
         Text            =   "1"
         Top             =   90
         Width           =   870
      End
      Begin prjChameleon.chameleonButton cmdImprimir 
         Height          =   315
         Left            =   2700
         TabIndex        =   69
         ToolTipText     =   "Imprimir as cartas"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsAvancadaMob.frx":0704
         PICN            =   "frmCnsAvancadaMob.frx":0720
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdVoltar2 
         Height          =   315
         Left            =   9990
         TabIndex        =   55
         ToolTipText     =   "Voltar a tela anterior"
         Top             =   90
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
         MICON           =   "frmCnsAvancadaMob.frx":087A
         PICN            =   "frmCnsAvancadaMob.frx":0896
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCarregar 
         Height          =   315
         Left            =   90
         TabIndex        =   67
         ToolTipText     =   "Carregar um modelo gravado"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Carregar"
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
         MICON           =   "frmCnsAvancadaMob.frx":09F0
         PICN            =   "frmCnsAvancadaMob.frx":0A0C
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
         Left            =   1395
         TabIndex        =   68
         ToolTipText     =   "Gravar ou atualizar o modelo"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Gravar"
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
         MICON           =   "frmCnsAvancadaMob.frx":0A93
         PICN            =   "frmCnsAvancadaMob.frx":0AAF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Incrementar a partir de:"
         Height          =   195
         Index           =   1
         Left            =   4365
         TabIndex        =   131
         Top             =   135
         Width           =   1680
      End
   End
   Begin Tributacao.jcFrames fr1 
      Height          =   5865
      Left            =   45
      Top             =   45
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   10345
      FrameColor      =   8388608
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Seleção de Critérios"
      TextBoxHeight   =   18
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
      Begin VB.Frame frDDList4 
         BackColor       =   &H00EEEEEE&
         Height          =   375
         Left            =   4770
         TabIndex        =   119
         Top             =   5340
         Width           =   6150
         Begin VB.ListBox lstDDList4 
            Height          =   2085
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   120
            Top             =   405
            Width           =   6060
         End
         Begin prjChameleon.chameleonButton cmdDDList4 
            Height          =   240
            Left            =   1635
            TabIndex        =   121
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   423
            BTYPE           =   14
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":0E54
            PICN            =   "frmCnsAvancadaMob.frx":0E70
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDDList4All 
            Height          =   240
            Left            =   5265
            TabIndex        =   122
            ToolTipText     =   "Selecionar todos"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":0FCA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDDList4None 
            Height          =   240
            Left            =   5625
            TabIndex        =   123
            ToolTipText     =   "Manter apenas o código"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "-"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":0FE6
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
      Begin VB.Frame frIssAtiv 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Atividade ISS"
         ForeColor       =   &H00800000&
         Height          =   945
         Left            =   4680
         TabIndex        =   117
         Top             =   4860
         Width           =   6360
         Begin VB.OptionButton optAtiv 
            Caption         =   "Descrição"
            Height          =   195
            Index           =   1
            Left            =   4050
            TabIndex        =   125
            Top             =   180
            Value           =   -1  'True
            Width           =   1245
         End
         Begin VB.OptionButton optAtiv 
            Caption         =   "Código"
            Height          =   195
            Index           =   0
            Left            =   3090
            TabIndex        =   124
            Top             =   180
            Width           =   885
         End
         Begin VB.CheckBox chkAtivIss 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Indiferente"
            Height          =   195
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Value           =   1  'Checked
            Width           =   1140
         End
      End
      Begin VB.Frame Frame10 
         Height          =   555
         Left            =   4680
         TabIndex        =   114
         Top             =   315
         Width           =   6360
         Begin VB.OptionButton optIM 
            Caption         =   "Empresas sem Inscrição Muncipal"
            Height          =   240
            Index           =   1
            Left            =   3150
            TabIndex        =   116
            Top             =   225
            Width           =   2850
         End
         Begin VB.OptionButton optIM 
            Caption         =   "Empresas com Inscrição Muncipal"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   115
            Top             =   225
            Value           =   -1  'True
            Width           =   2850
         End
      End
      Begin VB.Frame frDDList3 
         BackColor       =   &H00EEEEEE&
         Height          =   375
         Left            =   6120
         TabIndex        =   99
         Top             =   4050
         Width           =   4830
         Begin VB.ListBox lstDDList3 
            Appearance      =   0  'Flat
            Height          =   1380
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   108
            Top             =   405
            Width           =   4740
         End
         Begin prjChameleon.chameleonButton cmdDDList3 
            Height          =   240
            Left            =   315
            TabIndex        =   100
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   423
            BTYPE           =   14
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1002
            PICN            =   "frmCnsAvancadaMob.frx":101E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDDList3All 
            Height          =   240
            Left            =   3915
            TabIndex        =   101
            ToolTipText     =   "Selecionar todos"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1178
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDDList3None 
            Height          =   240
            Left            =   4275
            TabIndex        =   102
            ToolTipText     =   "Manter apenas o código"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "-"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1194
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
      Begin VB.Frame Frame9 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Pagamentos"
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   4620
         TabIndex        =   105
         Top             =   360
         Visible         =   0   'False
         Width           =   6360
         Begin VB.TextBox txtAno 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5175
            MaxLength       =   4
            TabIndex        =   25
            Top             =   270
            Width           =   750
         End
         Begin VB.ComboBox cmbImposto 
            Height          =   315
            ItemData        =   "frmCnsAvancadaMob.frx":11B0
            Left            =   2475
            List            =   "frmCnsAvancadaMob.frx":11B2
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   270
            Width           =   2130
         End
         Begin VB.ComboBox cmbPagto 
            Height          =   315
            ItemData        =   "frmCnsAvancadaMob.frx":11B4
            Left            =   90
            List            =   "frmCnsAvancadaMob.frx":11C7
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   270
            Width           =   2265
         End
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Ano..:"
            Height          =   195
            Index           =   6
            Left            =   4680
            TabIndex        =   106
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame frDDList2 
         BackColor       =   &H00EEEEEE&
         Height          =   375
         Left            =   6120
         TabIndex        =   94
         Top             =   3330
         Width           =   4830
         Begin VB.ListBox lstDDList2 
            Height          =   2085
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   95
            Top             =   405
            Width           =   4740
         End
         Begin prjChameleon.chameleonButton cmdDDList2 
            Height          =   240
            Left            =   315
            TabIndex        =   96
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   423
            BTYPE           =   14
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1238
            PICN            =   "frmCnsAvancadaMob.frx":1254
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDD2All 
            Height          =   240
            Left            =   3915
            TabIndex        =   97
            ToolTipText     =   "Selecionar todos"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":13AE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDD2None 
            Height          =   240
            Left            =   4275
            TabIndex        =   98
            ToolTipText     =   "Manter apenas o código"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "-"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":13CA
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
      Begin VB.Frame Frame8 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Importação de Arquivos"
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   315
         TabIndex        =   77
         Top             =   4365
         Width           =   4155
         Begin VB.TextBox txtDelimiter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Left            =   1035
            MaxLength       =   1
            TabIndex        =   80
            Text            =   ","
            Top             =   675
            Width           =   330
         End
         Begin VB.TextBox txtArq 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Height          =   285
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   315
            Width           =   3390
         End
         Begin prjChameleon.chameleonButton cmdOpen 
            Height          =   315
            Left            =   3555
            TabIndex        =   79
            ToolTipText     =   "Localizar arquivo texto"
            Top             =   315
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":13E6
            PICN            =   "frmCnsAvancadaMob.frx":1402
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdImportar 
            Height          =   315
            Left            =   1395
            TabIndex        =   81
            ToolTipText     =   "Importar o arquivo selecionado"
            Top             =   675
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "Importar"
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
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1489
            PICN            =   "frmCnsAvancadaMob.frx":14A5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdLimpar 
            Height          =   315
            Left            =   3555
            TabIndex        =   83
            ToolTipText     =   "Limpar texto"
            Top             =   675
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":16B8
            PICN            =   "frmCnsAvancadaMob.frx":16D4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdPreview 
            Height          =   315
            Left            =   3105
            TabIndex        =   82
            ToolTipText     =   "Visualizar arquivo"
            Top             =   675
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":18F1
            PICN            =   "frmCnsAvancadaMob.frx":190D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblTotImp 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3285
            TabIndex        =   86
            Top             =   1125
            Width           =   420
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Empresas localizadas no arquivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   85
            Top             =   1125
            Width           =   2895
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Delimitador.:"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   84
            Top             =   735
            Width           =   870
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Critérios Diversos"
         ForeColor       =   &H00800000&
         Height          =   4020
         Left            =   315
         TabIndex        =   46
         Top             =   315
         Width           =   4155
         Begin VB.ComboBox cmbDispensaIE 
            Height          =   315
            ItemData        =   "frmCnsAvancadaMob.frx":1A67
            Left            =   2610
            List            =   "frmCnsAvancadaMob.frx":1A74
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   3555
            Width           =   1275
         End
         Begin VB.ComboBox cmbISSE 
            Height          =   315
            ItemData        =   "frmCnsAvancadaMob.frx":1A89
            Left            =   2610
            List            =   "frmCnsAvancadaMob.frx":1A96
            Style           =   2  'Dropdown List
            TabIndex        =   109
            Top             =   3195
            Width           =   1275
         End
         Begin VB.ComboBox cmbISS 
            Height          =   315
            ItemData        =   "frmCnsAvancadaMob.frx":1AAB
            Left            =   2610
            List            =   "frmCnsAvancadaMob.frx":1AAD
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2835
            Width           =   1275
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1755
            Width           =   1275
         End
         Begin VB.ComboBox cmbVSanit 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   2475
            Width           =   1275
         End
         Begin VB.ComboBox cmbMEI 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   315
            Width           =   1275
         End
         Begin VB.ComboBox cmbSimples 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   675
            Width           =   1275
         End
         Begin VB.ComboBox cmbIsento 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1035
            Width           =   1275
         End
         Begin VB.ComboBox cmbAlvara 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1395
            Width           =   1275
         End
         Begin VB.ComboBox cmbVistoria 
            Height          =   315
            Left            =   2610
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2115
            Width           =   1275
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Dispensa ISS Eletron.:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   8
            Left            =   225
            TabIndex        =   112
            Top             =   3600
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "ISS Eletrônico.......:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   225
            TabIndex        =   110
            Top             =   3240
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Regime.......:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   225
            TabIndex        =   54
            Top             =   2880
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Vigilância Sanitária.:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   225
            TabIndex        =   53
            Top             =   2520
            Width           =   2310
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Integrante do MEI....:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   52
            Top             =   360
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Simples Nacional.....:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   51
            Top             =   720
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Isento de Taxa.......:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   225
            TabIndex        =   50
            Top             =   1080
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Alvará Automático....:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   225
            TabIndex        =   49
            Top             =   1440
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo de Empresa......:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   225
            TabIndex        =   48
            Top             =   1800
            Width           =   2310
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Possui Vistoria......:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   225
            TabIndex        =   47
            Top             =   2160
            Width           =   2310
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Atividade Principal"
         ForeColor       =   &H00800000&
         Height          =   1005
         Left            =   4680
         TabIndex        =   45
         Top             =   3825
         Width           =   6360
         Begin VB.CheckBox chkTipoAtiv 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Outros"
            Height          =   195
            Index           =   4
            Left            =   3555
            TabIndex        =   22
            Top             =   675
            Value           =   1  'Checked
            Width           =   1140
         End
         Begin VB.CheckBox chkTipoAtiv 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Serviços"
            Height          =   195
            Index           =   3
            Left            =   2430
            TabIndex        =   21
            Top             =   675
            Value           =   1  'Checked
            Width           =   1140
         End
         Begin VB.CheckBox chkTipoAtiv 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Comércio"
            Height          =   195
            Index           =   2
            Left            =   1305
            TabIndex        =   20
            Top             =   675
            Value           =   1  'Checked
            Width           =   1140
         End
         Begin VB.CheckBox chkTipoAtiv 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Industria"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   19
            Top             =   675
            Value           =   1  'Checked
            Width           =   1140
         End
         Begin VB.CheckBox chkAtividade 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Indiferente"
            Height          =   195
            Left            =   225
            TabIndex        =   18
            Top             =   315
            Value           =   1  'Checked
            Width           =   1140
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Contador"
         ForeColor       =   &H00800000&
         Height          =   645
         Index           =   0
         Left            =   4680
         TabIndex        =   44
         Top             =   3150
         Width           =   6360
         Begin VB.CheckBox chkContador 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Indiferente"
            Height          =   195
            Left            =   225
            TabIndex        =   17
            Top             =   300
            Value           =   1  'Checked
            Width           =   1140
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Data de Suspensão"
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   4680
         TabIndex        =   41
         Top             =   2430
         Width           =   6360
         Begin VB.ComboBox cmbDSus 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   270
            Width           =   1455
         End
         Begin esMaskEdit.esMaskedEdit mskDataSusIni 
            Height          =   285
            Left            =   2925
            TabIndex        =   15
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MouseIcon       =   "frmCnsAvancadaMob.frx":1AAF
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
         Begin esMaskEdit.esMaskedEdit mskDataSusFim 
            Height          =   285
            Left            =   4905
            TabIndex        =   16
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MouseIcon       =   "frmCnsAvancadaMob.frx":1ACB
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
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Final:"
            Height          =   195
            Index           =   5
            Left            =   4095
            TabIndex        =   43
            Top             =   315
            Width           =   840
         End
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Inicial:"
            Height          =   195
            Index           =   4
            Left            =   2025
            TabIndex        =   42
            Top             =   315
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Data de Encerramento"
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   4680
         TabIndex        =   38
         Top             =   1710
         Width           =   6360
         Begin VB.ComboBox cmbDEnc 
            Height          =   315
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   270
            Width           =   1455
         End
         Begin esMaskEdit.esMaskedEdit mskDataEncIni 
            Height          =   285
            Left            =   2925
            TabIndex        =   12
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MouseIcon       =   "frmCnsAvancadaMob.frx":1AE7
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
         Begin esMaskEdit.esMaskedEdit mskDataEncFim 
            Height          =   285
            Left            =   4905
            TabIndex        =   13
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MouseIcon       =   "frmCnsAvancadaMob.frx":1B03
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
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Inicial:"
            Height          =   195
            Index           =   3
            Left            =   2025
            TabIndex        =   40
            Top             =   315
            Width           =   840
         End
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Final:"
            Height          =   195
            Index           =   2
            Left            =   4095
            TabIndex        =   39
            Top             =   315
            Width           =   840
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Data de Abertura"
         ForeColor       =   &H00800000&
         Height          =   690
         Left            =   4680
         TabIndex        =   35
         Top             =   990
         Width           =   6360
         Begin VB.CheckBox chkDAbe 
            BackColor       =   &H00EEEEEE&
            Caption         =   "Todos"
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   315
            Value           =   1  'Checked
            Width           =   870
         End
         Begin esMaskEdit.esMaskedEdit mskDataAbeIni 
            Height          =   285
            Left            =   2925
            TabIndex        =   9
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MouseIcon       =   "frmCnsAvancadaMob.frx":1B1F
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
         Begin esMaskEdit.esMaskedEdit mskDataAbeFim 
            Height          =   285
            Left            =   4905
            TabIndex        =   10
            Top             =   270
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            MouseIcon       =   "frmCnsAvancadaMob.frx":1B3B
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
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Final:"
            Height          =   195
            Index           =   0
            Left            =   4095
            TabIndex        =   37
            Top             =   315
            Width           =   840
         End
         Begin VB.Label lblVenc 
            BackStyle       =   0  'Transparent
            Caption         =   "Data Inicial:"
            Height          =   195
            Index           =   1
            Left            =   2025
            TabIndex        =   36
            Top             =   315
            Width           =   840
         End
      End
   End
   Begin Tributacao.jcFrames fr2 
      Height          =   5865
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   10345
      FrameColor      =   8388608
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Lista das empresas selecionadas pelos critérios"
      TextBoxHeight   =   18
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
      Begin vbAcceleratorSGrid6.vbalGrid grdMain 
         Height          =   2610
         Left            =   90
         TabIndex        =   32
         Top             =   855
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   4604
         NoHorizontalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         HighlightBackColor=   128
         HighlightForeColor=   16777215
         GroupRowForeColor=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         ScrollBarStyle  =   1
         DisableIcons    =   -1  'True
         DrawFocusRectangle=   0   'False
         GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
      End
      Begin VB.Frame frDDList1 
         BackColor       =   &H00EEEEEE&
         Height          =   375
         Left            =   3150
         TabIndex        =   88
         Top             =   315
         Width           =   2580
         Begin VB.ListBox lstDDList1 
            Height          =   3210
            ItemData        =   "frmCnsAvancadaMob.frx":1B57
            Left            =   45
            List            =   "frmCnsAvancadaMob.frx":1B59
            Style           =   1  'Checkbox
            TabIndex        =   89
            Top             =   405
            Width           =   2490
         End
         Begin prjChameleon.chameleonButton cmdDDList1 
            Height          =   240
            Left            =   315
            TabIndex        =   90
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   423
            BTYPE           =   14
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
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1B5B
            PICN            =   "frmCnsAvancadaMob.frx":1B77
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDD1All 
            Height          =   240
            Left            =   1710
            TabIndex        =   91
            ToolTipText     =   "Selecionar todos"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1CD1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDD1None 
            Height          =   240
            Left            =   2070
            TabIndex        =   92
            ToolTipText     =   "Manter apenas o código"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "-"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1CED
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
      Begin prjChameleon.chameleonButton cmdGroup 
         Height          =   315
         Left            =   9900
         TabIndex        =   104
         ToolTipText     =   "Avançar para a próxima tela"
         Top             =   405
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Agrupar"
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
         MICON           =   "frmCnsAvancadaMob.frx":1D09
         PICN            =   "frmCnsAvancadaMob.frx":1D25
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin Tributacao.jcFrames jcFrames1 
         Height          =   2115
         Left            =   90
         Top             =   3600
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   3731
         FrameColor      =   12829635
         Style           =   0
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "Resumo mensal de débitos"
         TextColor       =   13579779
         Alignment       =   0
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
         Begin VB.ComboBox cmbExercicio 
            Height          =   315
            Left            =   1110
            Style           =   2  'Dropdown List
            TabIndex        =   128
            Top             =   330
            Width           =   1095
         End
         Begin prjChameleon.chameleonButton cmdPrintDebito 
            Height          =   375
            Left            =   150
            TabIndex        =   129
            ToolTipText     =   "Relatório de débitos"
            Top             =   870
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Gerar relatório"
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
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":1E7F
            PICN            =   "frmCnsAvancadaMob.frx":1E9B
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
            Caption         =   "Exercício...:"
            Height          =   225
            Left            =   150
            TabIndex        =   127
            Top             =   390
            Width           =   945
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Campos a serem exibidos..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   31
         Top             =   405
         Width           =   2850
      End
   End
   Begin Tributacao.jcFrames tb1 
      Height          =   465
      Left            =   45
      Top             =   5940
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   820
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
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
      Begin prjChameleon.chameleonButton cmdConsultar 
         Height          =   315
         Left            =   7380
         TabIndex        =   26
         ToolTipText     =   "Consultar as empresas"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Pesquisar"
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
         MICON           =   "frmCnsAvancadaMob.frx":22F3
         PICN            =   "frmCnsAvancadaMob.frx":230F
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
         Left            =   8685
         TabIndex        =   27
         ToolTipText     =   "Avançar para a próxima tela"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Continuar"
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
         MICON           =   "frmCnsAvancadaMob.frx":243E
         PICN            =   "frmCnsAvancadaMob.frx":245A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Tributacao.XP_ProgressBar PBar 
         Height          =   240
         Left            =   4320
         TabIndex        =   87
         Top             =   135
         Width           =   2940
         _ExtentX        =   5186
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
      Begin prjChameleon.chameleonButton cmdSair 
         Height          =   315
         Left            =   9990
         TabIndex        =   103
         ToolTipText     =   "Sair da Tela"
         Top             =   90
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
         MICON           =   "frmCnsAvancadaMob.frx":25B4
         PICN            =   "frmCnsAvancadaMob.frx":25D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Números de empresas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   29
         Top             =   90
         Width           =   2805
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   3015
         TabIndex        =   28
         Top             =   90
         Width           =   915
      End
   End
   Begin Tributacao.jcFrames fr3 
      Height          =   5865
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   10345
      FrameColor      =   6974058
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextBoxHeight   =   18
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
      Begin VB.Frame Frame7 
         BackColor       =   &H00EEEEEE&
         Height          =   500
         Left            =   45
         TabIndex        =   56
         Top             =   270
         Width           =   11220
         Begin VB.ComboBox cmbCampos 
            Height          =   315
            ItemData        =   "frmCnsAvancadaMob.frx":263E
            Left            =   8370
            List            =   "frmCnsAvancadaMob.frx":2640
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   135
            Width           =   2355
         End
         Begin prjChameleon.chameleonButton cmdInsertBullet 
            Height          =   315
            Left            =   2835
            TabIndex        =   75
            ToolTipText     =   "Inserir marcador"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":2642
            PICN            =   "frmCnsAvancadaMob.frx":265E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdForeColor 
            Height          =   315
            Left            =   3870
            TabIndex        =   74
            ToolTipText     =   "Cor da letra"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "A"
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
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   192
            FCOLO           =   192
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":29B0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAddField 
            Height          =   315
            Left            =   10755
            TabIndex        =   66
            ToolTipText     =   "Adicionar campo ao texto"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   14
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":29CC
            PICN            =   "frmCnsAvancadaMob.frx":29E8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdFU 
            Height          =   315
            Left            =   2205
            TabIndex        =   64
            ToolTipText     =   "Sublinhado"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":2BCB
            PICN            =   "frmCnsAvancadaMob.frx":2BE7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdFB 
            Height          =   315
            Left            =   1395
            TabIndex        =   62
            ToolTipText     =   "Negrito"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":2DCD
            PICN            =   "frmCnsAvancadaMob.frx":2DE9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin VB.ComboBox cmbFonte 
            Height          =   315
            Left            =   5040
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   135
            Width           =   1905
         End
         Begin VB.ComboBox cmbTam 
            Height          =   315
            Left            =   6975
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   135
            Width           =   870
         End
         Begin prjChameleon.chameleonButton cmdAL 
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   59
            ToolTipText     =   "Alinhar a esquerda"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":2FD3
            PICN            =   "frmCnsAvancadaMob.frx":2FEF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAL 
            Height          =   315
            Index           =   2
            Left            =   495
            TabIndex        =   60
            ToolTipText     =   "Centralizar"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":31D3
            PICN            =   "frmCnsAvancadaMob.frx":31EF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdAL 
            Height          =   315
            Index           =   1
            Left            =   900
            TabIndex        =   61
            ToolTipText     =   "Alinhar a direita"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":33D0
            PICN            =   "frmCnsAvancadaMob.frx":33EC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdFI 
            Height          =   315
            Left            =   1800
            TabIndex        =   63
            ToolTipText     =   "Itálico"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
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
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":35CD
            PICN            =   "frmCnsAvancadaMob.frx":35E9
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdBackColor 
            Height          =   315
            Left            =   3465
            TabIndex        =   73
            ToolTipText     =   "Cor do fundo"
            Top             =   135
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "A"
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
            BCOL            =   12632319
            BCOLO           =   12632319
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":37CF
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdData 
            Height          =   315
            Left            =   4275
            TabIndex        =   113
            ToolTipText     =   "Inserir data atual"
            Top             =   135
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "Data"
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
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaMob.frx":37EB
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
      Begin RichTextLib.RichTextBox Rtb 
         Height          =   5055
         Left            =   45
         TabIndex        =   65
         Top             =   765
         Width           =   11220
         _ExtentX        =   19791
         _ExtentY        =   8916
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmCnsAvancadaMob.frx":3807
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "mnuPrint"
      Visible         =   0   'False
      Begin VB.Menu mnuTela 
         Caption         =   "Gerar na Tela"
      End
      Begin VB.Menu mnuEtiq 
         Caption         =   "Gerar Etiquetas"
      End
      Begin VB.Menu mnuCartas 
         Caption         =   "Gerar Cartas"
      End
      Begin VB.Menu mnuTxt 
         Caption         =   "Gerar em TXT"
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "Gerar em Excel"
      End
      Begin VB.Menu mnuCadastro 
         Caption         =   "Imprimir Dados Cadastrais"
      End
      Begin VB.Menu mnuExtrato 
         Caption         =   "Imprimir Extrato"
      End
      Begin VB.Menu mnuQtdeAtividade 
         Caption         =   "Imprimir quantidade por atividade"
      End
   End
End
Attribute VB_Name = "frmCnsAvancadaMob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNotificado As String
    sExFiscal As String
End Type

Private Type Empresa
    nCodigo As Long
    sRazao As String
    sCNPJ As String
    sCPF As String
    sEndereco As String
    sComplemento As String
    sBairro As String
    sCidade As String
    sUF As String
    sCep As String
    sDataAbertura As String
    sDataEncerramento As String
    sDataSuspensao As String
    sAtividade As String
    sAtivExtenso As String
    sHorario As String
    sContador As String
    sMEI As String
    sSN As String
    sIE As String
    sISSElet As String
    sDispensaIE As String
    sIsentoTaxa As String
    sAlvara As String
    sVistoria As String
    sVigilancia As String
    sFixo As String
    sEstimado As String
    sVariavel As String
    sProcEncerramento As String
    sNomeContato As String
    sFoneContato As String
    sProcAbertura As String
    nParcPagtoTotal As Integer
    nParcPagtoSim As Integer
    nParcPagtoNao As Integer
    sNomeFantasia As String
    sAtivIss As String
    sEnderecoEnt As String
    sComplementoEnt As String
    sBairroEnt As String
    sCidadeEnt As String
    sUFEnt As String
    sCEPEnt As String
End Type

Private Type Suspenso
    nCodigo As Long
    dData As Date
End Type

Private Type EmpresaAtividade
    nCodigo As Long
    sAtividade As String
End Type

Dim aCodigos() As Empresa, aSuspenso() As Suspenso, aSuspensoCod() As Long, sNomeModelo As String, nIndent As Integer
Dim aVigilancia() As Long, aFixo() As Long, aEstimado() As Long, aVariavel() As Long, sTitulo As String
Dim aCodigosImp() As Long, strCodigos As String, sCodEmpresaIss As String, aEmpresaISS() As EmpresaAtividade
Dim xImovel As clsImovel

Private Sub chkAtividade_Click()
If chkAtividade.value = vbChecked Then
    cmdDDList3.Enabled = False
'    chkTipoAtiv(1).Enabled = True
'    chkTipoAtiv(2).Enabled = True
'    chkTipoAtiv(3).Enabled = True
'    chkTipoAtiv(4).Enabled = True
Else
    cmdDDList3.Enabled = True
 '   chkTipoAtiv(1).Enabled = False
 '   chkTipoAtiv(2).Enabled = False
  '  chkTipoAtiv(3).Enabled = False
  '  chkTipoAtiv(4).Enabled = False
End If

End Sub

Private Sub chkAtivIss_Click()
If chkAtivIss.value = vbChecked Then
    cmdDDList4.Enabled = False
Else
    cmdDDList4.Enabled = True
End If

End Sub

Private Sub chkContador_Click()
If chkContador.value = vbChecked Then
    cmdDDList2.Enabled = False
Else
    cmdDDList2.Enabled = True
End If
End Sub

Private Sub cmbDEnc_Click()
If cmbDEnc.ListIndex = 1 Then
    mskDataEncIni.BackColor = vbWhite
    mskDataEncFim.BackColor = vbWhite
    mskDataEncIni.Locked = False
    mskDataEncFim.Locked = False
    mskDataEncIni.SetFocus
Else
    mskDataEncIni.BackColor = Kde
    mskDataEncFim.BackColor = Kde
    mskDataEncIni.Locked = True
    mskDataEncFim.Locked = True
End If

End Sub

Private Sub cmbDSus_Click()
If cmbDSus.ListIndex = 1 Then
    mskDataSusIni.BackColor = vbWhite
    mskDataSusFim.BackColor = vbWhite
    mskDataSusIni.Locked = False
    mskDataSusFim.Locked = False
    mskDataSusIni.SetFocus
Else
    mskDataSusIni.BackColor = Kde
    mskDataSusFim.BackColor = Kde
    mskDataSusIni.Locked = True
    mskDataSusFim.Locked = True
End If

End Sub

Private Sub cmbFonte_Click()
If cmbFonte.ListIndex = -1 Then Exit Sub
Rtb.SelFontName = cmbFonte.Text
End Sub

Private Sub cmbImposto_Click()
lblTot.Caption = 0
End Sub

Private Sub cmbISS_Click()
CarregaListaIss

End Sub

Private Sub cmbPagto_Click()
lblTot.Caption = 0
If cmbPagto.ListIndex > -1 Then
    If cmbPagto.ListIndex = 0 Then
        cmbImposto.ListIndex = -1
        cmbImposto.Enabled = False
        txtAno.BackColor = Kde
        txtAno.Locked = True
    Else
        cmbImposto.Enabled = True
        txtAno.BackColor = Branco
        txtAno.Locked = False
    End If
End If
End Sub

Private Sub cmbTam_Click()
If cmbTam.ListIndex = -1 Then Exit Sub
Rtb.SelFontSize = cmbTam.Text
End Sub

Private Sub cmdAddField_Click()
Rtb.SelText = "[#" & cmbCampos.Text & "#]"
End Sub

Private Sub cmdAL_Click(Index As Integer)
Rtb.SelAlignment = Index
End Sub

Private Sub cmdBackColor_Click()
Dim lColor As Long, cc As cCommonDlg

Set cc = New cCommonDlg
lColor = cmdBackColor.BackColor
cc.VBChooseColor lColor
If lColor > -1 Then
    cmdBackColor.BackColor = lColor
    cmdBackColor.BackOver = lColor
    HighLight Rtb, cmdBackColor.BackColor
End If
End Sub

Private Sub cmdCarregar_Click()
CarregaModelo
frModelo.Visible = True
frModelo.ZOrder (0)
End Sub

Private Sub cmdCloseModel_Click()
frModelo.Visible = False
End Sub

Private Sub cmdData_Click()
Rtb.SelText = "Jaboticabal," & Format(Now, "dd", vbLongDate) & " de " & Format(Now, "mmmm", vbLongDate) & " de " & Format(Now, "yyyy", vbLongDate)
End Sub

Private Sub cmdDD1All_Click()
Dim x As Integer

For x = 0 To lstDDList1.ListCount - 1
    lstDDList1.Selected(x) = True
Next
End Sub

Private Sub cmdDD1None_Click()
Dim x As Integer
For x = 0 To lstDDList1.ListCount - 1
    lstDDList1.Selected(x) = False
Next
lstDDList1.Selected(0) = True
End Sub

Private Sub cmdDD2All_Click()
Dim x As Integer
For x = 0 To lstDDList2.ListCount - 1
    lstDDList2.Selected(x) = True
Next

End Sub

Private Sub cmdDD2None_Click()
Dim x As Integer
For x = 0 To lstDDList2.ListCount - 1
    lstDDList2.Selected(x) = False
Next
lstDDList2.Selected(0) = True

End Sub

Private Sub cmdDDList1_Click()
If cmdDDList1.value = True Then
    frDDList1.Height = 3660
    frDDList1.ZOrder 0
Else
    frDDList1.Height = 375
    HideColumns
End If
End Sub

Private Sub cmdDDList2_Click()

If cmdDDList2.value = True Then
    frDDList2.Height = 3660
    frDDList2.ZOrder 0
Else
    frDDList2.Height = 375
End If

End Sub

Private Sub cmdDDList3_Click()
If cmdDDList3.value = True Then
    frDDList3.Height = 3660
    frDDList3.ZOrder 0
Else
    frDDList3.Height = 375
End If

End Sub

Private Sub cmdDDList3All_Click()
Dim x As Integer
For x = 0 To lstDDList3.ListCount - 1
    lstDDList3.Selected(x) = True
Next

End Sub

Private Sub cmdDDList3None_Click()
Dim x As Integer
For x = 0 To lstDDList3.ListCount - 1
    lstDDList3.Selected(x) = False
Next
lstDDList3.Selected(0) = True

End Sub

Private Sub cmdDDList4_Click()
If cmdDDList4.value = True Then
    frDDList4.Height = 2100
    frDDList4.ZOrder 0
    frIssAtiv.Top = 3150
    frDDList4.Top = 3330
    frIssAtiv.Height = 2400
    lstDDList4.Height = 1800
    
Else
    frDDList4.Height = 375
    frIssAtiv.Top = 4860
    frDDList4.Top = 5340
    frIssAtiv.Height = 945
End If

End Sub

Private Sub cmdDDList4All_Click()
Dim x As Integer
For x = 0 To lstDDList4.ListCount - 1
    lstDDList4.Selected(x) = True
Next

End Sub

Private Sub cmdDDList4None_Click()
Dim x As Integer
For x = 0 To lstDDList4.ListCount - 1
    lstDDList4.Selected(x) = False
Next
lstDDList4.Selected(0) = True

End Sub

Private Sub cmdDebitos_Click()
If cmdDebitos.value = True Then
    grdMain.Height = 2610
Else
    grdMain.Height = 4950
    grdMain.ZOrder 0
End If
End Sub

Private Sub cmdFB_Click()
Rtb.SelBold = cmdFB.value
End Sub

Private Sub cmdFI_Click()
Rtb.SelItalic = cmdFI.value
End Sub

Private Sub cmdForeColor_Click()
Dim lColor As Long, cc As cCommonDlg

Set cc = New cCommonDlg
lColor = cmdForeColor.ForeColor
cc.VBChooseColor lColor
If lColor > -1 Then
    cmdForeColor.ForeColor = lColor
    Rtb.SelColor = cmdForeColor.ForeColor
End If
End Sub

Private Sub cmdFU_Click()
Rtb.SelUnderline = cmdFU.value
End Sub

Private Sub cmdGravar_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset

z = InputBox("Digite um nome para o modelo", "Nome do Modelo", sNomeModelo)
If z = "" Then Exit Sub

Sql = "SELECT * FROM MMGMODELO WHERE NOME='" & z & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
    If MsgBox("Já existe um modelo com este nome, voce deseja substitui-lo?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        Sql = "UPDATE MMGMODELO SET MODELO='" & Mask(Rtb.TextRTF) & "' WHERE NOME='" & z & "'"
        cn.Execute Sql, rdExecDirect
        fr3.Caption = sTitulo & z
    End If
Else
    Sql = "INSERT MMGMODELO(NOME,MODELO) VALUES('" & z & "','" & Mask(Rtb.TextRTF) & "')"
    cn.Execute Sql, rdExecDirect
    sNomeModelo = z
    fr3.Caption = sTitulo & z
End If

End Sub

Private Sub cmdGroup_Click()
grdMain.AllowGrouping = cmdGroup.value
End Sub

Private Sub cmdImportar_Click()
Dim strLinha As String, z As Variant, x As Integer, nCodigo As Long
lblTotImp.Caption = 0
If txtDelimiter.Text = "" Then
    MsgBox "Especifique um delimitador", vbCritical, "Erro"
    Exit Sub
End If
If txtArq.Text = "" Then
    MsgBox "Selecione um arquivo", vbCritical, "Erro"
    Exit Sub
End If

ReDim aCodigosImp(0): strCodigos = ""
Open txtArq.Text For Input As #1
   Do While Not EOF(1)
        Line Input #1, strLinha
        z = Split(strLinha, txtDelimiter.Text)
        For x = 0 To UBound(z)
            If Not IsNumeric(z(x)) Then
               GoTo proximo
            End If
            nCodigo = CLng(z(x))
            If nCodigo < 100000 Or nCodigo > 500000 Then
               GoTo Erro
            End If
            ReDim Preserve aCodigosImp(UBound(aCodigosImp) + 1)
            aCodigosImp(UBound(aCodigosImp)) = nCodigo
            strCodigos = strCodigos & nCodigo & ","
        Next
proximo:
   Loop
Close #1
strCodigos = Chomp(strCodigos, chomp_righT, 1)
lblTotImp.Caption = UBound(aCodigosImp)

Exit Sub
Erro:
MsgBox "Arquivo inválido !!!", vbCritical, "Erro de importação"
Close #1

End Sub

Private Sub cmdImprimir_Click()
Dim nGrid As Integer, nCampo As Integer, Sql As String, nPos As Integer

Ocupado
Rtb.TextRTF = Replace(Rtb.TextRTF, "\pard", "\pard\qj")
Sql = "DELETE FROM MMGREGISTRO WHERE USUARIO='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
On Error Resume Next
nPos = Val(txtAutoInc.Text)
With grdMain
    For nGrid = 1 To .Rows
        RtbTmp.TextRTF = Rtb.TextRTF
        For nCampo = 0 To cmbCampos.ListCount - 1
            If cmbCampos.List(nCampo) = "Auto Incremento" Then
                RtbTmp.TextRTF = Replace(RtbTmp.TextRTF, "[#" & cmbCampos.List(nCampo) & "#]", nPos)
            Else
                RtbTmp.TextRTF = Replace(RtbTmp.TextRTF, "[#" & cmbCampos.List(nCampo) & "#]", .CellText(nGrid, nCampo + 1))
            End If
        Next
        Sql = "INSERT MMGREGISTRO(USUARIO,SEQ,TEXTO) VALUES('" & NomeDoUsuario & "'," & nGrid & ",'" & Mask(RtbTmp.TextRTF) & "')"
        cn.Execute Sql, rdExecDirect
        DoEvents
        nPos = nPos + 1
    Next
End With
Liberado
frmReport.ShowReport "MMG", frmMdi.HWND, Me.HWND
Sql = "DELETE FROM MMGREGISTRO WHERE USUARIO='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdInsertBullet_Click()
Rtb.SelBullet = True
End Sub

Private Sub cmdLimpar_Click()
txtArq.Text = "": lblTotImp.Caption = 0
End Sub

Private Sub cmdOpen_Click()
Dim fName As String, cc As cCommonDlg

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Documento de Texto|*.txt;*.csv|Todos os Arquivos|*.*", , App.Path & "\Bin", "Selecione um arquivo texto", , Me.HWND, OFN_HIDEREADONLY, False
txtArq.Text = fName

End Sub

Private Sub cmdPreview_Click()
If (txtArq.Text) <> "" Then
    z = Shell(App.Path & "\NOTEPAD2" & " " & txtArq.Text, vbNormalFocus)
End If
End Sub

Private Sub cmdPrint_Click()
PopupMenu mnuPrint, , cmdPrint.Left
End Sub

Private Sub cmdSelectModel_Click()
Dim Sql As String, RdoAux As rdoResultset

If lstModelo.ListIndex > -1 Then
    Sql = "SELECT * FROM MMGMODELO WHERE NOME='" & lstModelo.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    Rtb.TextRTF = RdoAux!MODELO
    RdoAux.Close
    sNomeModelo = lstModelo.Text
End If
fr3.Caption = sTitulo & sNomeModelo
End Sub

Private Sub cmdVoltar2_Click()
tb3.Visible = False
tb2.Visible = True
fr3.Visible = False
fr2.Visible = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub mnuCadastro_Click()
Dim Sql As String, x As Integer

Sql = "DELETE FROM TBDADOSEMPRESA WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

With grdMain
    For x = 1 To .Rows
        Sql = "INSERT TBDADOSEMPRESA(USUARIO,CODIGO) VALUES('" & NomeDeLogin & "'," & Val(.CellText(x, 1)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

frmReport.ShowReport "CADMOBILIARIO", frmMdi.HWND, Me.HWND

Sql = "DELETE FROM TBDADOSEMPRESA WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mnuCartas_Click()
tb2.Visible = False
tb3.Visible = True
fr2.Visible = False
fr3.Visible = True
End Sub

Private Sub mnuEtiq_Click()
Dim x As Integer
Ocupado
If cGetInputState() <> 0 Then DoEvents
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
    
With grdMain
    For x = 1 To .Rows
        If cGetInputState() <> 0 Then DoEvents
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & x & ",'" & .CellText(x, 1) & "','" & Mask(.CellText(x, 2)) & "','"
        Sql = Sql & Left(.CellText(x, 39) & " " & .CellText(x, 40), 60) & "','" & .CellText(x, 41) & " - " & .CellText(x, 42) & "','" & .CellText(x, 43) & "   " & .CellText(x, 44) & "')"
'        sql = sql & NomeDeLogin & "'," & x & ",'" & .CellText(x, 1) & "','" & Mask(.CellText(x, 2)) & "','"
'        sql = sql & Left(.CellText(x, 5) & " " & .CellText(x, 6), 60) & "','" & .CellText(x, 7) & " - " & .CellText(x, 8) & "','" & .CellText(x, 9) & "   " & .CellText(x, 10) & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With
Liberado
If cGetInputState() <> 0 Then DoEvents
frmReport.ShowReport "ETIQUETACONSIST", frmMdi.HWND, Me.HWND
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mnuExcel_Click()
Dim x As Long, Y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
         
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For Y = 1 To grdMain.Columns
    If grdMain.ColumnVisible(Y) = True Then
        ax = ax & RemoveSpace(grdMain.ColumnHeader(Y)) & " char(255), "
    End If
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "[Table1$]", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To grdMain.Rows
    Rs.AddNew
    nCont = 0
    For Y = 1 To grdMain.Columns
        If grdMain.ColumnVisible(Y) = True Then
            Rs.Fields(nCont).value = Left(grdMain.cell(x, Y).Text, 100)
            nCont = nCont + 1
        End If
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)

End Sub

Private Sub mnuExcelOld_Click()
Dim myExcelFile As New ExcelFile, x As Long, Y As Long, ax As String, Scr_hdc As Long, z As Long
Scr_hdc = GetDesktopWindow()
         
          
With myExcelFile
    FileName$ = sPathBin & "\Relatorio.xls"  'create spreadsheet in the current directory
    .CreateFile FileName$
    
    ax = ""
    For Y = 1 To grdMain.Columns
        If grdMain.ColumnVisible(Y) = True Then
            .WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, 1, Y, grdMain.ColumnHeader(Y)
        End If
    Next
    For x = 2 To grdMain.Rows
        If cGetInputState() <> 0 Then DoEvents
        ax = ""
        For Y = 1 To grdMain.Columns
            If grdMain.ColumnVisible(Y) = True Then
                .WriteValue xlsText, xlsFont3, xlsLeftAlign, xlsNormal, x, Y, grdMain.cell(x, Y).Text
            End If
        Next
    Next
   .CloseFile
End With

z = ShellExecute(Scr_hdc, "Open", "Relatorio.xls", "", sPathBin, SW_SHOWNORMAL)

End Sub

Private Sub mnuExtrato_Click()
Dim x As Integer

'MORREK ZMANI
Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

For x = 1 To grdMain.Rows
    GravaExtrato2 Val(grdMain.CellText(x, 1))
Next


'EXIBE RELATORIO
frmReport.ShowReport "Extrato3", frmMdi.HWND, Me.HWND
'MORREK ZMANI
Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mnuQtdeAtividade_Click()

Dim nCodReduz As Long, sRazao As String, sAtividade As String, nGrupo As Integer, Sql As String


Sql = "delete from rel_empresa_qtde_por_atividade where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

With grdMain
    For x = 1 To .Rows
        nCodReduz = .CellText(x, 1)
        sRazao = .CellText(x, 2)
        sAtividade = .CellText(x, 14)
        nGrupo = Val(Left(sAtividade, 1))
        
        Sql = "insert rel_empresa_qtde_por_atividade(usuario,codigo,razao_social,atividade,grupo) values('" & NomeDeLogin & "'," & nCodReduz & ",'" & sRazao & "','" & sAtividade & " '," & nGrupo & ")"
        cn.Execute Sql, rdExecDirect
        
    Next
End With
frmReport.ShowReport2 "EMPRESA_QTDEATIVIDADE", frmMdi.HWND, Me.HWND
Sql = "delete from rel_empresa_qtde_por_atividade where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mnuTela_Click()
Dim ax As String, z As Long, x As Integer, Y As Integer

Open sPathBin & "\TEMPMOB.TXT" For Output As #1

With grdMain
    ax = ""
    For Y = 1 To .Columns
        If .ColumnVisible(Y) = True Then
            ax = ax & FillSpace(.ColumnHeader(Y), Val(Right(.ColumnKey(Y), 3))) & vbTab
        End If
    Next
    Print #1, ax
    For x = 1 To .Rows
        ax = ""
        For Y = 1 To .Columns
            If .ColumnVisible(Y) = True Then
                ax = ax & FillSpace(.cell(x, Y).Text, Val(Right(.ColumnKey(Y), 3))) & vbTab
            End If
        Next
        Print #1, ax
    Next

End With

Close #1

z = Shell(App.Path & "\NOTEPAD2" & " " & sPathBin & "\TEMPMOB.TXT", vbNormalFocus)

End Sub

Private Sub mnuTxt_Click()
Dim ax As String, z As Long, x As Integer, Y As Integer, sChar As String

If txtSep.Text = "" Then
    sChar = " "
Else
    sChar = txtSep.Text
End If
Ocupado
Open sPathBin & "\RELATMOB.CSV" For Output As #1

With grdMain
    ax = ""
    For Y = 1 To .Columns
        If .ColumnVisible(Y) = True Then
            ax = ax & FillSpace(.ColumnHeader(Y), Val(Right(.ColumnKey(Y), 3))) & sChar
        End If
    Next
    ax = Chomp(ax, chomp_righT, 1)
    Print #1, ax
    For x = 1 To .Rows
        If cGetInputState() <> 0 Then DoEvents
        ax = ""
        For Y = 1 To .Columns
            If .ColumnVisible(Y) = True Then
                ax = ax & FillSpace(.cell(x, Y).Text, Val(Right(.ColumnKey(Y), 3))) & sChar
            End If
        Next
        If sChar <> " " Then ax = Chomp(ax, chomp_righT, 1)
        Print #1, ax
    Next

End With

Close #1
Liberado
MsgBox "O arquivo foi salvo em " & sPathBin & "\RELATMOB.CSV"
End Sub

Private Sub mskDataAbeFim_GotFocus()
mskDataAbeFim.SelStart = 0
mskDataAbeFim.SelLength = Len(mskDataAbeFim.Text)
End Sub

Private Sub mskDataAbeIni_GotFocus()
mskDataAbeIni.SelStart = 0
mskDataAbeIni.SelLength = Len(mskDataAbeIni.Text)
End Sub

Private Sub mskDataEncFim_GotFocus()
mskDataEncFim.SelStart = 0
mskDataEncFim.SelLength = Len(mskDataEncFim.Text)
End Sub

Private Sub mskDataEncIni_GotFocus()
mskDataEncIni.SelStart = 0
mskDataEncIni.SelLength = Len(mskDataEncIni.Text)
End Sub

Private Sub chkDAbe_Click()
If chkDAbe.value = vbChecked Then
    mskDataAbeIni.BackColor = Kde
    mskDataAbeFim.BackColor = Kde
    mskDataAbeIni.Locked = True
    mskDataAbeFim.Locked = True
Else
    mskDataAbeIni.BackColor = vbWhite
    mskDataAbeFim.BackColor = vbWhite
    mskDataAbeIni.Locked = False
    mskDataAbeFim.Locked = False
    mskDataAbeIni.SetFocus
End If
End Sub

Private Sub cmdNext_Click()

grdMain.Height = 4950
If Val(lblTot.Caption) > 0 Then
    If Not Valida Then Exit Sub
    cmdNext.Enabled = False
    Ocupado
    If cGetInputState() <> 0 Then DoEvents
    CarregaCampos
    CarregaLista
    HideColumns
    Liberado
    tb2.Visible = True
    tb1.Visible = False
    fr2.Visible = True
    fr1.Visible = False
    grdMain.SetFocus
    grdMain.SelectedRow = 1
    cmdNext.Enabled = True
Else
    MsgBox "Nenhuma empresa possui os critérios selecionados ou não foi gerada consulta.", vbExclamation, "Atenção"
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub cmdVoltar_Click()
tb2.Visible = False
tb1.Visible = True
fr2.Visible = False
fr1.Visible = True
End Sub

Private Sub Form_Load()
Set xImovel = New clsImovel
Init
fr3.Caption = sTitulo & "Sem Nome"
Centraliza Me
GridHeader
CarregaCriterio
cmdDDList2.Enabled = False
cmdDDList3.Enabled = False
cmdDDList4.Enabled = False

cmdDebitos.Enabled = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnu2ViaLaser"))

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC00000
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


Private Sub grdMain_ColumnClick(ByVal lcol As Long)

Dim sTag As String
Dim iSortIndex As Long
      
   With grdMain.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lcol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lcol
      End If
   
      ' Determine which sort order to apply:
      sTag = grdMain.ColumnTag(lcol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdMain.ColumnTag(lcol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdMain.ColumnSortType(lcol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdMain.Sort
   Screen.MousePointer = vbDefault

End Sub

Private Sub HideColumns()
Dim x As Integer, Y As Integer, bAchou As Boolean

bAchou = False
For x = 0 To lstDDList1.ListCount - 1
    If lstDDList1.Selected(x) = True Then
        bAchou = True
        Exit For
    End If
Next
If Not bAchou Then
    MsgBox "Voce deve selecionar pelo menos um campo para exibição.", vbExclamation, "Atenção"
    Exit Sub
End If


For x = 0 To lstDDList1.ListCount - 1
'    If lstDDList1.Selected(x) = True Then MsgBox "teste"
     grdMain.ColumnVisible(x + 1) = lstDDList1.Selected(x)
Next

End Sub

Private Sub cmdConsultar_Click()

If optIM(0).value = True Then
    ConsultaEmpresa
Else
    ConsultaCidadao
End If

End Sub

Private Sub ConsultaCidadao()
Dim Sql As String, RdoAux As rdoResultset, nTotal As Integer, RdoAux2 As rdoResultset, cn2 As rdoConnection, nTot As Long, nPos As Long, qd As New rdoQuery, nLanc As Integer
Dim s As Integer, lResult As Long, sDataSus As String, sContador As String, x As Integer, sAtividade As String, sSimples As String, sVig As String, sFixo As String, sVariavel As String, sEstimado As String
Dim nQtdePagtoTotal As Integer, nQtdePagtoNao As Integer, nQtdePagtoSim As Integer, sISSEletr As String
If Not Valida Then Exit Sub

lblTot.Caption = 0
Set cn2 = en.OpenConnection(dsname:="odbcTributacao", Prompt:=rdDriverNoPrompt, Connect:="uid=" & NomeDeLogin & ";PWD=" & sWd & ";driver={SQL Server};")

Ocupado
If cGetInputState() <> 0 Then DoEvents
nTotal = 0: ReDim aCodigos(0): nPos = 1
cmdConsultar.Enabled = False
Sql = "SELECT * FROM vwFULLCIDADAO WHERE "
If Val(lblTotImp.Caption) = 0 Then
    Sql = Sql & "JURIDICA = 1"
Else
    Sql = Sql & "CODCIDADAO in (" & strCodigos & ")"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
        End If
        
        sSimples = ""
        sISSEletr = ""
        sFixo = ""
        sEstimado = ""
        sVariavel = ""
        sVig = ""
        If !CodCidadao < 500000 Then GoTo proximo
        ReDim Preserve aCodigos(UBound(aCodigos) + 1)
        aCodigos(UBound(aCodigos)).nCodigo = !CodCidadao
        aCodigos(UBound(aCodigos)).sRazao = !nomecidadao
        aCodigos(UBound(aCodigos)).sCNPJ = SubNull(!Cnpj)
        aCodigos(UBound(aCodigos)).sCPF = SubNull(!CPF)
        aCodigos(UBound(aCodigos)).sEndereco = Trim(SubNull(!Endereco) & " nº " & Val(SubNull(!NUMIMOVEL)))
        aCodigos(UBound(aCodigos)).sComplemento = SubNull(!Complemento)
        aCodigos(UBound(aCodigos)).sBairro = SubNull(!DescBairro)
        aCodigos(UBound(aCodigos)).sCidade = SubNull(!descCidade)
        aCodigos(UBound(aCodigos)).sUF = SubNull(!SiglaUF)
        If Not IsNull(!Cep) And !CodCidade <> 413 Then
            aCodigos(UBound(aCodigos)).sCep = Format(SubNull(!Cep), "00000-000")
        Else
            aCodigos(UBound(aCodigos)).sCep = Format(RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!NUMIMOVEL))), "00000-000")
        End If
        nTotal = nTotal + 1
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
    lblTot.Caption = nTotal
   .Close
End With
PBar.value = 0: PBar.Color = vbWhite
cn2.Close
cmdConsultar.Enabled = True
Liberado
If nTotal = 0 Then
    MsgBox "Nenhuma empresa (cidadão) esta como pessoa jurídica .", vbInformation, "Informação"
End If

End Sub

Private Sub ConsultaEmpresa()
Dim Sql As String, RdoAux As rdoResultset, nTotal As Integer, RdoAux2 As rdoResultset, cn2 As rdoConnection, nTot As Long, nPos As Long, qd As New rdoQuery, nLanc As Integer
Dim s As Integer, lResult As Long, sDataSus As String, sContador As String, x As Integer, sAtividade As String, sSimples As String, sVig As String, sFixo As String, sVariavel As String, sEstimado As String
Dim nQtdePagtoTotal As Integer, nQtdePagtoNao As Integer, nQtdePagtoSim As Integer, sISSEletr As String, sAtivIss As String
Dim Sql2 As String, bMei As Boolean

If Not Valida Then Exit Sub
ReDim aEmpresaISS(0)
sCodEmpresaIss = ""

lblTot.Caption = 0
'Set cn2 = en.OpenConnection(dsname:="odbcTributacao", Prompt:=rdDriverNoPrompt, Connect:="uid=" & UL & ";PWD=" & UP & ";driver={SQL Server};")
   Conn$ = "UID=" & UL & ";PWD=" & UP & ";" _
    & "DATABASE=tributacao;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cn2 = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

Ocupado

If chkAtivIss.value = vbUnchecked Then
    sAtivIss = ""
    
    For x = 0 To lstDDList4.ListCount - 1
        If lstDDList4.Selected(x) = True Then
            sAtivIss = sAtivIss & lstDDList4.ItemData(x) & ","
            
        End If
    Next
    If sAtivIss = "" Then
        Liberado
        MsgBox "Selecione ao menos uma atividade de ISS.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sAtivIss = Chomp(sAtivIss, chomp_righT, 1)
    
    Sql2 = "select codmobiliario,codatividade from mobiliarioatividadeiss where codmobiliario>100000 and codatividade in (" & sAtivIss & ")"
    Set RdoAux2 = cn2.OpenResultset(Sql2, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Do Until .EOF
            sCodEmpresaIss = sCodEmpresaIss & !codmobiliario & ","
           .MoveNext
        Loop
       .Close
    End With
    sCodEmpresaIss = Chomp(sCodEmpresaIss, chomp_righT, 1)
End If

Sql2 = "select codmobiliario,codatividade from mobiliarioatividadeiss where codmobiliario>100000 and codmobiliario<300000"
Set RdoAux2 = cn2.OpenResultset(Sql2, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
        ReDim Preserve aEmpresaISS(UBound(aEmpresaISS) + 1)
        aEmpresaISS(UBound(aEmpresaISS)).nCodigo = !codmobiliario
        aEmpresaISS(UBound(aEmpresaISS)).sAtividade = !codatividade
       .MoveNext
    Loop
   .Close
End With

If cGetInputState() <> 0 Then DoEvents
nTotal = 0: ReDim aCodigos(0): nPos = 1
Sql = "SELECT DISTINCT * FROM vwFULLEMPRESA3 WHERE "
If Val(lblTotImp.Caption) = 0 Then
    If sCodEmpresaIss = "" Then
        Sql = Sql & "CODIGOMOB BETWEEN 100000 AND 400000 "
    Else
        Sql = Sql & "CODIGOMOB in (" & sCodEmpresaIss & ")"
    End If
Else
    Sql = Sql & "CODIGOMOB in (" & strCodigos & ")"
End If

'If cmbMEI.Text = "Sim" Then
'    Sql = Sql & " AND MEI=1"
'ElseIf cmbMEI.Text = "Não" Then
'    Sql = Sql & " AND (MEI=0 OR MEI IS NULL)"
'End If
If cmbDispensaIE.Text = "Sim" Then
    Sql = Sql & " AND NOT DISPENSAIEDATA IS NULL"
ElseIf cmbDispensaIE.Text = "Não" Then
    Sql = Sql & " AND (DISPENSAIEDATA IS NULL)"
End If
If cmbIsento.Text = "Sim" Then
    Sql = Sql & " AND ISENTOTAXA=1"
ElseIf cmbIsento.Text = "Não" Then
    Sql = Sql & " AND (ISENTOTAXA=0 OR ISENTOTAXA IS NULL)"
End If
If cmbAlvara.Text = "Sim" Then
    Sql = Sql & " AND ALVARA=1"
ElseIf cmbAlvara.Text = "Não" Then
    Sql = Sql & " AND (ALVARA=0 OR ALVARA IS NULL)"
End If
If cmbTipo.Text = "Física" Then
    Sql = Sql & " AND (CONVERT(bigint, CPF) > 0)"
ElseIf cmbTipo.Text = "Jurídica" Then
    Sql = Sql & " AND (CONVERT(bigint, CNPJ) > 0)"
ElseIf cmbTipo.Text = "Indefinido" Then
    Sql = Sql & " AND (CONVERT(bigint, CNPJ) = 0) AND (CONVERT(bigint, CPF) = 0)"
End If
If chkDAbe.value = vbUnchecked Then
    Sql = Sql & " AND DATAABERTURA BETWEEN '" & Format(mskDataAbeIni.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataAbeFim.Text, "mm/dd/yyyy") & "'"
End If
If cmbDEnc.ListIndex = 1 Then
    Sql = Sql & " AND DATAENCERRAMENTO BETWEEN '" & Format(mskDataEncIni.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataEncFim.Text, "mm/dd/yyyy") & "'"
ElseIf cmbDEnc.ListIndex = 2 Then
    Sql = Sql & " AND DATAENCERRAMENTO IS NULL"
End If

'Sql = Sql & " AND EMITENF=1"
If chkContador.value = vbUnchecked Then
    sContador = ""
    For x = 0 To lstDDList2.ListCount - 1
        If lstDDList2.Selected(x) = True Then
            sContador = sContador & lstDDList2.ItemData(x) & ","
        End If
    Next
    If sContador = "" Then
        Liberado
        MsgBox "Selecione ao menos um contador.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sContador = Chomp(sContador, chomp_righT, 1)
    
    Sql = Sql & " AND RESPCONTABIL in (" & sContador & ")"
End If
If chkAtividade.value = vbUnchecked Then
    sAtividade = ""
    For x = 0 To lstDDList3.ListCount - 1
        If lstDDList3.Selected(x) = True Then
            sAtividade = sAtividade & lstDDList3.ItemData(x) & ","
        End If
    Next
    If sAtividade = "" Then
        Liberado
        MsgBox "Selecione ao menos uma atividade.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sAtividade = Chomp(sAtividade, chomp_righT, 1)
    
    Sql = Sql & " AND CODATIVIDADE in (" & sAtividade & ")"
'Else
End If

   If chkTipoAtiv(1).value = vbChecked Then
'        Sql = Sql & " AND (CODATIVIDADE >=10000 AND CODATIVIDADE<20000) "
    End If
    If chkTipoAtiv(2).value = vbChecked Then
'        Sql = Sql & " AND (CODATIVIDADE >=20000 AND CODATIVIDADE<30000) "
    End If
    If chkTipoAtiv(3).value = vbChecked Then
'        Sql = Sql & " AND (CODATIVIDADE >=30000 AND CODATIVIDADE<40000) "
    End If
    If chkTipoAtiv(4).value = vbChecked Then
'        Sql = Sql & " AND (CODATIVIDADE >=40000 AND CODATIVIDADE<50000) "
    End If

If cmbVistoria.Text = "Sim" Then
    Sql = Sql & " AND VISTORIA=1"
ElseIf cmbVistoria.Text = "Não" Then
    Sql = Sql & " AND (VISTORIA=0 OR VISTORIA IS NULL)"
End If

Sql = Sql & " ORDER BY CODIGOMOB"
cmdConsultar.Enabled = False
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
        End If
        
    '    If (!codatividade >= 10000 And !codatividade < 20000) And chkTipoAtiv(1).value = vbUnchecked Then GoTo proximo
    '    If (!codatividade >= 20000 And !codatividade < 30000) And chkTipoAtiv(2).value = vbUnchecked Then GoTo proximo
    '    If (!codatividade >= 30000 And !codatividade < 40000) And chkTipoAtiv(3).value = vbUnchecked Then GoTo proximo
     '   If (!codatividade >= 40000 And !codatividade < 50000) And chkTipoAtiv(4).value = vbUnchecked Then GoTo proximo
        
  '      If Val(!codatividade) = 0 And (chkTipoAtiv(1) = vbChecked Or chkTipoAtiv(2) = vbChecked Or chkTipoAtiv(3) = vbChecked Or chkTipoAtiv(4) = vbChecked) Then
  '          GoTo Proximo
  '      End If
        
        sVig = "N"
        lResult = BinarySearchLong(aVigilancia(), !codigomob)
        If lResult > -1 Then
           sVig = "S"
        End If
        If cmbVSanit.ListIndex = 1 And sVig = "N" Then GoTo proximo
        If cmbVSanit.ListIndex = 2 And sVig = "N" Then GoTo proximo
        
        bMei = False
        'If !codigomob = 100671 Then MsgBox "teste"
        If cmbMEI.Text = "Sim" Then
            bMei = IsMEI(!codigomob)
            If Not bMei Then GoTo proximo
        ElseIf cmbMEI.Text = "Não" Then
            bMei = IsMEI(!codigomob)
            If bMei Then GoTo proximo
        End If
        bMei = IsMEI(!codigomob)
        
        sSimples = "N"
        On Error Resume Next
        RdoAux2.Close
        On Error GoTo 0
        Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNASN(" & Format(Val(!codigomob), "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
        Set RdoAux2 = cn2.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
        If RdoAux2!RETORNO = 1 And cmbSimples.Text = "Não" Then
            GoTo proximo
        ElseIf RdoAux2!RETORNO = 0 And cmbSimples.Text = "Sim" Then
            GoTo proximo
        Else
            If RdoAux2!RETORNO = 1 Then
                sSimples = "S"
            End If
        End If
                
        sISSEletr = "N"
        On Error Resume Next
        RdoAux2.Close
        On Error GoTo 0
        Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNAIE(" & Format(Val(!codigomob), "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
        Set RdoAux2 = cn2.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
        If RdoAux2!RETORNO = 1 And cmbISSE.Text = "Não" Then
            GoTo proximo
        ElseIf RdoAux2!RETORNO = 0 And cmbISSE.Text = "Sim" Then
            GoTo proximo
        Else
            If RdoAux2!RETORNO = 1 Then
                sISSEletr = "S"
            End If
        End If
                
        lResult = BinarySearchLong(aFixo(), !codigomob)
        If lResult = -1 Then
            sFixo = "N"
        Else
            sFixo = "S"
        End If
        lResult = BinarySearchLong(aEstimado(), !codigomob)
        If lResult = -1 Then
            sEstimado = "N"
        Else
            sEstimado = "S"
        End If
        lResult = BinarySearchLong(aVariavel(), !codigomob)
        If lResult = -1 Then
            sVariavel = "N"
        Else
            sVariavel = "S"
        End If
        
        If cmbISS.ListIndex = 1 And sFixo = "N" Then GoTo proximo
        If cmbISS.ListIndex = 2 And sVariavel = "N" Then GoTo proximo
        If cmbISS.ListIndex = 3 And sEstimado = "N" Then GoTo proximo
        If cmbISS.ListIndex = 4 And (sEstimado = "N" And sFixo = "N" And sVariavel = "N") Then GoTo proximo
        
        If cmbDSus.ListIndex > -1 Then
            sDataSus = ""
            If cmbDSus.ListIndex = 2 Then 'funcionando
                lResult = BinarySearchLong(aSuspensoCod(), !codigomob)
                If lResult > -1 Then
                   GoTo proximo
                End If
            ElseIf cmbDSus.ListIndex = 0 Then 'todos
                lResult = BinarySearchLong(aSuspensoCod(), !codigomob)
                If lResult > -1 Then
                    For s = 1 To UBound(aSuspenso)
                        If aSuspenso(s).nCodigo = !codigomob Then
                            sDataSus = Format(aSuspenso(s).dData, "dd/mm/yyyy")
                            Exit For
                        End If
                    Next
                End If
            Else
                lResult = BinarySearchLong(aSuspensoCod(), !codigomob)
                If lResult = -1 Then
                   GoTo proximo
                Else
                    For s = 1 To UBound(aSuspenso)
                        If aSuspenso(s).nCodigo = !codigomob Then
                            sDataSus = Format(aSuspenso(s).dData, "dd/mm/yyyy")
                            If mskDataSusIni.ClipText = "" Then mskDataSusIni.Text = "01/01/2000"
                            If mskDataSusFim.ClipText = "" Then mskDataSusFim.Text = Format(Now, "dd/mm/yyyy")
                            If (CDate(sDataSus) < CDate(mskDataSusIni.Text)) Or (CDate(sDataSus) > CDate(mskDataSusFim.Text)) Then
                                GoTo proximo
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
        
        
        nQtdePagtoTotal = 0: nQtdePagtoNao = 0: nQtdePagtoSim = 0
        If cmbPagto.ListIndex > 0 Then
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux2.Close
            On Error GoTo 0
            qd.Sql = "{ Call spRETORNAPARCELASPAGAS(?,?,?) }"
            qd(0) = !codigomob
            qd(1) = Val(txtAno.Text)
            qd(2) = cmbImposto.ItemData(cmbImposto.ListIndex)
            Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
            With RdoAux2
                nQtdePagtoTotal = .rdoColumns(0)
                nQtdePagtoSim = .rdoColumns(1)
                nQtdePagtoNao = .rdoColumns(2)
               .Close
            End With
            
'            If !CODIGOMOB = 116625 Then MsgBox "teste"
            
            If cmbPagto.ListIndex = 4 Then 'Não possui imposto
                If nQtdePagtoTotal > 0 Then
                    GoTo proximo
                End If
            ElseIf cmbPagto.ListIndex = 1 Then 'Nenhuma paga
                If nQtdePagtoSim > 0 Or nQtdePagtoTotal = 0 Then
                    GoTo proximo
                End If
            ElseIf cmbPagto.ListIndex = 2 Then 'Algumas pagas
                If nQtdePagtoSim = 0 Or nQtdePagtoNao = 0 Or nQtdePagtoTotal = 0 Then
                    GoTo proximo
                End If
            ElseIf cmbPagto.ListIndex = 3 Then 'Todas pagas
                If nQtdePagtoNao > 0 Or nQtdePagtoTotal = 0 Then
                    GoTo proximo
                End If
            End If
        End If
        
        
        sCodEmpresaIss = ""
        For s = 1 To UBound(aEmpresaISS)
            If aEmpresaISS(s).nCodigo = !codigomob Then
                sCodEmpresaIss = sCodEmpresaIss & aEmpresaISS(s).sAtividade & ","
            End If
        Next
        
        sCodEmpresaIss = Chomp(sCodEmpresaIss, chomp_righT, 1)
        
        ReDim Preserve aCodigos(UBound(aCodigos) + 1)
        aCodigos(UBound(aCodigos)).nCodigo = !codigomob
        aCodigos(UBound(aCodigos)).sRazao = !razaosocial
        aCodigos(UBound(aCodigos)).sNomeFantasia = SubNull(!NOMEFANTASIA)
        aCodigos(UBound(aCodigos)).sCNPJ = SubNull(!Cnpj)
        aCodigos(UBound(aCodigos)).sCPF = SubNull(!CPF)
        aCodigos(UBound(aCodigos)).sIE = SubNull(!INSCESTADUAL)
        aCodigos(UBound(aCodigos)).sEndereco = Trim(SubNull(!Logradouro) & " nº " & Val(SubNull(!Numero)))
        aCodigos(UBound(aCodigos)).sComplemento = SubNull(!Complemento)
        aCodigos(UBound(aCodigos)).sBairro = SubNull(!DescBairro)
        aCodigos(UBound(aCodigos)).sCidade = SubNull(!descCidade)
        aCodigos(UBound(aCodigos)).sUF = SubNull(!SiglaUF)
        aCodigos(UBound(aCodigos)).sCep = RetornaCEP(!CodLogradouro, !Numero)
        aCodigos(UBound(aCodigos)).sDataAbertura = Format(!DATAABERTURA, "dd/mm/yyyy")
        aCodigos(UBound(aCodigos)).sDataEncerramento = Format(!dataencerramento, "dd/mm/yyyy")
        aCodigos(UBound(aCodigos)).sDataSuspensao = sDataSus
        aCodigos(UBound(aCodigos)).sAtividade = SubNull(!Atividade)
        aCodigos(UBound(aCodigos)).sAtivExtenso = SubNull(!ativextenso)
        aCodigos(UBound(aCodigos)).sHorario = SubNull(!DESCHORARIO)
        aCodigos(UBound(aCodigos)).sContador = SubNull(!NOMEESC)
        aCodigos(UBound(aCodigos)).sMEI = IIf(bMei, "S", "N")
        aCodigos(UBound(aCodigos)).sSN = sSimples
        aCodigos(UBound(aCodigos)).sIsentoTaxa = IIf(!ISENTOTAXA = 1, "S", "N")
        aCodigos(UBound(aCodigos)).sAlvara = IIf(!ALVARA = 1, "S", "N")
        aCodigos(UBound(aCodigos)).sVistoria = IIf(!VISTORIA = 1, "S", "N")
        aCodigos(UBound(aCodigos)).sVigilancia = sVig
        aCodigos(UBound(aCodigos)).sFixo = sFixo
        aCodigos(UBound(aCodigos)).sVariavel = sVariavel
        aCodigos(UBound(aCodigos)).sEstimado = sEstimado
        aCodigos(UBound(aCodigos)).sProcEncerramento = SubNull(!NUMPROCENCERRAMENTO)
        aCodigos(UBound(aCodigos)).sNomeContato = SubNull(!NOMECONTATO)
        aCodigos(UBound(aCodigos)).sFoneContato = SubNull(!fonecontato)
        aCodigos(UBound(aCodigos)).sProcAbertura = SubNull(!numprocesso)
        aCodigos(UBound(aCodigos)).nParcPagtoTotal = nQtdePagtoTotal
        aCodigos(UBound(aCodigos)).nParcPagtoSim = nQtdePagtoSim
        aCodigos(UBound(aCodigos)).nParcPagtoNao = nQtdePagtoNao
        aCodigos(UBound(aCodigos)).sISSElet = sISSEletr
        aCodigos(UBound(aCodigos)).sDispensaIE = IIf(IsNull(!DISPENSAIEDATA), "N", "S")
        aCodigos(UBound(aCodigos)).sAtivIss = sCodEmpresaIss
        
        
        xImovel.RetornaEndereco !codigomob, Mobiliario, Entrega
        aCodigos(UBound(aCodigos)).sEnderecoEnt = xImovel.Endereco & " nº " & xImovel.Numero
        aCodigos(UBound(aCodigos)).sComplementoEnt = xImovel.Complemento
        aCodigos(UBound(aCodigos)).sBairroEnt = xImovel.Bairro
        aCodigos(UBound(aCodigos)).sCidadeEnt = xImovel.Cidade
        aCodigos(UBound(aCodigos)).sUFEnt = xImovel.UF
        aCodigos(UBound(aCodigos)).sCEPEnt = xImovel.Cep
        
        
        
        
        
        nTotal = nTotal + 1
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
    lblTot.Caption = nTotal
   .Close
End With
PBar.value = 0: PBar.Color = vbWhite
cn2.Close
cmdConsultar.Enabled = True
Liberado
If nTotal = 0 Then
    MsgBox "Nenhuma empresa possui os critérios especificados.", vbInformation, "Informação"
End If

End Sub


Private Sub CarregaCriterio()
Dim x As Integer

For x = 1990 To Year(Now)
    cmbExercicio.AddItem (x)
Next

cmbExercicio.Text = Year(Now)

cmbMEI.AddItem "Todos"
cmbMEI.AddItem "Sim"
cmbMEI.AddItem "Não"
cmbMEI.ListIndex = 0

cmbSimples.AddItem "Todos"
cmbSimples.AddItem "Sim"
cmbSimples.AddItem "Não"
cmbSimples.ListIndex = 0

cmbIsento.AddItem "Todos"
cmbIsento.AddItem "Sim"
cmbIsento.AddItem "Não"
cmbIsento.ListIndex = 0

cmbAlvara.AddItem "Todos"
cmbAlvara.AddItem "Sim"
cmbAlvara.AddItem "Não"
cmbAlvara.ListIndex = 0

cmbTipo.AddItem "Todos"
cmbTipo.AddItem "Física"
cmbTipo.AddItem "Jurídica"
cmbTipo.AddItem "Indefinido"
cmbTipo.ListIndex = 0

cmbDEnc.AddItem "Todos"
cmbDEnc.AddItem "Encerradas"
cmbDEnc.AddItem "Abertas"
cmbDEnc.ListIndex = 2

cmbDSus.AddItem "Todos"
cmbDSus.AddItem "Suspensas"
cmbDSus.AddItem "Funcionando"
cmbDSus.ListIndex = 2

cmbVistoria.AddItem "Todos"
cmbVistoria.AddItem "Sim"
cmbVistoria.AddItem "Não"
cmbVistoria.ListIndex = 0

cmbVSanit.AddItem "Todos"
cmbVSanit.AddItem "Sim"
cmbVSanit.AddItem "Não"
cmbVSanit.ListIndex = 0

cmbISS.AddItem "Indiferente"
cmbISS.AddItem "Fixo"
cmbISS.AddItem "Variável"
cmbISS.AddItem "Estimado"
cmbISS.AddItem "Prest.Serviço"
cmbISS.ListIndex = 0

cmbDispensaIE.ListIndex = 0
cmbISSE.ListIndex = 0

End Sub


Private Sub CarregaLista()
Dim sDoc As String
Dim x As Integer
Ocupado
grdMain.Redraw = False
grdMain.Clear
grdMain.Redraw = True
grdMain.Redraw = False

For x = 1 To UBound(aCodigos)
    If cGetInputState() <> 0 Then DoEvents
    grdMain.AddRow
    grdMain.CellDetails grdMain.Rows, 1, Format(aCodigos(x).nCodigo, "000000"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 2, aCodigos(x).sRazao
    sDoc = ""
    If Val(SubNull(aCodigos(x).sCNPJ)) > 0 Then
        sDoc = Format(aCodigos(x).sCNPJ, "0#\.###\.###/####-##")
    End If
    If Val(SubNull(aCodigos(x).sCPF)) > 0 Then
        sDoc = Format(aCodigos(x).sCPF, "00#\.###\.###-##")
    End If
    grdMain.CellDetails grdMain.Rows, 3, sDoc
    grdMain.CellDetails grdMain.Rows, 4, aCodigos(x).sIE
    grdMain.CellDetails grdMain.Rows, 5, aCodigos(x).sEndereco
    grdMain.CellDetails grdMain.Rows, 6, aCodigos(x).sComplemento
    grdMain.CellDetails grdMain.Rows, 7, aCodigos(x).sBairro
    grdMain.CellDetails grdMain.Rows, 8, aCodigos(x).sCidade
    grdMain.CellDetails grdMain.Rows, 9, aCodigos(x).sUF, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 10, aCodigos(x).sCep, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 11, aCodigos(x).sDataAbertura, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 12, aCodigos(x).sDataEncerramento, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 13, aCodigos(x).sDataSuspensao, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 14, aCodigos(x).sAtividade
    grdMain.CellDetails grdMain.Rows, 15, aCodigos(x).sAtivExtenso
    grdMain.CellDetails grdMain.Rows, 16, aCodigos(x).sHorario
    grdMain.CellDetails grdMain.Rows, 17, aCodigos(x).sContador
    grdMain.CellDetails grdMain.Rows, 18, aCodigos(x).sMEI, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 19, aCodigos(x).sSN, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 20, aCodigos(x).sIsentoTaxa, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 21, aCodigos(x).sAlvara, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 22, aCodigos(x).sVistoria, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 23, aCodigos(x).sVigilancia, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 24, aCodigos(x).sFixo, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 25, aCodigos(x).sVariavel, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 26, aCodigos(x).sEstimado, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 27, aCodigos(x).sProcEncerramento
    grdMain.CellDetails grdMain.Rows, 28, aCodigos(x).sNomeContato
    grdMain.CellDetails grdMain.Rows, 29, aCodigos(x).sFoneContato
    grdMain.CellDetails grdMain.Rows, 30, aCodigos(x).sProcAbertura
    grdMain.CellDetails grdMain.Rows, 31, aCodigos(x).nParcPagtoTotal, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 32, aCodigos(x).nParcPagtoSim, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 33, aCodigos(x).nParcPagtoNao, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 34, aCodigos(x).sISSElet, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 35, aCodigos(x).sDispensaIE, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 36, aCodigos(x).sNomeFantasia
    grdMain.CellDetails grdMain.Rows, 37, aCodigos(x).sAtivIss
    grdMain.CellDetails grdMain.Rows, 38, ""
    grdMain.CellDetails grdMain.Rows, 39, aCodigos(x).sEnderecoEnt
    grdMain.CellDetails grdMain.Rows, 40, aCodigos(x).sComplementoEnt
    grdMain.CellDetails grdMain.Rows, 41, aCodigos(x).sBairroEnt
    grdMain.CellDetails grdMain.Rows, 42, aCodigos(x).sCidadeEnt
    grdMain.CellDetails grdMain.Rows, 43, aCodigos(x).sUFEnt, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 44, aCodigos(x).sCEPEnt, DT_CENTER
proximo:
Next
Liberado
grdMain.Redraw = True
End Sub

Private Sub CarregaCampos()
Dim x As Integer
With lstDDList1
    .Clear
    .AddItem "Codigo"
    .ItemData(.NewIndex) = 0
    .AddItem "Razao Social"
    .ItemData(.NewIndex) = 1
    .AddItem "CPF/CNPJ"
    .ItemData(.NewIndex) = 2
    .AddItem "Inscricao Estadual"
    .ItemData(.NewIndex) = 3
    .AddItem "Endereco"
    .ItemData(.NewIndex) = 4
    .AddItem "Complemento"
    .ItemData(.NewIndex) = 5
    .AddItem "Bairro"
    .ItemData(.NewIndex) = 6
    .AddItem "Cidade"
    .ItemData(.NewIndex) = 7
    .AddItem "UF"
    .ItemData(.NewIndex) = 8
    .AddItem "CEP"
    .ItemData(.NewIndex) = 9
    .AddItem "Data de Abertura"
    .ItemData(.NewIndex) = 10
    .AddItem "Data de Encerramento"
    .ItemData(.NewIndex) = 11
    .AddItem "Data de Suspensao"
    .ItemData(.NewIndex) = 12
    .AddItem "Atividade"
    .ItemData(.NewIndex) = 13
    .AddItem "Atividade por Extenso"
    .ItemData(.NewIndex) = 14
    .AddItem "Horario de Funcionamento"
    .ItemData(.NewIndex) = 15
    .AddItem "Escritorio Contabil"
    .ItemData(.NewIndex) = 16
    .AddItem "Inscrito no MEI"
    .ItemData(.NewIndex) = 17
    .AddItem "Simples Nacional"
    .ItemData(.NewIndex) = 18
    .AddItem "Isento Taxa"
    .ItemData(.NewIndex) = 19
    .AddItem "Alvara Automatico"
    .ItemData(.NewIndex) = 20
    .AddItem "Vistoria"
    .ItemData(.NewIndex) = 21
    .AddItem "Vigilancia Sanitaria"
    .ItemData(.NewIndex) = 22
    .AddItem "ISS Fixo"
    .ItemData(.NewIndex) = 23
    .AddItem "ISS Variável"
    .ItemData(.NewIndex) = 24
    .AddItem "ISS Estimado"
    .ItemData(.NewIndex) = 25
    .AddItem "Processo Encerramento"
    .ItemData(.NewIndex) = 26
    .AddItem "Nome do Contato"
    .ItemData(.NewIndex) = 27
    .AddItem "Telefone de Contato"
    .ItemData(.NewIndex) = 28
    .AddItem "Processo de Abertura"
    .ItemData(.NewIndex) = 29
    .AddItem "Parc.Pagas Total"
    .ItemData(.NewIndex) = 30
    .AddItem "Parc.Pagas Sim"
    .ItemData(.NewIndex) = 31
    .AddItem "Parc.Pagas Não"
    .ItemData(.NewIndex) = 32
    .AddItem "ISS Eletrônico"
    .ItemData(.NewIndex) = 33
    .AddItem "Dispensa ISS Elet."
    .ItemData(.NewIndex) = 34
    .AddItem "Nome Fantasia"
    .ItemData(.NewIndex) = 35
    .AddItem "Atividade ISS"
    .ItemData(.NewIndex) = 36
    .AddItem "Auto Incremento"
    .ItemData(.NewIndex) = 37
    .AddItem "Endereco Entrega"
    .ItemData(.NewIndex) = 38
    .AddItem "Complemento Entrega"
    .ItemData(.NewIndex) = 39
    .AddItem "Bairro Entrega"
    .ItemData(.NewIndex) = 40
    .AddItem "Cidade Entrega"
    .ItemData(.NewIndex) = 41
    .AddItem "UF Entrega"
    .ItemData(.NewIndex) = 42
    .AddItem "CEP Entrega"
    .ItemData(.NewIndex) = 43
    .Selected(0) = True
    .Selected(1) = True
    .Selected(2) = True
    .Selected(4) = True
End With

For x = 0 To lstDDList1.ListCount - 1
    cmbCampos.AddItem lstDDList1.List(x)
Next
cmbCampos.ListIndex = 0

End Sub

Private Sub GridHeader()

With grdMain
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "kCodigo006", "Código", ecgHdrTextALignCentre, , 50
    .AddColumn "kRazao080", "Razão Social", ecgHdrTextALignLeft, , 200
    .AddColumn "kCnpj018", "CPF/CNPJ", ecgHdrTextALignLeft, , 120
    .AddColumn "kIE015", "Insc.Est.", ecgHdrTextALignLeft, , 100
    .AddColumn "kEndereco100", "Endereço", ecgHdrTextALignLeft, , 200
    .AddColumn "kCompl030", "Complemento", ecgHdrTextALignLeft, , 90
    .AddColumn "kBairro050", "Bairro", ecgHdrTextALignLeft, , 120
    .AddColumn "kCidade050", "Cidade", ecgHdrTextALignLeft, , 120
    .AddColumn "kUF002", "UF", ecgHdrTextALignCentre, , 35
    .AddColumn "kCEP009", "CEP", ecgHdrTextALignCentre, , 70
    .AddColumn "kDABE010", "Dt.Abe.", ecgHdrTextALignCentre, , 70
    .AddColumn "kDENC010", "Dt.Enc.", ecgHdrTextALignCentre, , 70
    .AddColumn "kDSUS010", "Dt.Sus.", ecgHdrTextALignCentre, , 70
    .AddColumn "kAtividade100", "Atividade", ecgHdrTextALignLeft, , 200
    .AddColumn "kAtivExt100", "Ativ.Ext.", ecgHdrTextALignLeft, , 200
    .AddColumn "kHorario040", "Horario", ecgHdrTextALignLeft, , 120
    .AddColumn "kContador040", "Contador", ecgHdrTextALignLeft, , 120
    .AddColumn "kMEI001", "MEI", ecgHdrTextALignCentre, , 50
    .AddColumn "kSNA001", "SNA", ecgHdrTextALignCentre, , 50
    .AddColumn "kIST001", "IST", ecgHdrTextALignCentre, , 50
    .AddColumn "kAAU001", "AAU", ecgHdrTextALignCentre, , 50
    .AddColumn "kVIS001", "VIS", ecgHdrTextALignCentre, , 50
    .AddColumn "kVIG001", "VIG", ecgHdrTextALignCentre, , 50
    .AddColumn "kISF001", "Fix", ecgHdrTextALignCentre, , 50
    .AddColumn "kISV001", "Var", ecgHdrTextALignCentre, , 50
    .AddColumn "kISE001", "Est", ecgHdrTextALignCentre, , 50
    .AddColumn "kPENC015", "Proc.Enc.", ecgHdrTextALignLeft, , 90
    .AddColumn "kNOMEC040", "Nome Contato", ecgHdrTextALignLeft, , 90
    .AddColumn "kFONEC040", "Fone Contato", ecgHdrTextALignLeft, , 90
    .AddColumn "kPABE015", "Proc.Abe.", ecgHdrTextALignLeft, , 90
    .AddColumn "kPGT006", "Pg.Total", ecgHdrTextALignRight, , 50
    .AddColumn "kPGS006", "Pg.Sim", ecgHdrTextALignRight, , 50
    .AddColumn "kPGN006", "Pg.Não", ecgHdrTextALignRight, , 50
    .AddColumn "kISO001", "ISE", ecgHdrTextALignCentre, , 50
    .AddColumn "kDIO001", "DIE", ecgHdrTextALignCentre, , 50
    .AddColumn "kFAN001", "Nome Fantasia", ecgHdrTextALignLeft, , 120
    .AddColumn "kATI001", "Atividade ISS", ecgHdrTextALignLeft, , 90
    .AddColumn "kAUI001", "Seq.", ecgHdrTextALignRight, , 50
    .AddColumn "kEnderecoE100", "Endereço Entrega", ecgHdrTextALignLeft, , 200
    .AddColumn "kComplE030", "Complemento Entrega", ecgHdrTextALignLeft, , 90
    .AddColumn "kBairroE050", "Bairro Entrega", ecgHdrTextALignLeft, , 120
    .AddColumn "kCidadeE050", "Cidade Entrega", ecgHdrTextALignLeft, , 120
    .AddColumn "kUFE002", "UF Entrega", ecgHdrTextALignCentre, , 35
    .AddColumn "kCEPE009", "CEP Entrega", ecgHdrTextALignCentre, , 70
End With

End Sub

Private Sub Init()
Dim Sql As String, RdoAux As rdoResultset, x As Integer

sTitulo = "Editor de Cartas de Correspondência - "

For i = 0 To Screen.FontCount - 1
    cmbFonte.AddItem Screen.Fonts(i)
    If Screen.Fonts(i) = "Arial" Then x = i
Next i
For i = 8 To 30 Step 2
    cmbTam.AddItem i
Next
cmbFonte.ListIndex = x
cmbTam.ListIndex = 0

ReDim aSuspenso(0): ReDim aSuspensoCod(0)
Sql = "SELECT codmobiliario, DataEv, codtipoevento From vwMOBILIARIOSUSPENSO Where (codtipoevento = 2) order by codmobiliario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aSuspenso(UBound(aSuspenso) + 1)
        ReDim Preserve aSuspensoCod(UBound(aSuspensoCod) + 1)
        aSuspensoCod(UBound(aSuspensoCod)) = !codmobiliario
        aSuspenso(UBound(aSuspenso)).nCodigo = !codmobiliario
        aSuspenso(UBound(aSuspenso)).dData = Format(!DATAEV, "dd/mm/yyyy")
       .MoveNext
    Loop
   .Close
End With

ReDim aVigilancia(0)
'Sql = "SELECT DISTINCT codmobiliario From mobiliarioatividadevs2 ORDER BY codmobiliario"
Sql = "SELECT DISTINCT codigo From mobiliariovs ORDER BY codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aVigilancia(UBound(aVigilancia) + 1)
        'aVigilancia(UBound(aVigilancia)) = !codmobiliario
        aVigilancia(UBound(aVigilancia)) = !Codigo
       .MoveNext
    Loop
   .Close
End With

ReDim aFixo(0)
Sql = "SELECT DISTINCT codmobiliario From mobiliarioatividadeiss where codtributo=11 ORDER BY codmobiliario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aFixo(UBound(aFixo) + 1)
        aFixo(UBound(aFixo)) = !codmobiliario
       .MoveNext
    Loop
   .Close
End With

ReDim aVariavel(0)
Sql = "SELECT DISTINCT codmobiliario From mobiliarioatividadeiss where codtributo=13 ORDER BY codmobiliario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aVariavel(UBound(aVariavel) + 1)
        aVariavel(UBound(aVariavel)) = !codmobiliario
       .MoveNext
    Loop
   .Close
End With

ReDim aEstimado(0)
Sql = "SELECT DISTINCT codmobiliario From mobiliarioatividadeiss where codtributo=12 ORDER BY codmobiliario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aEstimado(UBound(aEstimado) + 1)
        aEstimado(UBound(aEstimado)) = !codmobiliario
       .MoveNext
    Loop
   .Close
End With


Sql = "SELECT * FROM ESCRITORIOCONTABIL WHERE CODIGOESC>0  ORDER BY NOMEESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstDDList2.AddItem !NOMEESC
        lstDDList2.ItemData(lstDDList2.NewIndex) = !codigoesc
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT CODATIVIDADE,DESCATIVIDADE FROM ATIVIDADE WHERE CODATIVIDADE>0 ORDER BY DESCATIVIDADE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstDDList3.AddItem !descatividade
        lstDDList3.ItemData(lstDDList3.NewIndex) = !codatividade
       .MoveNext
    Loop
   .Close
End With

CarregaListaIss

With cmbImposto
    .AddItem (" ")
    .ItemData(.NewIndex) = 0
    .AddItem ("Taxa de Licença")
    .ItemData(.NewIndex) = 6
    .AddItem ("ISS Fixo")
    .ItemData(.NewIndex) = 14
    .AddItem ("ISS Estimado")
    .ItemData(.NewIndex) = 3
    .AddItem ("ISS Variável")
    .ItemData(.NewIndex) = 5
    .AddItem ("ISS Fixo/TLL")
    .ItemData(.NewIndex) = 2
    .AddItem ("Vigilância Sanitária")
    .ItemData(.NewIndex) = 13
End With

mskDataAbeIni.BackColor = Kde
mskDataAbeFim.BackColor = Kde
mskDataAbeIni.Locked = True
mskDataAbeFim.Locked = True
cmbPagto.ListIndex = 0

End Sub

Private Function Valida() As Boolean
lblTot.Caption = 0
Valida = False

If chkDAbe.value = vbUnchecked Then
    If Not IsDate(mskDataAbeIni.Text) Or Not IsDate(mskDataAbeFim.Text) Then
        MsgBox "Data de abertura inicial/final inválida.", vbCritical, "Erro de validação"
        Exit Function
    End If
    If CDate(mskDataAbeIni.Text) > CDate(mskDataAbeFim.Text) Then
        MsgBox "Data de abertura inicial maior que data final.", vbCritical, "Erro de validação"
        Exit Function
    End If
End If

If cmbDEnc.ListIndex = 1 Then
    If Not IsDate(mskDataEncIni.Text) Or Not IsDate(mskDataEncFim.Text) Then
        MsgBox "Data de encerramento inicial/final inválida.", vbCritical, "Erro de validação"
        Exit Function
    End If
    If CDate(mskDataEncIni.Text) > CDate(mskDataEncFim.Text) Then
        MsgBox "Data de encerramento inicial maior que data final.", vbCritical, "Erro de validação"
        Exit Function
    End If
End If

If cmbDSus.ListIndex = 1 Then
    If Not IsDate(mskDataSusIni.Text) Or Not IsDate(mskDataSusFim.Text) Then
        MsgBox "Data de suspensão inicial/final inválida.", vbCritical, "Erro de validação"
        Exit Function
    End If
    If CDate(mskDataSusIni.Text) > CDate(mskDataSusFim.Text) Then
        MsgBox "Data de suspensão inicial maior que data final.", vbCritical, "Erro de validação"
        Exit Function
    End If
End If

If cmbPagto.ListIndex > 0 Then
    If cmbImposto.ListIndex < 1 Then
        MsgBox "Selecione o imposto a verificar.", vbCritical, "Erro de Validação"
        Exit Function
    Else
        If Val(txtAno.Text) < 1950 Or Val(txtAno.Text) > 2020 Then
            MsgBox "Ano de pagamento inválido.", vbCritical, "Erro de Validação"
            Exit Function
        End If
    End If
End If

Valida = True
End Function


Private Sub CarregaModelo()
Dim Sql As String, RdoAux As rdoResultset

lstModelo.Clear
Sql = "SELECT * FROM MMGMODELO ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstModelo.AddItem !Nome
       .MoveNext
    Loop
   .Close
End With

End Sub

Public Sub HighLight(Rtb As RichTextBox, lColor As Long)
'add new color to color table
'add tags \highlight# and \highlight0
'where # is new color number
Dim iPos As Long
Dim strRTF As String
Dim bkColor As Integer

    With Rtb
        iPos = .SelStart
        'bracket selection
        .SelText = Chr(&H9D) & .SelText & Chr(&H81)
        strRTF = Rtb.TextRTF
'add new color
        bkColor = AddColorToTable(strRTF, lColor)
'add highlighting
         strRTF = Replace(strRTF, "\'9d", "\up1\highlight" & CStr(bkColor) & "")
         strRTF = Replace(strRTF, "\'81", "\highlight0\up0 ")

         .TextRTF = strRTF
        .SelStart = iPos
       End With

End Sub

Function AddColorToTable(strRTF As String, lColor As Long) As Integer
Dim iPos As Long, jpos As Long

Dim ctbl As String
Dim tagColors
Dim nColors As Integer
Dim tagNew As String
Dim i As Integer
Dim iLen As Integer
Dim split1 As String
Dim split2 As String

    'make new color into tag
    tagNew = "\red" & CStr(lColor And &HFF) & _
        "\green" & CStr(Int(lColor / &H100) And &HFF) & _
        "\blue" & CStr(Int(lColor / &H10000))
    
    'find colortable
    iPos = InStr(strRTF, "{\colortbl")
    
    If iPos > 0 Then 'if table already exists
        jpos = InStr(iPos, strRTF, ";}")
        'color table
        ctbl = Mid(strRTF, iPos + 12, jpos - iPos - 12)
        'array of color tags
        tagColors = Split(ctbl, ";")
        nColors = UBound(tagColors) + 2
        'see if our color already exists in table
        For i = 0 To UBound(tagColors)
            If tagColors(i) = tagNew Then
                AddColorToTable = i + 1
                Exit Function
            End If
        Next i
'{\fonttbl{\f0\fnil\fcharset0 Verdana;}}
'{\colortbl ;\red0\green0\blue0;\red128\green0\blue255;}
        
        split1 = Left(strRTF, jpos)
        split2 = Mid(strRTF, jpos + 1)
        strRTF = split1 & tagNew & ";" & split2
        AddColorToTable = nColors
    
    Else
        'color table doesn't exists, let's make one
        iPos = InStr(strRTF, "{\fonttbl") 'beginning of font table
        jpos = InStr(iPos, strRTF, ";}}") + 2 'end of font table
        split1 = Left(strRTF, jpos)
        split2 = Mid(strRTF, jpos + 1)
        strRTF = split1 & "{\colortbl ;" & tagNew & ";}" & split2
        AddColorToTable = 1
    End If

End Function

Private Sub optAtiv_Click(Index As Integer)
CarregaListaIss
End Sub

Private Sub CarregaListaIss()
lstDDList4.Clear

Sql = "SELECT distinct atividadeiss.codatividade, atividadeiss.descatividade, tabelaiss.tipoiss "
Sql = Sql & "FROM atividadeiss INNER JOIN tabelaiss ON atividadeiss.codatividade = tabelaiss.codigoativ "

If cmbISS.Text = "Fixo" Then
    Sql = Sql & "where tipoiss=11"
ElseIf cmbISS.Text = "Variável" Then
    Sql = Sql & "where tipoiss=13"
ElseIf cmbISS.Text = "Estimado" Then
    Sql = Sql & "where tipoiss=12"
End If

If optAtiv(0).value = True Then
    Sql = Sql & "ORDER BY atividadeiss.CODATIVIDADE"
Else
    Sql = Sql & "ORDER BY DESCATIVIDADE"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstDDList4.AddItem !codatividade & " - " & !descatividade
        lstDDList4.ItemData(lstDDList4.NewIndex) = !codatividade
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub optIM_Click(Index As Integer)
If Index = 0 Then
    AtivaEmpresa (True)
Else
    AtivaEmpresa (False)
End If
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub AtivaEmpresa(bValue As Boolean)
Dim nColor As Long
If bValue Then
    nColor = Branco
Else
    nColor = fr1.BackColor
End If

cmbMEI.Enabled = bValue: cmbMEI.BackColor = nColor
cmbSimples.Enabled = bValue: cmbSimples.BackColor = nColor
cmbIsento.Enabled = bValue: cmbIsento.BackColor = nColor
cmbAlvara.Enabled = bValue: cmbAlvara.BackColor = nColor
cmbTipo.Enabled = bValue: cmbTipo.BackColor = nColor
cmbVistoria.Enabled = bValue: cmbVistoria.BackColor = nColor
cmbVSanit.Enabled = bValue: cmbVSanit.BackColor = nColor
cmbISS.Enabled = bValue: cmbISS.BackColor = nColor
cmbISSE.Enabled = bValue: cmbISSE.BackColor = nColor
cmbDispensaIE.Enabled = bValue: cmbDispensaIE.BackColor = nColor
'cmdOpen.Enabled = bValue
'txtDelimiter.Enabled = bValue: txtDelimiter.BackColor = nColor
'cmdImportar.Enabled = bValue
'cmdPreview.Enabled = bValue
'cmdLimpar.Enabled = bValue
chkDAbe.Enabled = bValue
mskDataAbeIni.Enabled = bValue
mskDataAbeFim.Enabled = bValue
cmbDEnc.Enabled = bValue: cmbDEnc.BackColor = nColor
mskDataEncIni.Enabled = bValue
mskDataEncFim.Enabled = bValue
cmbDSus.Enabled = bValue: cmbDSus.BackColor = nColor
mskDataSusIni.Enabled = bValue
mskDataSusFim.Enabled = bValue
chkContador.Enabled = bValue
cmdDDList2.Enabled = bValue
cmdDD2All.Enabled = bValue
cmdDD2None.Enabled = bValue
cmdDDList3.Enabled = bValue
cmdDDList3All.Enabled = bValue
cmdDDList3None.Enabled = bValue
chkAtividade.Enabled = bValue
chkAtivIss.Enabled = bValue
cmdDDList4.Enabled = bValue
cmdDDList4All.Enabled = bValue
cmdDDList4None.Enabled = bValue
cmbPagto.Enabled = bValue: cmbPagto.BackColor = nColor
cmbImposto.Enabled = bValue: cmbImposto.BackColor = nColor
txtAno.Enabled = bValue: txtAno.BackColor = nColor
chkTipoAtiv(1).Enabled = bValue
chkTipoAtiv(2).Enabled = bValue
chkTipoAtiv(3).Enabled = bValue
chkTipoAtiv(4).Enabled = bValue

End Sub

Private Sub GravaExtrato2(nCodReduz As Long)

Dim aDebito() As Debito, z1 As Variant
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim Achou As Boolean, x As Integer, z As Integer
Dim nSomaDebito As Double, nEval As Integer, nValorCorrecao As Double
Dim nSomaVencer As Double, nValorAtualizado As Double
Dim bMulta As Boolean, bJuros As Boolean
Dim qd As New rdoQuery
Dim sComputer As String
Dim nSeq As Integer
Dim sNumInsc As String
Dim sNomeProp As String
Dim sEnd As String
Dim nAno As Integer
Dim nLancamento As Integer
Dim nSequencia As Integer
Dim nParcela As Integer
Dim nComplemento As Integer

Dim nNumero As Integer
Dim sBairro As String
Dim nCodBanco As Integer

Dim nValorMulta As Double
Dim nValorJuros As Double
Dim nSaldo As Double
Dim sDA As String, sAj As String

ReDim aDebito(0)
nSomaDebito = 0
nSomaVencer = 0
bSel = True

sComputer = NomeDoUsuario
'nCodReduz = 15
For x = 1 To grdMain.Rows
    If Val(grdMain.CellText(x, 1)) = nCodReduz Then
        sNomeProp = grdMain.CellText(x, 2)
        sEnd = grdMain.CellText(x, 5)
        sBairro = grdMain.CellText(x, 7)
        sNumInsc = grdMain.CellText(x, 3)
        Exit For
    End If
Next



dDataAtualiza = Now


Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = nCodReduz
qd(1) = nCodReduz
qd(2) = 1950: qd(3) = 2050
qd(4) = 0: qd(5) = 99
qd(6) = 0: qd(7) = 9999
qd(8) = 0: qd(9) = 99
qd(10) = 0: qd(11) = 99
qd(12) = 0: qd(13) = 99
qd(14) = Format(dDataAtualiza, "mm/dd/yyyy")
qd(15) = NomeDoUsuario
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
        nEval = UBound(aDebito)
        Do Until .EOF
            bJuros = False: bMulta = False
            
            If !CodLancamento = 20 And !statuslanc = 5 Then GoTo proximo
            If !NumParcela = 0 And !statuslanc = 5 Then GoTo proximo
            If !statuslanc = 12 Or !statuslanc = 5 Then GoTo proximo
            If !NumParcela > 0 And !statuslanc = 1 Then GoTo proximo
            'Carrega Matriz Debito
            nEval = UBound(aDebito)
            If !NumParcela = 0 And !statuslanc = 3 And DateDiff("d", !DataVencimento, Now) > 0 Then GoTo proximo
            If !AnoExercicio = 2003 And !CodLancamento = 1 And (!statuslanc <> 2 And !statuslanc <> 1) Then GoTo proximo
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aDebito(UBound(aDebito) + 1)
                nEval = UBound(aDebito)
                aDebito(nEval).nAno = !AnoExercicio
                aDebito(nEval).nLanc = !CodLancamento
                If !CodLancamento = 20 Or !CodLancamento = 8 Then
                   If Not IsNull(!numprocesso) Then
                      If Val(Right$(!numprocesso, 4)) >= 2006 Then
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1) & "-" & RetornaDVProcesso(Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1)) & "/" & Right$(!numprocesso, 4) & ")"
                      Else
                        aDebito(nEval).sLanc = !DESCLANCAMENTO & " (" & !numprocesso & ")"
                      End If
                   Else
                      aDebito(nEval).sLanc = !DESCLANCAMENTO
                   End If
                Else
                   aDebito(nEval).sLanc = !DESCLANCAMENTO
                End If
                aDebito(nEval).nSeq = !SeqLancamento
                aDebito(nEval).nParc = !NumParcela
                aDebito(nEval).nCompl = !CODCOMPLEMENTO
                aDebito(nEval).nSituacao = !statuslanc
                aDebito(nEval).sSituacao = !Situacao
                aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aDebito(nEval).nCodTributo = !CodTributo
                aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                
                If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
                    If Not IsNull(!valorpagoreal) Then
                        aDebito(nEval).nValorAtual = FormatNumber(!valorpagoreal, 2)
                    Else
                        aDebito(nEval).nValorAtual = FormatNumber(0, 2)
                    End If
                Else
                    aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                End If
                
                aDebito(nEval).nValorJuros = IIf(IsNull(!ValorJuros), 0, FormatNumber(!ValorJuros, 2))
                aDebito(nEval).nValorMulta = IIf(IsNull(!ValorMulta), 0, FormatNumber(!ValorMulta, 2))
                aDebito(nEval).nValorCorrecao = IIf(IsNull(!ValorCorrecao), 0, FormatNumber(!ValorCorrecao, 2))
            Else
                If aDebito(nEval).nCodTributo = !CodTributo Then GoTo proximo
            
                aDebito(nEval).nValorTributo = FormatNumber(aDebito(nEval).nValorTributo + !ValorTributo, 2)
               If (!statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7) And Not bForum Then
                    'aDebito(nEval).nValorAtual = FormatNumber(aDebito(nEval).nValorAtual + !ValorTributo, 2)
                    aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                    If Not IsNull(!ValorMulta) Then
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                    Else
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + 0, 2)
                    End If
                    aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
               ElseIf (!statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7) And bForum Then
                    '***************************************************************
                    Sql = "SELECT DEBITOTRIBUTO.CODTRIBUTO,VALORTRIBUTO,ABREVTRIBUTO,DEBITOTRIBUTO.VALORJUROS,DEBITOTRIBUTO.VALORCORRECAO,DEBITOTRIBUTO.VALORMULTA "
                    Sql = Sql & "FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO "
                    Sql = Sql & "AND DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO "
                    Sql = Sql & "AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND DEBITOPARCELA.NumParcela = DEBITOTRIBUTO.NumParcela "
                    Sql = Sql & "AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO Inner Join TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO "
                    Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO = " & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio & " AND DEBITOPARCELA.CODLANCAMENTO = " & !CodLancamento & " AND "
                    Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO = " & !SeqLancamento & " AND DEBITOPARCELA.NUMPARCELA = " & !NumParcela & " AND DEBITOPARCELA.CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND DEBITOTRIBUTO.CODTRIBUTO=" & !CodTributo
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
                        aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                        aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                        aDebito(nEval).nValorAtual = FormatNumber(aDebito(nEval).nValorAtual + !ValorTributo + !ValorJuros + !ValorMulta + !ValorCorrecao, 2)
                       .Close
                    End With
                    '***************************************************************
               Else
                    aDebito(nEval).nValorJuros = FormatNumber(aDebito(nEval).nValorJuros + !ValorJuros, 2)
                    aDebito(nEval).nValorMulta = FormatNumber(aDebito(nEval).nValorMulta + !ValorMulta, 2)
                    aDebito(nEval).nValorCorrecao = FormatNumber(aDebito(nEval).nValorCorrecao + !ValorCorrecao, 2)
                    aDebito(nEval).nValorAtual = FormatNumber(aDebito(nEval).nValorAtual + !ValorTotal, 2)
               End If
            End If
            If !statuslanc = 1 Or !statuslanc = 2 Or !statuslanc = 7 Then
                Sql = "SELECT * FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND RESTITUIDO IS NULL"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                         aDebito(nEval).dDataPag = Format(!DataPagamento, "dd/mm/yyyy")
                         aDebito(nEval).nCodBanco = Val(SubNull(!CodBanco))
                    End If
                End With
            End If
proximo:
            .MoveNext
        Loop
      End If
   .Close
End With

nSeq = 0
nSaldo = 0
'Set qd.ActiveConnection = cn

Sql = "select max(seq) as maximo from extratotmp where computer='" & NomeDoUsuario & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Not IsNull(RdoAux!maximo) Then
    
    nSeq = RdoAux!maximo + 1
Else
    nSeq = 1
End If
RdoAux.Close

For x = 1 To UBound(aDebito)
    With aDebito(x)
    
        If .dDataPag = "00:00:00" Then .dDataPag = "01/01/1900"
        nSeq = nSeq + 1
        Sql = "INSERT Tributacao..EXTRATOTMP"
        Sql = Sql & "(COMPUTER,SEQ,CODREDUZIDO,NOMEPROP,ENDERECO,NUMERO,BAIRRO,CODLANCAMENTO,DESCLANCAMENTO,"
        Sql = Sql & "ANOEXERCICIO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,DATAVENCIMENTO,CODBANCO,DATAPAGAMENTO,STATUSLANC,"
        Sql = Sql & "VALORLANCADO,VALORCORRECAO,VALORMULTA,VALORJUROS,VALORTOTAL,SALDO,DA,AJ) VALUES('" & NomeDoUsuario & "',"
        Sql = Sql & nSeq & "," & nCodReduz & ",'" & Mask(Left$(sNomeProp, 30)) & "','" & Left$(sEnd, 50) & "'," & nNumero & ",'" & Left(sBairro, 25) & "',"
        Sql = Sql & .nLanc & ",'" & .sLanc & "'," & .nAno & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'"
        Sql = Sql & Format(.sVencto, "mm/dd/yyyy") & "'," & .nCodBanco & ",'" & Format(.dDataPag, "mm/dd/yyyy") & "','"
        Sql = Sql & Left$(Format(.nSituacao, "00") & " - " & .sSituacao, 30) & "'," & Virg2Ponto(CStr(.nValorTributo)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorCorrecao)) & "," & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorJuros)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorAtual)) & "," & Virg2Ponto(CStr(nSaldo)) & ",'" & .sDA & "','" & .sAj & "')"
        cn.Execute Sql, rdExecDirect
    End With
PROXIMO2:
Next

Liberado

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
   
End Sub

Private Sub txtAutoInc_KeyPress(KeyAscii As Integer)
Tweak txtAutoInc, KeyAscii, IntegerPositive
End Sub

Private Function IsMEI(nCodigo As Long) As Boolean
Dim nRet As Boolean, Sql As String, RdoAux As rdoResultset
nRet = False

Sql = "select * from periodomei where codigo=" & nCodigo & " order by datainicio desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If IsNull(!datafim) Then
            nRet = True
        End If
    End If
   .Close
End With

IsMEI = nRet

End Function


