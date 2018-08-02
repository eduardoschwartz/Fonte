VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRegistroAtendimento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secretaria de Obras - Registro de Atendimento"
   ClientHeight    =   7125
   ClientLeft      =   8040
   ClientTop       =   4155
   ClientWidth     =   11295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   11295
   Begin Tributacao.jcFrames pnlMaterial 
      Height          =   5745
      Left            =   3060
      Top             =   510
      Visible         =   0   'False
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   10134
      FillColor       =   16777152
      TextBoxColor    =   12582912
      Style           =   3
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Selecione os materiais utilizados"
      TextColor       =   16777215
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
      Begin VB.ListBox lstMaterial 
         Appearance      =   0  'Flat
         Height          =   4755
         Left            =   90
         Style           =   1  'Checkbox
         TabIndex        =   80
         Top             =   390
         Width           =   4905
      End
      Begin prjChameleon.chameleonButton btNovoMaterial 
         Height          =   315
         Left            =   300
         TabIndex        =   81
         ToolTipText     =   "Cadastrar um novo material"
         Top             =   5280
         Visible         =   0   'False
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
         MICON           =   "frmRegistroAtendimento.frx":0000
         PICN            =   "frmRegistroAtendimento.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btAlterarMaterial 
         Height          =   315
         Left            =   1410
         TabIndex        =   82
         ToolTipText     =   "Alterar um material cadastrado"
         Top             =   5280
         Visible         =   0   'False
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
         MICON           =   "frmRegistroAtendimento.frx":0176
         PICN            =   "frmRegistroAtendimento.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btExcluirMaterial 
         Height          =   315
         Left            =   2520
         TabIndex        =   83
         ToolTipText     =   "Excluir um material"
         Top             =   5280
         Visible         =   0   'False
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
         MICON           =   "frmRegistroAtendimento.frx":02EC
         PICN            =   "frmRegistroAtendimento.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btSairMaterial 
         Height          =   315
         Left            =   3780
         TabIndex        =   84
         ToolTipText     =   "Fechar a tela de materiais"
         Top             =   5280
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Fechar"
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
         MICON           =   "frmRegistroAtendimento.frx":03AA
         PICN            =   "frmRegistroAtendimento.frx":03C6
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
   Begin VB.ListBox lstNum 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   10260
      Sorted          =   -1  'True
      TabIndex        =   77
      Top             =   1590
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.ListBox lstNomeLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   2175
      ItemData        =   "frmRegistroAtendimento.frx":0434
      Left            =   1320
      List            =   "frmRegistroAtendimento.frx":0436
      TabIndex        =   76
      Top             =   3720
      Visible         =   0   'False
      Width           =   5265
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   1095
      Left            =   120
      Top             =   3390
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   1931
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço da realização do serviço"
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
      Begin VB.TextBox txtNum 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   7245
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   330
         Width           =   2355
      End
      Begin VB.TextBox txtComplemento 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   70
         Top             =   660
         Width           =   4725
      End
      Begin VB.ComboBox cmbBairro 
         Height          =   315
         Left            =   7230
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   660
         Width           =   3135
      End
      Begin VB.TextBox txtNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1215
         MaxLength       =   50
         TabIndex        =   27
         Top             =   330
         Width           =   4725
      End
      Begin prjChameleon.chameleonButton cmdAddNum 
         Height          =   285
         Left            =   9630
         TabIndex        =   74
         ToolTipText     =   "Adicionar um número ao campo numeração"
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
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
         FCOL            =   32768
         FCOLO           =   32768
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmRegistroAtendimento.frx":0438
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelNum 
         Height          =   285
         Left            =   10020
         TabIndex        =   75
         ToolTipText     =   "Limpar campo numeração"
         Top             =   330
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmRegistroAtendimento.frx":0454
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btCopy 
         Height          =   435
         Left            =   10470
         TabIndex        =   78
         ToolTipText     =   "Copiaer endereço para campo assunto"
         Top             =   420
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   767
         BTYPE           =   3
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   18
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
         FCOL            =   32768
         FCOLO           =   32768
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmRegistroAtendimento.frx":0470
         PICN            =   "frmRegistroAtendimento.frx":048C
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
         Caption         =   "Número(s)..:"
         Height          =   195
         Index           =   25
         Left            =   6270
         TabIndex        =   73
         Top             =   390
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Complemento.:"
         Height          =   195
         Index           =   24
         Left            =   60
         TabIndex        =   71
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Bairro....:"
         Height          =   195
         Index           =   23
         Left            =   6480
         TabIndex        =   69
         Top             =   690
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço.......:"
         Height          =   195
         Index           =   22
         Left            =   60
         TabIndex        =   68
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbAssunto 
      Height          =   315
      Left            =   1530
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1980
      Width           =   8250
   End
   Begin VB.ComboBox cmbReq 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1530
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1260
      Visible         =   0   'False
      Width           =   8475
   End
   Begin VB.ComboBox cmbEquipe 
      Height          =   315
      Left            =   6615
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1620
      Width           =   3165
   End
   Begin VB.ComboBox cmbChefe 
      Height          =   315
      Left            =   1530
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1620
      Width           =   3165
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   9045
      TabIndex        =   3
      ToolTipText     =   "Consulta Cidadãos Cadastrados"
      Top             =   6780
      Width           =   1065
      _ExtentX        =   1879
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":0840
      PICN            =   "frmRegistroAtendimento.frx":085C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton optD 
      Caption         =   "Aguardando"
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
      Index           =   2
      Left            =   1530
      TabIndex        =   30
      Top             =   5295
      Width           =   1410
   End
   Begin VB.TextBox txtSolucao 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   90
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   6015
      Width           =   11085
   End
   Begin VB.OptionButton optD 
      Caption         =   "Indeferido"
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
      Index           =   1
      Left            =   4455
      TabIndex        =   32
      Top             =   5295
      Width           =   1410
   End
   Begin VB.OptionButton optD 
      Caption         =   "Deferido"
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
      Index           =   0
      Left            =   2985
      TabIndex        =   31
      Top             =   5295
      Width           =   1410
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   1530
      MaxLength       =   5000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   4530
      Width           =   9645
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   53
      Text            =   "frmRegistroAtendimento.frx":09B6
      Top             =   4575
      Width           =   1320
   End
   Begin VB.TextBox txtFone 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3060
      Width           =   2805
   End
   Begin VB.TextBox txtBairro 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3060
      Width           =   4290
   End
   Begin VB.TextBox txtCompl 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   8685
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   2700
      Width           =   2490
   End
   Begin VB.TextBox txtEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2700
      Width           =   5775
   End
   Begin VB.TextBox txtCidadao 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1260
      Width           =   8430
   End
   Begin VB.CheckBox chkUrg 
      Caption         =   "Urgente"
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
      Left            =   10080
      TabIndex        =   26
      Top             =   3105
      Width           =   1095
   End
   Begin VB.TextBox txtAssunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   2340
      Width           =   9645
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3870
      TabIndex        =   11
      Top             =   900
      Width           =   1500
   End
   Begin VB.TextBox txtObsTipo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   8370
      MaxLength       =   50
      TabIndex        =   9
      Top             =   495
      Width           =   2805
   End
   Begin VB.ComboBox cmbTipoAtend 
      Height          =   315
      Left            =   5040
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   495
      Width           =   2310
   End
   Begin prjChameleon.chameleonButton cmdRefresh1 
      Height          =   240
      Left            =   3510
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   495
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   3
      TX              =   "!"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":09D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbAtendente 
      Height          =   315
      ItemData        =   "frmRegistroAtendimento.frx":09F4
      Left            =   1170
      List            =   "frmRegistroAtendimento.frx":09F6
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   495
      Width           =   2310
   End
   Begin prjChameleon.chameleonButton cmdRefresh2 
      Height          =   240
      Left            =   7380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   495
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   3
      TX              =   "!"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":09F8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskData 
      Height          =   285
      Left            =   1170
      TabIndex        =   10
      Top             =   900
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmRegistroAtendimento.frx":0A14
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
   Begin esMaskEdit.esMaskedEdit mskDataExec 
      Height          =   285
      Left            =   5490
      TabIndex        =   33
      Top             =   5700
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      MouseIcon       =   "frmRegistroAtendimento.frx":0A30
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
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   9075
      TabIndex        =   37
      ToolTipText     =   "Gravar os Dados"
      Top             =   6780
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
      MICON           =   "frmRegistroAtendimento.frx":0A4C
      PICN            =   "frmRegistroAtendimento.frx":0A68
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
      Left            =   10155
      TabIndex        =   57
      ToolTipText     =   "Sair da Tela"
      Top             =   6780
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
      MICON           =   "frmRegistroAtendimento.frx":0E0D
      PICN            =   "frmRegistroAtendimento.frx":0E29
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
      Left            =   10170
      TabIndex        =   38
      ToolTipText     =   "Cancelar Edição"
      Top             =   6780
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
      MICON           =   "frmRegistroAtendimento.frx":0E97
      PICN            =   "frmRegistroAtendimento.frx":0EB3
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
      Left            =   6855
      TabIndex        =   1
      ToolTipText     =   "Editar Registro"
      Top             =   6780
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
      MICON           =   "frmRegistroAtendimento.frx":100D
      PICN            =   "frmRegistroAtendimento.frx":1029
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
      Left            =   5760
      TabIndex        =   0
      ToolTipText     =   "Novo Registro"
      Top             =   6780
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
      MICON           =   "frmRegistroAtendimento.frx":1183
      PICN            =   "frmRegistroAtendimento.frx":119F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataEnd 
      Height          =   285
      Left            =   7740
      TabIndex        =   34
      Top             =   5700
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      MouseIcon       =   "frmRegistroAtendimento.frx":12F9
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
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   7950
      TabIndex        =   2
      ToolTipText     =   "Excluir Registro"
      Top             =   6780
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&xcluir"
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
      MICON           =   "frmRegistroAtendimento.frx":1315
      PICN            =   "frmRegistroAtendimento.frx":1331
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCns 
      Height          =   315
      Left            =   10080
      TabIndex        =   14
      ToolTipText     =   "Consulta Munícipe/Secretaria"
      Top             =   1245
      Width           =   1065
      _ExtentX        =   1879
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":13D3
      PICN            =   "frmRegistroAtendimento.frx":13EF
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
      Left            =   90
      TabIndex        =   61
      ToolTipText     =   "Imprimir registro"
      Top             =   6780
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
      MICON           =   "frmRegistroAtendimento.frx":1549
      PICN            =   "frmRegistroAtendimento.frx":1565
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRefresh3 
      Height          =   240
      Left            =   4725
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1620
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   3
      TX              =   "!"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":16BF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRefresh4 
      Height          =   240
      Left            =   9810
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1620
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   3
      TX              =   "!"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":16DB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRefresh5 
      Height          =   240
      Left            =   9810
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2025
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   3
      TX              =   "!"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRegistroAtendimento.frx":16F7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataCancel 
      Height          =   285
      Left            =   10080
      TabIndex        =   35
      Top             =   5700
      Width           =   990
      _ExtentX        =   1746
      _ExtentY        =   503
      MouseIcon       =   "frmRegistroAtendimento.frx":1713
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
   Begin prjChameleon.chameleonButton chkMaterial 
      Height          =   345
      Left            =   2160
      TabIndex        =   79
      ToolTipText     =   "Material utilizado"
      Top             =   6750
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Material Utilizado"
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
      MICON           =   "frmRegistroAtendimento.frx":172F
      PICN            =   "frmRegistroAtendimento.frx":174B
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
      Caption         =   "Cancelada em..:"
      Height          =   195
      Index           =   21
      Left            =   8820
      TabIndex        =   67
      Top             =   5745
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Situação..:"
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
      Index           =   20
      Left            =   135
      TabIndex        =   66
      Top             =   135
      Width           =   1455
   End
   Begin VB.Label lblSit 
      Caption         =   "00001/2010"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1620
      TabIndex        =   65
      Top             =   135
      Width           =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "Assunto..............:"
      Height          =   195
      Index           =   19
      Left            =   135
      TabIndex        =   64
      Top             =   2070
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Nome da Equipe.:"
      Height          =   195
      Index           =   18
      Left            =   5220
      TabIndex        =   63
      Top             =   1710
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Chefe da Equipe.:"
      Height          =   195
      Index           =   17
      Left            =   135
      TabIndex        =   62
      Top             =   1710
      Width           =   1320
   End
   Begin VB.Label lblCodBairro 
      Caption         =   "0"
      Height          =   195
      Left            =   6795
      TabIndex        =   60
      Top             =   5295
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lblCodLogr 
      Caption         =   "0"
      Height          =   195
      Left            =   6165
      TabIndex        =   59
      Top             =   5295
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Concluída em..:"
      Height          =   195
      Index           =   16
      Left            =   6570
      TabIndex        =   58
      Top             =   5745
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Executado em..:"
      Height          =   195
      Index           =   15
      Left            =   4275
      TabIndex        =   56
      Top             =   5745
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Do Chefe de Setor/Serviço - Solução ou Justificativa"
      Height          =   195
      Index           =   14
      Left            =   90
      TabIndex        =   55
      Top             =   5745
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Do Secretário......:"
      Height          =   195
      Index           =   13
      Left            =   90
      TabIndex        =   54
      Top             =   5295
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Telefone....:"
      Height          =   195
      Index           =   12
      Left            =   5985
      TabIndex        =   52
      Top             =   3105
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Complemento..:"
      Height          =   195
      Index           =   11
      Left            =   7425
      TabIndex        =   51
      Top             =   2745
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Bairro.................:"
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   50
      Top             =   3105
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço...........:"
      Height          =   195
      Index           =   9
      Left            =   135
      TabIndex        =   49
      Top             =   2730
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Munícipe/Setor.:"
      Height          =   195
      Index           =   8
      Left            =   135
      TabIndex        =   48
      Top             =   1305
      Width           =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "Objeto Processo.:"
      Height          =   195
      Index           =   7
      Left            =   135
      TabIndex        =   47
      Top             =   2385
      Width           =   1365
   End
   Begin VB.Label lblDataProc 
      Caption         =   "01/01/1900"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   7020
      TabIndex        =   46
      Top             =   945
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Data do Processo..:"
      Height          =   195
      Index           =   6
      Left            =   5535
      TabIndex        =   45
      Top             =   945
      Width           =   1500
   End
   Begin VB.Label lblNumReg 
      Alignment       =   1  'Right Justify
      Caption         =   "00001/2010"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   9675
      TabIndex        =   44
      Top             =   135
      Width           =   1500
   End
   Begin VB.Label Label1 
      Caption         =   "N° do Registro..:"
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
      Index           =   5
      Left            =   7605
      TabIndex        =   43
      Top             =   135
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "N° do Processo..:"
      Height          =   195
      Index           =   4
      Left            =   2430
      TabIndex        =   42
      Top             =   945
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Data..........:"
      Height          =   195
      Index           =   3
      Left            =   135
      TabIndex        =   41
      Top             =   945
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Obs..:"
      Height          =   195
      Index           =   2
      Left            =   7830
      TabIndex        =   40
      Top             =   540
      Width           =   465
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Atend..:"
      Height          =   195
      Index           =   1
      Left            =   4005
      TabIndex        =   39
      Top             =   540
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Atendente..:"
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   4
      Top             =   585
      Width           =   960
   End
   Begin VB.Menu mnuName 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCidadao 
         Caption         =   "Munícipe"
      End
      Begin VB.Menu mnuCentroCusto 
         Caption         =   "Centro de Custo"
      End
   End
End
Attribute VB_Name = "frmRegistroAtendimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub btCopy_Click()
Dim sEnd As String

If Val(txtNomeLog.Tag) = 0 Then
    MsgBox "Selecione o endereço.", vbCritical, "Atenção"
    Exit Sub
End If

If MsgBox("Copiar o endereço para o campo assunto?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    sEnd = txtNomeLog.Text & ", "
    sEnd = sEnd & txtNum.Text
    If Trim(txtComplemento.Text) <> "" Then
        sEnd = sEnd & " " & txtComplemento.Text
     End If
    If cmbBairro.ListIndex > -1 Then
        sEnd = sEnd & " - " & cmbBairro.Text
    End If
    txtDesc.Text = txtDesc.Text & " Endereço: " & sEnd
End If
End Sub

Private Sub btSairMaterial_Click()
EventosMaterial False
End Sub

Private Sub chkMaterial_Click()
EventosMaterial True
If cmdGravar.Visible = True Then
    lstMaterial.Enabled = True
Else
    lstMaterial.Enabled = False
End If

End Sub

Private Sub CarregaMaterial()
Dim Sql As String, RdoAux As rdoResultset

lstMaterial.Clear
Sql = "SELECT CODIGO,descricao FROM material_obras order by descricao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstMaterial.AddItem !descricao
        lstMaterial.ItemData(lstMaterial.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdAddNum_Click()
Dim z As Variant, x As Integer
z = InputBox("Digite um número do endereço para execução do serviço.", "Endereço de execução do serviço")
If Not IsNumeric(z) Then
    MsgBox "Número inválido", vbCritical, "Erro"
Else
    For x = 0 To lstNum.ListCount
        If Val(z) = Val(lstNum.List(x)) Then
            MsgBox "Número já cadastrado.", vbCritical, "Atenção"
            Exit Sub
        End If
    Next

    lstNum.AddItem z
    If txtNum.Text = "S/N" Then
        txtNum.Text = z
    Else
        txtNum.Text = txtNum.Text & ", " & z
    End If
End If

End Sub

Private Sub cmdAlterar_Click()
If lblNumReg.Caption = "" Then
    MsgBox "Selecione um registro", vbExclamation, "Atenção"
    Exit Sub
End If

Eventos "INCLUIR"
Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
Evento = ""
End Sub

Private Sub cmdCns_Click()

If Trim$(txtNumProc.Text) <> "" And Trim$(txtNumProc.Text) <> "/" Then
    MsgBox "Escolha munícipe ou um Centro de Custos apenas quando não houver processo.", vbCritical, "Atenção"
Else
    PopupMenu mnuName
End If

End Sub

Private Sub cmdConsultar_Click()
frmCnsRegAtend.show: frmCnsRegAtend.ZOrder 0
End Sub

Private Sub cmdDelNum_Click()
lstNum.Clear
txtNum.Text = "S/N"
End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String

If lblNumReg.Caption = "" Then
    MsgBox "Selecione um registro", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Excluir este registro?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "DELETE FROM REGISTROATENDIMENTO WHERE NUMREG=" & Val(Left(lblNumReg.Caption, 5)) & " AND ANOREG=" & Val(Right(lblNumReg.Caption, 4))
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM REGISTROATENDIMENTO_ENDERECO WHERE NUMREG=" & Val(Left(lblNumReg.Caption, 5)) & " AND ANOREG=" & Val(Right(lblNumReg.Caption, 4))
    cn.Execute Sql, rdExecDirect
    Limpa
End If
End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset, nMaxCod As Integer, nNumproc As Long, nAnoproc As Integer, x As Integer
Dim nCodAssunto As Integer, nChefe As Integer, nequipe As Integer, nCCusto As Integer

If cmbAtendente.ListIndex = -1 Then
    MsgBox "Selecione um atendente.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbTipoAtend.ListIndex = -1 Then
    MsgBox "Selecione um tipo de atendimento.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskData.Text) Then
    MsgBox "Data de atendimento inválido.", vbCritical, "Atenção"
    Exit Sub
End If

If txtCidadao.Text = "" And cmbReq.ListIndex = -1 Then
    MsgBox "Selecione um munícipe ou Secretaria.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtNomeLog.Tag) = 0 Then
    MsgBox "Selecione o endereço para a execução do serviço.", vbCritical, "Atenção"
    Exit Sub
End If


'If cmbBairro.ListIndex = -1 Then
'    MsgBox "Selecione o bairro do atendimento.", vbCritical, "Atenção"
'    Exit Sub
'End If

'If Trim(txtEndereco.Text = "") Then
'    MsgBox "Digite o endereço do atendimento.", vbCritical, "Atenção"
'    Exit Sub
'End If


If cmbAssunto.ListIndex = -1 Then
    nCodAssunto = 0
Else
    nCodAssunto = cmbAssunto.ItemData(cmbAssunto.ListIndex)
End If
If cmbChefe.ListIndex = -1 Then
    nChefe = 0
Else
    nChefe = cmbChefe.ItemData(cmbChefe.ListIndex)
End If
If cmbEquipe.ListIndex = -1 Then
    nequipe = 0
Else
    nequipe = cmbEquipe.ItemData(cmbEquipe.ListIndex)
End If
If cmbReq.ListIndex = -1 Then
    nCCusto = 0
Else
    nCCusto = cmbReq.ItemData(cmbReq.ListIndex)
End If


'If IsNumeric(Left(txtCidadao.Text, 1)) And txtCidadao.Visible Then
'    If Val(Left(txtCidadao.Text, 6)) < 500000 Then
'        MsgBox "Código de cidadão antigo, favor atualizar o código cidadão.", vbCritical, "Atenção"
'        Exit Sub
 '   End If
'E'nd If

'On Error Resume Next
'cn.Close
'Conecta UL, UP

If Trim(txtNumProc.Text) <> "" Then
    nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
    nAnoproc = Right$(txtNumProc.Text, 4)
Else
    nNumproc = 0
    nAnoproc = 0
End If
If Evento = "Novo" Then
    Sql = "SELECT MAX(NUMREG) AS MAXIMO FROM REGISTROATENDIMENTO WHERE ANOREG=" & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nMaxCod = 1
        Else
            nMaxCod = !maximo + 1
        End If
       .Close
    End With
    
    Sql = "INSERT REGISTROATENDIMENTO(NUMREG,ANOREG,ATENDENTE,TIPOATENDIMENTO,OBSTIPO,DATA,NUMPROC,ANOPROC,URGENTE,ASSUNTO,AGUARDO,DEFERIDO,INDEFERIDO,DATAEXEC,DATAEND,SOLUCAO,CIDADAO,CCUSTO,CODLOGR,CODBAIRRO,CHEFE,EQUIPE,CODASSUNTO,DATACANCEL,LOGRADOURO_SERVICO,BAIRRO,COMPLEMENTO_SERVICO,NUMERO_SERVICO) VALUES(" & nMaxCod & "," & Year(Now) & ","
    Sql = Sql & cmbAtendente.ItemData(cmbAtendente.ListIndex) & "," & cmbTipoAtend.ItemData(cmbTipoAtend.ListIndex) & "," & IIf(Trim(txtObsTipo.Text) <> "", "'" & Mask(txtObsTipo.Text) & "'", "Null") & ",'"
    Sql = Sql & Format(mskData.Text, "mm/dd/yyyy") & "'," & nNumproc & "," & nAnoproc & "," & chkUrg.value & "," & IIf(Trim(txtDesc.Text) <> "", "'" & Mask(txtDesc.Text) & "'", "Null") & "," & IIf(optD(2).value, 1, 0) & "," & IIf(optD(0).value, 1, 0) & ","
    Sql = Sql & IIf(optD(1).value, 1, 0) & "," & IIf(IsDate(mskDataExec.Text), "'" & Format(mskDataExec.Text, "mm/dd/yyyy") & "'", "Null") & "," & IIf(IsDate(mskDataEnd.Text), "'" & Format(mskDataEnd.Text, "mm/dd/yyyy") & "'", "Null") & "," & IIf(Trim(txtSolucao.Text) <> "", "'" & Mask(txtSolucao.Text) & "'", "Null") & ","
    If IsNumeric(Left(txtCidadao.Text, 1)) Then
        Sql = Sql & Val(Left(txtCidadao.Text, 6)) & "," & "Null,"
    Else
        Sql = Sql & "Null" & "," & nCCusto & ","
    End If
    If cmbBairro.ListIndex > 0 Then
        Sql = Sql & Val(lblCodLogr.Caption) & "," & Val(lblCodBairro.Caption) & "," & nChefe & "," & nequipe & "," & nCodAssunto & "," & IIf(IsDate(mskDataCancel.Text), "'" & Format(mskDataCancel.Text, "mm/dd/yyyy") & "'", "Null") & "," & Val(txtNomeLog.Tag) & "," & cmbBairro.ItemData(cmbBairro.ListIndex) & ",'" & Mask(txtComplemento.Text) & "','" & txtNum.Text & "')"
    Else
        Sql = Sql & Val(lblCodLogr.Caption) & "," & Val(lblCodBairro.Caption) & "," & nChefe & "," & nequipe & "," & nCodAssunto & "," & IIf(IsDate(mskDataCancel.Text), "'" & Format(mskDataCancel.Text, "mm/dd/yyyy") & "'", "Null") & "," & Val(txtNomeLog.Tag) & "," & 999 & ",'" & Mask(txtComplemento.Text) & "','" & txtNum.Text & "')"
    End If
    lblNumReg.Caption = Format(nMaxCod, "00000") & "/" & Year(Now)
Else
    Sql = "UPDATE REGISTROATENDIMENTO SET ATENDENTE=" & cmbAtendente.ItemData(cmbAtendente.ListIndex) & ",TIPOATENDIMENTO=" & cmbTipoAtend.ItemData(cmbTipoAtend.ListIndex) & ",OBSTIPO=" & IIf(Trim(txtObsTipo.Text) <> "", "'" & Mask(txtObsTipo.Text) & "'", "Null") & ","
    Sql = Sql & "DATA='" & Format(mskData.Text, "mm/dd/yyyy") & "',NUMPROC=" & nNumproc & ",ANOPROC=" & nAnoproc & ",URGENTE=" & chkUrg.value & ",ASSUNTO=" & IIf(Trim(txtDesc.Text) <> "", "'" & Mask(txtDesc.Text) & "'", "Null") & ",AGUARDO=" & IIf(optD(2).value, 1, 0) & ",DEFERIDO=" & IIf(optD(0).value, 1, 0) & ","
    Sql = Sql & "INDEFERIDO=" & IIf(optD(1).value, 1, 0) & ",DATAEXEC=" & IIf(IsDate(mskDataExec.Text), "'" & Format(mskDataExec.Text, "mm/dd/yyyy") & "'", "Null") & ",DATAEND=" & IIf(IsDate(mskDataEnd.Text), "'" & Format(mskDataEnd.Text, "mm/dd/yyyy") & "'", "Null") & ",SOLUCAO=" & IIf(Trim(txtSolucao.Text) <> "", "'" & Mask(txtSolucao.Text) & "'", "Null") & ","
    Sql = Sql & "CIDADAO=" & IIf(IsNumeric(Left(txtCidadao.Text, 1)), Val(Left(txtCidadao.Text, 6)), "Null") & ",CCUSTO=" & IIf(IsNumeric(Left(txtCidadao.Text, 1)), "Null", nCCusto) & ",CODLOGR=" & Val(lblCodLogr.Caption) & ",CODBAIRRO=" & Val(lblCodBairro.Caption) & ",CHEFE=" & nChefe & ","
    If cmbBairro.ListIndex = -1 Then
        Sql = Sql & "EQUIPE=" & nequipe & ",CODASSUNTO=" & nCodAssunto & ",DATACANCEL=" & IIf(IsDate(mskDataCancel.Text), "'" & Format(mskDataCancel.Text, "mm/dd/yyyy") & "'", "Null") & ",LOGRADOURO_SERVICO=" & Val(txtNomeLog.Tag) & ",BAIRRO=" & 999 & ",COMPLEMENTO_SERVICO='" & Mask(txtComplemento.Text) & "',NUMERO_SERVICO='" & txtNum.Text & "'"
    Else
        Sql = Sql & "EQUIPE=" & nequipe & ",CODASSUNTO=" & nCodAssunto & ",DATACANCEL=" & IIf(IsDate(mskDataCancel.Text), "'" & Format(mskDataCancel.Text, "mm/dd/yyyy") & "'", "Null") & ",LOGRADOURO_SERVICO=" & Val(txtNomeLog.Tag) & ",BAIRRO=" & cmbBairro.ItemData(cmbBairro.ListIndex) & ",COMPLEMENTO_SERVICO='" & Mask(txtComplemento.Text) & "',NUMERO_SERVICO='" & txtNum.Text & "'"
    End If
    Sql = Sql & " WHERE NUMREG=" & Val(Left(lblNumReg.Caption, 5)) & " AND ANOREG=" & Val(Right(lblNumReg.Caption, 4))
End If
cn.Execute Sql, rdExecDirect

Sql = "delete from registroatendimento_endereco where anoreg=" & Val(Right(lblNumReg.Caption, 4)) & " and numreg=" & Val(Left(lblNumReg.Caption, 5))
cn.Execute Sql, rdExecDirect

For x = 0 To lstNum.ListCount - 1
    Sql = "insert registroatendimento_endereco (anoreg,numreg,numero_servico) values(" & Val(Right(lblNumReg.Caption, 4)) & "," & Val(Left(lblNumReg.Caption, 5)) & "," & lstNum.List(x) & ")"
    cn.Execute Sql, rdExecDirect
Next

Sql = "delete from registroatendimento_material where anoreg=" & Val(Right(lblNumReg.Caption, 4)) & " and numreg=" & Val(Left(lblNumReg.Caption, 5))
cn.Execute Sql, rdExecDirect

For x = 0 To lstMaterial.ListCount - 1
    If lstMaterial.Selected(x) Then
        Sql = "insert registroatendimento_material (anoreg,numreg,codigo_material) values(" & Val(Right(lblNumReg.Caption, 4)) & "," & Val(Left(lblNumReg.Caption, 5)) & "," & lstMaterial.ItemData(x) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next



If IsDate(mskDataEnd.Text) Then
    lblSit.Caption = "CONCLUIDO"
Else
    If IsDate(mskDataCancel.Text) Then
        lblSit.Caption = "CANCELADO"
    Else
        lblSit.Caption = "AGUARDANDO"
    End If
End If

Eventos "INICIAR"
Evento = ""
End Sub

Private Sub cmdNovo_Click()
Limpa
mskData.Text = Format(Now, "dd/mm/yyyy")
Eventos "INCLUIR"
Evento = "Novo"
cmbAtendente.SetFocus
End Sub

Private Sub cmdPrint_Click()
Dim Sql As String
If lblNumReg.Caption = "" Then
    MsgBox "Selecione um registro", vbExclamation, "Atenção"
    Exit Sub
End If

If Len(txtCidadao.Text) > 0 Then txtCidadao.Text = Left(txtCidadao.Text, 50)

'On Error Resume Next
cn.Close
Conecta UL, UP

Sql = "DELETE FROM REGISTROATENDIMENTOTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "INSERT REGISTROATENDIMENTOTMP(USUARIO,ATENDENTE,TIPO,DATA,PROCESSO,REGISTRO,NOME,ENDERECO,BAIRRO,TELEFONE,COMPL,URGENTE,ASSUNTO,"
Sql = Sql & "DEFERIDO,INDEFERIDO,SOLUCAO,DATAEXEC,DATACONC,CHEFE,EQUIPE,ASSUNTO2,SITUACAO) VALUES('"
Sql = Sql & NomeDeLogin & "','" & cmbAtendente.Text & "','" & cmbTipoAtend.Text & "','" & mskData.Text & "','" & txtNumProc.Text & "','"
Sql = Sql & lblNumReg.Caption & "','" & IIf(txtCidadao.Visible, Mask(txtCidadao.Text), cmbReq.Text) & "','" & Mask(txtEnd.Text) & "','" & Mask(txtBairro.Text) & "','"
Sql = Sql & Mask(txtFone.Text) & "','" & Mask(txtCompl.Text) & "','" & IIf(chkUrg.value = 1, "S", "N") & "','" & Mask(txtDesc.Text) & "','"
Sql = Sql & IIf(optD(0).value = True, "X", " ") & "','" & IIf(optD(1).value = True, "X", " ") & "','" & Mask(txtSolucao.Text) & "','"
Sql = Sql & mskDataExec.Text & "','" & mskDataEnd.Text & "','" & Mask(cmbChefe.Text) & "','" & cmbEquipe.Text & "','" & cmbAssunto.Text & "','" & lblSit.Caption & "')"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport2 "REGATENDIMENTO", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM REGISTROATENDIMENTOTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdRefresh1_Click()
CarregaAT
End Sub

Private Sub cmdRefresh2_Click()
CarregaTA
End Sub

Private Sub cmdRefresh3_Click()
CarregaCF
End Sub

Private Sub cmdRefresh5_Click()
CarregaAS
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If NumRegAtend > 0 Then
    Le NumRegAtend, AnoRegAtend
End If
End Sub

Private Sub Form_Load()

Centraliza Me
CarregaBairro
CarregaAT
CarregaAS
CarregaCC
CarregaTA
CarregaCF
CarregaEQ
CarregaMaterial
Limpa
Eventos "INICIAR"
Evento = ""
End Sub

Private Sub Limpa()
chkUrg.value = vbUnchecked
LimpaMascara mskData
lblDataProc.Caption = ""
lblNumReg.Caption = ""
lblSit.Caption = ""
txtObsTipo.Text = ""
txtNumProc.Text = ""
cmbAtendente.ListIndex = -1
cmbAssunto.ListIndex = -1
cmbTipoAtend.ListIndex = -1
cmbChefe.ListIndex = -1
cmbEquipe.ListIndex = -1
txtCidadao.Text = ""
cmbReq.ListIndex = -1
txtEnd.Text = ""
txtCompl.Text = ""
txtBairro.Text = ""
txtFone.Text = ""
LimpaMascara mskDataExec
LimpaMascara mskDataEnd
LimpaMascara mskDataCancel
txtAssunto.Text = ""
txtDesc.Text = ""
txtSolucao.Text = ""
optD(0).value = False
optD(1).value = False
optD(2).value = False
'txtEndereco.Text = ""
cmbBairro.ListIndex = -1
txtNomeLog.Text = ""
txtNum.Text = "S/N"
txtComplemento.Text = ""
lstNum.Clear
pnlMaterial.Visible = False

For x = 0 To lstMaterial.ListCount - 1
    lstMaterial.Selected(x) = False
Next
End Sub

Private Sub CarregaAT()
Dim Sql As String, RdoAux As rdoResultset

cmbAtendente.Clear
Sql = "SELECT CODIGO,NOME FROM PARAMOBRA WHERE SIGLA='AT'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAtendente.AddItem !Nome
        cmbAtendente.ItemData(cmbAtendente.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaBairro()
Dim Sql As String, RdoAux As rdoResultset

cmbBairro.Clear
Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO<>999 ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbBairro.AddItem !DescBairro
        cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub CarregaAS()
Dim Sql As String, RdoAux As rdoResultset

cmbAssunto.Clear
Sql = "SELECT CODIGO,NOME FROM PARAMOBRA WHERE SIGLA='AS'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAssunto.AddItem !Nome
        cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaCC()
Dim Sql As String, RdoAux As rdoResultset
cmbReq.Clear
Sql = "SELECT CODIGO,DESCRICAO FROM CENTROCUSTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbReq.AddItem !descricao
       cmbReq.ItemData(cmbReq.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaCF()

Dim Sql As String, RdoAux As rdoResultset

cmbChefe.Clear
Sql = "SELECT CODIGO,NOME FROM PARAMOBRA WHERE SIGLA='FC'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbChefe.AddItem !Nome
        cmbChefe.ItemData(cmbChefe.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaEQ()

Dim Sql As String, RdoAux As rdoResultset

cmbEquipe.Clear
Sql = "SELECT CODIGO,NOME FROM PARAMOBRA WHERE SIGLA='EQ'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbEquipe.AddItem !Nome
        cmbEquipe.ItemData(cmbEquipe.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaTA()
Dim Sql As String, RdoAux As rdoResultset

cmbTipoAtend.Clear
Sql = "SELECT CODIGO,NOME FROM PARAMOBRA WHERE SIGLA='TA'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipoAtend.AddItem !Nome
        cmbTipoAtend.ItemData(cmbTipoAtend.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub mnuCentroCusto_Click()
Dim z As Variant
'z = InputBox("Digite o nome da Secretaria,Setor ou Centro de Custos.", "Informação requerida")
'If z <> "" Then
'    z = Left(z, 50)
'End If
'txtCidadao.Text = z
cmbReq.Visible = True
txtCidadao.Text = ""
txtCidadao.Visible = False
End Sub

Private Sub mnuCidadao_Click()
Dim frm As Object
cmbReq.Visible = False
txtCidadao.Visible = True
Set frm = frmCidadao
frm.sForm = Me.Name
frmCidadao.show: frmCidadao.ZOrder (0)
End Sub

Private Sub mskData_GotFocus()
mskData.SetFocus
mskData.SelStart = 0: mskData.SelLength = Len(mskData.Text)
End Sub

Private Sub txtNumProc_GotFocus()
txtNumProc.SetFocus
txtNumProc.SelStart = 0: txtNumProc.SelLength = Len(txtNumProc.Text)
End Sub

Private Sub txtNumProc_LostFocus()
Dim sNumProc As String, nNumproc As Long, nAnoproc As Integer
Dim sValidaProc As String, Sql As String, RdoAux As rdoResultset

If Trim$(txtNumProc.Text) <> "" And Trim$(txtNumProc.Text) <> "/" Then
    cmbReq.Visible = False
    txtCidadao.Visible = True
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
        nAnoproc = Right$(txtNumProc.Text, 4)
        sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
        Sql = "SELECT * FROM vwFULLPROCESSO WHERE ANO=" & nAnoproc & " AND NUMERO=" & nNumproc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                lblDataProc.Caption = Format(!DATAENTRADA, "dd/mm/yyyy")
                txtAssunto.Text = !Complemento
                If IsNull(!nomecidadao) Then
                    txtCidadao.Text = !descricao
                    txtEnd.Text = ""
                    txtCompl.Text = ""
                    txtBairro.Text = ""
                    txtFone.Text = ""
                Else
                    txtCidadao.Text = Format(!CodCidadao, "000000") & " - " & !nomecidadao
                    txtEnd.Text = SubNull(!Endereco) & ", " & SubNull(!NUMIMOVEL)
                    txtCompl.Text = SubNull(!COMPL)
                    txtBairro.Text = SubNull(!DescBairro)
                    txtFone.Text = SubNull(!telefone)
                    lblCodLogr.Caption = Val(SubNull(!CodLogradouro))
                    lblCodBairro.Caption = Val(SubNull(!CodBairro))
                End If
            Else
                MsgBox "Processo não cadastrado.", vbCritical, "Atenção"
            End If
           .Close
        End With
    Else
        MsgBox "Processo inválido.", vbExclamation, "Atenção"
        lblDataProc.Caption = ""
        txtNumProc.SetFocus
    End If
Else
    txtCidadao.Text = ""
    txtEnd.Text = ""
    txtBairro.Text = ""
    txtCompl.Text = ""
    txtFone.Text = ""
End If

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdPrint.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdConsultar.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmdCns.Enabled = False
   cmbAtendente.Enabled = False
   cmbAtendente.BackColor = Kde
   cmbBairro.Enabled = False
   cmbBairro.BackColor = Kde
   cmbTipoAtend.Enabled = False
   cmbTipoAtend.BackColor = Kde
   cmbChefe.Enabled = False
   cmbChefe.BackColor = Kde
   cmbAssunto.Enabled = False
   cmbAssunto.BackColor = Kde
   cmbReq.Enabled = False
   cmbReq.BackColor = Kde
   cmbEquipe.Enabled = False
   cmbEquipe.BackColor = Kde
   txtNomeLog.Locked = True
   txtNomeLog.BackColor = Kde
   txtComplemento.Locked = True
   txtComplemento.BackColor = Kde
   cmdAddNum.Enabled = False
   cmdDelNum.Enabled = False
   txtObsTipo.Locked = True
   txtObsTipo.BackColor = Kde
   txtNumProc.Locked = True
   txtNumProc.BackColor = Kde
   mskData.Locked = True
   mskData.BackColor = Kde
   chkUrg.Enabled = False
   txtDesc.Locked = True
   txtDesc.BackColor = Kde
'   txtEndereco.Locked = True
'   txtEndereco.BackColor = Kde
   optD(0).Enabled = False
   optD(1).Enabled = False
   optD(2).Enabled = False
   txtSolucao.Locked = True
   txtSolucao.BackColor = Kde
   mskDataExec.Locked = True
   mskDataExec.BackColor = Kde
   mskDataEnd.Locked = True
   mskDataEnd.BackColor = Kde
   mskDataCancel.Locked = True
   mskDataCancel.BackColor = Kde
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdPrint.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdConsultar.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   cmdCns.Enabled = True
   cmbAtendente.Enabled = True
   cmbAtendente.BackColor = Branco
   cmbBairro.Enabled = True
   cmbBairro.BackColor = Branco
   cmbTipoAtend.Enabled = True
   cmbTipoAtend.BackColor = Branco
   cmbAssunto.Enabled = True
   cmbAssunto.BackColor = Branco
   cmbChefe.Enabled = True
   cmbChefe.BackColor = Branco
   cmbReq.Enabled = True
   cmbReq.BackColor = Branco
   cmbEquipe.Enabled = True
   cmbEquipe.BackColor = Branco
   cmdAddNum.Enabled = True
   cmdDelNum.Enabled = True
   
   txtNomeLog.Locked = False
   txtNomeLog.BackColor = Branco
   txtComplemento.Locked = False
   txtComplemento.BackColor = Branco
   txtObsTipo.Locked = False
   txtObsTipo.BackColor = Branco
'   txtEndereco.Locked = False
'   txtEndereco.BackColor = Branco
   txtNumProc.Locked = False
   txtNumProc.BackColor = Branco
   mskData.Locked = False
   mskData.BackColor = Branco
   chkUrg.Enabled = True
   txtDesc.Locked = False
   txtDesc.BackColor = Branco
   optD(0).Enabled = True
   optD(1).Enabled = True
   optD(2).Enabled = True
   txtSolucao.Locked = False
   txtSolucao.BackColor = Branco
   mskDataExec.Locked = False
   mskDataExec.BackColor = Branco
   mskDataEnd.Locked = False
   mskDataEnd.BackColor = Branco
   mskDataCancel.Locked = False
   mskDataCancel.BackColor = Branco
End If

End Sub

Private Sub Le(nNumero As Long, nAno As Integer)
Dim Sql As String, RdoAux As rdoResultset, x As Integer, RdoAux2 As rdoResultset, bFind As Boolean

Limpa
Sql = "SELECT registroatendimento.*,Logradouro.Endereco as nomelogradouro  FROM REGISTROATENDIMENTO LEFT OUTER JOIN logradouro ON registroatendimento.logradouro_servico = logradouro.codlogradouro WHERE NUMREG=" & nNumero & " AND ANOREG=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblNumReg.Caption = Format(!numreg, "00000") & "/" & !anoreg
        For x = 0 To cmbAtendente.ListCount - 1
            If cmbAtendente.ItemData(x) = !ATENDENTE Then
                cmbAtendente.ListIndex = x
                Exit For
            End If
        Next
        For x = 0 To cmbTipoAtend.ListCount - 1
            If cmbTipoAtend.ItemData(x) = !TIPOATENDIMENTO Then
                cmbTipoAtend.ListIndex = x
                Exit For
            End If
        Next
        For x = 0 To cmbBairro.ListCount - 1
            If cmbBairro.ItemData(x) = Val(SubNull(!Bairro)) Then
                cmbBairro.ListIndex = x
                Exit For
            End If
        Next
        txtComplemento.Text = SubNull(!Complemento_servico)
        txtNomeLog.Text = SubNull(!NomeLogradouro)
        txtNomeLog.Tag = Val(SubNull(!logradouro_servico))
       
'        txtEndereco.Text = SubNull(!Endereco)
        For x = 0 To cmbChefe.ListCount - 1
            If cmbChefe.ItemData(x) = !Chefe Then
                cmbChefe.ListIndex = x
                Exit For
            End If
        Next
        For x = 0 To cmbEquipe.ListCount - 1
            If cmbEquipe.ItemData(x) = !equipe Then
                cmbEquipe.ListIndex = x
                Exit For
            End If
        Next
        For x = 0 To cmbAssunto.ListCount - 1
            If cmbAssunto.ItemData(x) = !CODASSUNTO Then
                cmbAssunto.ListIndex = x
                Exit For
            End If
        Next
        txtObsTipo.Text = SubNull(!OBSTIPO)
        mskData.Text = Format(!Data, "dd/mm/yyyy")
        If Val(SubNull(!NumProc)) > 0 Then
            txtNumProc.Text = !NumProc & RetornaDVProcesso(!NumProc) & "/" & !AnoProc
            txtNumProc_LostFocus
        Else
            If Not IsNull(!cidadao) Then
                Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & !cidadao
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        txtCidadao.Text = Format(!CodCidadao, "000000") & " - " & !nomecidadao
                        txtEnd.Text = !Endereco & ", " & !NUMIMOVEL
                        txtBairro.Text = SubNull(!DescBairro)
                        txtCompl.Text = SubNull(!Complemento)
                        txtFone.Text = SubNull(!telefone)
                        txtCidadao.Visible = True
                        cmbReq.Visible = False
                    End If
                   .Close
                End With
            Else
                For x = 0 To cmbReq.ListCount - 1
                    cmbReq.ListIndex = x
                    If cmbReq.ItemData(cmbReq.ListIndex) = !ccusto Then
                        txtCidadao.Visible = False
                        cmbReq.Visible = True
                        txtCidadao.Text = ""
                        Exit For
                    End If
                Next
            End If
        End If
        chkUrg.value = IIf(!urgente, 1, 0)
        txtDesc.Text = SubNull(!assunto)
        optD(0).value = !deferido
        optD(1).value = !indeferido
        optD(2).value = !aguardo
        If IsDate(!Dataexec) Then
            mskDataExec.Text = Format(!Dataexec, "dd/mm/yyyy")
        End If
        If IsDate(!dataend) Then
            mskDataEnd.Text = Format(!dataend, "dd/mm/yyyy")
        End If
        txtSolucao.Text = SubNull(!solucao)
        If Not IsNull(!ccusto) Then
            txtCidadao.Text = !ccusto
        End If
        lblCodLogr.Caption = Val(SubNull(!CodLogr))
        lblCodBairro.Caption = Val(SubNull(!CodBairro))
    End If
   .Close
End With



Sql = "SELECT  registroatendimento_endereco.numero_servico FROM registroatendimento_endereco WHERE NUMREG=" & nNumero & " AND ANOREG=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    
    Do Until .EOF
        lstNum.AddItem Val(SubNull(!numero_servico))
       .MoveNext
    Loop
    If lstNum.ListCount = 0 Then
        txtNum.Text = "S/N"
    Else
        txtNum.Text = ""
        For x = 0 To lstNum.ListCount - 1
            txtNum.Text = txtNum.Text & lstNum.List(x) & ", "
        Next
        txtNum.Text = Left(txtNum.Text, Len(txtNum.Text) - 2)
    End If
End With

Sql = "SELECT  * FROM registroatendimento_material WHERE NUMREG=" & nNumero & " AND ANOREG=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bFind = False
        For x = 0 To lstMaterial.ListCount - 1
            If lstMaterial.ItemData(x) = !Codigo_material Then
                bFind = True
                Exit For
            End If
        Next
        If bFind Then
            lstMaterial.Selected(x) = True
        End If
       .MoveNext
    Loop
End With


NumRegAtend = 0
AnoRegAtend = 0

If IsDate(mskDataEnd.Text) Then
    lblSit.Caption = "CONCLUIDO"
Else
    If IsDate(mskDataCancel.Text) Then
        lblSit.Caption = "CANCELADO"
    Else
        lblSit.Caption = "AGUARDANDO"
    End If
End If




End Sub

Private Sub txtNomeLog_Change()
If Trim$(txtNomeLog) = "" Then
   txtNomeLog.Tag = 0
End If
End Sub

Private Sub txtNomeLog_GotFocus()
txtNomeLog.SelStart = 0
txtNomeLog.SelLength = Len(txtNomeLog.Text)
End Sub

Private Sub txtNomeLog_KeyPress(KeyAscii As Integer)
Dim Sql As String, RdoAux As rdoResultset

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
             lstNomeLog.Left = 1320
             lstNomeLog.Top = 3720
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
   txtNomeLog.Tag = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
    txtNomeLog.Tag = lstNomeLog.ItemData(lstNomeLog.ListIndex)
    cmdAddNum.SetFocus
End If

lstNomeLog.Visible = False
End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
On Error GoTo Erro
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
        txtNomeLog.Text = lstNomeLog.Text
        txtNomeLog.Tag = lstNomeLog.ItemData(lstNomeLog.ListIndex)
        cmdAddNum.SetFocus
    End If
    lstNomeLog.Visible = False
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
End If

Exit Sub
Erro:
lstNomeLog.Visible = False

End Sub

Private Sub EventosMaterial(bLock As Boolean)

pnlMaterial.Visible = bLock
cmdNovo.Enabled = Not bLock
cmdPrint.Enabled = Not bLock
cmdAlterar.Enabled = Not bLock
cmdExcluir.Enabled = Not bLock
cmdSair.Enabled = Not bLock
cmdConsultar.Enabled = Not bLock
cmdGravar.Enabled = Not bLock
cmdCancel.Enabled = Not bLock
cmdCns.Enabled = Not bLock
cmbAtendente.Enabled = Not bLock
cmbBairro.Enabled = Not bLock
cmbTipoAtend.Enabled = Not bLock
cmbChefe.Enabled = Not bLock
cmbAssunto.Enabled = Not bLock
cmbReq.Enabled = Not bLock
cmbEquipe.Enabled = Not bLock
txtNomeLog.Enabled = Not bLock
txtComplemento.Enabled = Not bLock
cmdAddNum.Enabled = Not bLock
cmdDelNum.Enabled = Not bLock
txtObsTipo.Enabled = Not bLock
txtNumProc.Enabled = Not bLock
mskData.Enabled = Not bLock
chkUrg.Enabled = Not bLock
txtDesc.Enabled = Not bLock
optD(0).Enabled = Not bLock
optD(1).Enabled = Not bLock
optD(2).Enabled = Not bLock
txtSolucao.Enabled = Not bLock
mskDataExec.Enabled = Not bLock
mskDataEnd.Enabled = Not bLock
mskDataCancel.Enabled = Not bLock
chkMaterial.Enabled = Not bLock
btCopy.Enabled = Not bLock
cmdRefresh4.Enabled = Not bLock
cmdRefresh5.Enabled = Not bLock

End Sub

