VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmCadastroObra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Obras"
   ClientHeight    =   6030
   ClientLeft      =   4455
   ClientTop       =   3030
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   11430
   Begin Tributacao.jcFrames frEdit 
      Height          =   4515
      Left            =   630
      Top             =   990
      Visible         =   0   'False
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   7964
      Caption         =   ""
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
      Begin prjChameleon.chameleonButton cmdBackCep 
         Height          =   315
         Left            =   8100
         TabIndex        =   38
         ToolTipText     =   "Retorna para o campo de CEP"
         Top             =   60
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Retorna CEP"
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
         MICON           =   "frmCadastroObra.frx":0000
         PICN            =   "frmCadastroObra.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Tributacao.jcFrames frTela 
         Height          =   3975
         Index           =   0
         Left            =   45
         Top             =   495
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   7011
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "jcFrames1"
         TextBoxHeight   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7380
            MaxLength       =   50
            TabIndex        =   43
            Top             =   1305
            Width           =   2445
         End
         Begin VB.TextBox txtNum 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5985
            MaxLength       =   50
            TabIndex        =   41
            Top             =   1305
            Width           =   600
         End
         Begin VB.TextBox txtEndereco 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1125
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1305
            Width           =   4335
         End
         Begin VB.TextBox txtBairro 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3915
            MaxLength       =   50
            TabIndex        =   31
            Top             =   945
            Width           =   3255
         End
         Begin VB.TextBox txtUF 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2610
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   30
            Top             =   945
            Width           =   375
         End
         Begin VB.TextBox txtCEP 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   720
            MaxLength       =   8
            TabIndex        =   28
            Top             =   945
            Width           =   915
         End
         Begin VB.ComboBox cmbTipoDoc 
            Height          =   315
            ItemData        =   "frmCadastroObra.frx":05C2
            Left            =   675
            List            =   "frmCadastroObra.frx":05CC
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   225
            Width           =   870
         End
         Begin VB.TextBox txtNome 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   5445
            TabIndex        =   14
            Top             =   225
            Width           =   4515
         End
         Begin esMaskEdit.esMaskedEdit mskCPF 
            Height          =   285
            Left            =   2205
            TabIndex        =   16
            Top             =   225
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            MouseIcon       =   "frmCadastroObra.frx":05DB
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
            Left            =   2205
            TabIndex        =   20
            Top             =   225
            Width           =   1920
            _ExtentX        =   3387
            _ExtentY        =   503
            MouseIcon       =   "frmCadastroObra.frx":05F7
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
         Begin prjChameleon.chameleonButton cmdCep 
            Height          =   270
            Index           =   0
            Left            =   1665
            TabIndex        =   29
            ToolTipText     =   "Busca CEP"
            Top             =   970
            Width           =   270
            _ExtentX        =   476
            _ExtentY        =   476
            BTYPE           =   5
            TX              =   "..."
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
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCadastroObra.frx":0613
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
            Caption         =   "Compl..:"
            Height          =   195
            Index           =   5
            Left            =   6705
            TabIndex        =   44
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº..:"
            Height          =   195
            Index           =   4
            Left            =   5580
            TabIndex        =   42
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Endereço..:"
            Height          =   195
            Index           =   3
            Left            =   225
            TabIndex        =   40
            Top             =   1350
            Width           =   870
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Bairro..:"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   27
            Top             =   990
            Width           =   645
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "UF..:"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   26
            Top             =   990
            Width           =   420
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "CEP..:"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   25
            Top             =   990
            Width           =   555
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            Caption         =   "Localização"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   225
            TabIndex        =   24
            Top             =   630
            Width           =   1050
         End
         Begin VB.Shape Shape1 
            Height          =   1320
            Left            =   90
            Top             =   720
            Width           =   9870
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo..:"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nº..:"
            Height          =   240
            Index           =   1
            Left            =   1800
            TabIndex        =   18
            Top             =   270
            Width           =   420
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Nome/Razão..:"
            Height          =   240
            Index           =   2
            Left            =   4275
            TabIndex        =   17
            Top             =   270
            Width           =   1140
         End
      End
      Begin Tributacao.jcFrames frCEP 
         Height          =   3975
         Left            =   45
         Top             =   495
         Visible         =   0   'False
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   7011
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "BUSCA DE CEP"
         TextBoxHeight   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ThemeColor      =   2
         Begin VB.TextBox txtLogCep 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   3285
            MaxLength       =   40
            TabIndex        =   35
            Top             =   405
            Width           =   5190
         End
         Begin VB.TextBox txtCepCrit 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   720
            MaxLength       =   8
            TabIndex        =   34
            Top             =   405
            Width           =   1005
         End
         Begin prjChameleon.chameleonButton cmdConsultar 
            Height          =   315
            Left            =   8595
            TabIndex        =   36
            ToolTipText     =   "Consulta Cidadãos Cadastrados"
            Top             =   405
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   556
            BTYPE           =   3
            TX              =   "Pesquisar"
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
            MICON           =   "frmCadastroObra.frx":062F
            PICN            =   "frmCadastroObra.frx":064B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.ListView lvCep 
            Height          =   3090
            Left            =   135
            TabIndex        =   37
            Top             =   810
            Width           =   9750
            _ExtentX        =   17198
            _ExtentY        =   5450
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "CEP"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Titulo"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Logradouro"
               Object.Width           =   6703
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Bairro"
               Object.Width           =   5292
            EndProperty
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Logradouro...:"
            Height          =   240
            Index           =   1
            Left            =   2160
            TabIndex        =   33
            Top             =   450
            Width           =   1050
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "CEP..:"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   32
            Top             =   450
            Width           =   555
         End
      End
      Begin Tributacao.jcFrames frTela 
         Height          =   3975
         Index           =   3
         Left            =   45
         Top             =   495
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   7011
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "jcFrames1"
         TextBoxHeight   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Text3 
            Height          =   330
            Left            =   1080
            TabIndex        =   23
            Text            =   "Dados adicionais"
            Top             =   585
            Width           =   2940
         End
      End
      Begin Tributacao.jcFrames frTela 
         Height          =   3975
         Index           =   2
         Left            =   45
         Top             =   495
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   7011
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "jcFrames1"
         TextBoxHeight   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Text2 
            Height          =   330
            Left            =   1080
            TabIndex        =   22
            Text            =   "Dados da Obra"
            Top             =   585
            Width           =   2940
         End
      End
      Begin Tributacao.jcFrames frTela 
         Height          =   3975
         Index           =   1
         Left            =   45
         Top             =   495
         Width           =   10050
         _ExtentX        =   17727
         _ExtentY        =   7011
         FillColor       =   14745599
         Style           =   4
         RoundedCornerTxtBox=   -1  'True
         Caption         =   "jcFrames1"
         TextBoxHeight   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.TextBox Text1 
            Height          =   330
            Left            =   1080
            TabIndex        =   21
            Text            =   "Cadastro Obra"
            Top             =   585
            Width           =   2940
         End
      End
      Begin VB.Label lblTit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dados Adicionais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   3
         Left            =   8010
         MouseIcon       =   "frmCadastroObra.frx":07A5
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   90
         Width           =   1770
      End
      Begin VB.Label lblTit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dados da Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   2
         Left            =   5625
         MouseIcon       =   "frmCadastroObra.frx":08F7
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   90
         Width           =   1770
      End
      Begin VB.Label lblTit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro da Obra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   1
         Left            =   3240
         MouseIcon       =   "frmCadastroObra.frx":0A49
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   90
         Width           =   1770
      End
      Begin VB.Label lblTit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cadastro do Proprietário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   180
         MouseIcon       =   "frmCadastroObra.frx":0B9B
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   90
         Width           =   2310
      End
   End
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   45
      TabIndex        =   1
      Top             =   -45
      Width           =   11355
      Begin VB.ComboBox cmbAno 
         Height          =   315
         Left            =   4005
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   135
         Width           =   825
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmCadastroObra.frx":0CED
         Left            =   2700
         List            =   "frmCadastroObra.frx":0D15
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   135
         Width           =   1275
      End
      Begin prjChameleon.chameleonButton cmdNovo 
         Height          =   315
         Left            =   8010
         TabIndex        =   5
         ToolTipText     =   "Novo Registro"
         Top             =   135
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
         MICON           =   "frmCadastroObra.frx":0D7E
         PICN            =   "frmCadastroObra.frx":0D9A
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
         Left            =   10215
         TabIndex        =   7
         ToolTipText     =   "Excluir Registro"
         Top             =   135
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
         MICON           =   "frmCadastroObra.frx":0EF4
         PICN            =   "frmCadastroObra.frx":0F10
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
         Left            =   9120
         TabIndex        =   8
         ToolTipText     =   "Gravar os Dados"
         Top             =   135
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
         MICON           =   "frmCadastroObra.frx":0FB2
         PICN            =   "frmCadastroObra.frx":0FCE
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
         Left            =   10215
         TabIndex        =   9
         ToolTipText     =   "Cancelar Edição"
         Top             =   135
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
         MICON           =   "frmCadastroObra.frx":1373
         PICN            =   "frmCadastroObra.frx":138F
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
         Left            =   9120
         TabIndex        =   6
         ToolTipText     =   "Editar Registro"
         Top             =   135
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
         MICON           =   "frmCadastroObra.frx":14E9
         PICN            =   "frmCadastroObra.frx":1505
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
         Caption         =   "Obras cadastradas no período de "
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   200
         Width           =   2445
      End
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdMain 
      Height          =   5460
      Left            =   45
      TabIndex        =   0
      Top             =   495
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   9631
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BackColor       =   16777215
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
      DisableIcons    =   -1  'True
      GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
   End
End
Attribute VB_Name = "frmCadastroObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec As Boolean
Dim Evento As String
Dim nMunicipio As Long

Private Sub cmbAno_Click()
Le
End Sub

Private Sub cmbMes_Click()
Le
End Sub

Private Sub cmbTipoDoc_Click()
If cmbTipoDoc.ListIndex = 0 Then
    mskCPF.Visible = True
    mskCNPJ.Visible = False
Else
    mskCPF.Visible = False
    mskCNPJ.Visible = True
End If
End Sub

Private Sub cmdAlterar_Click()
Dim nRow As Integer

With grdMain
    nRow = .SelectedRow
    If nRow = 0 Then
        MsgBox "Selecione um registro.", vbCritical, "Erro"
       Exit Sub
    End If
    If .CellText(nRow, 1) <> "" Then
        cmbTipoDoc.Text = .CellText(nRow, 1)
        If cmbTipoDoc.ListIndex = 0 Then
            mskCPF.Visible = True
            mskCNPJ.Visible = False
            mskCPF.Text = .CellText(nRow, 2)
        Else
            mskCPF.Visible = False
            mskCNPJ.Visible = True
            mskCNPJ.Text = .CellText(nRow, 2)
        End If
        txtNome.Text = .CellText(nRow, 3)
    End If
End With
Eventos "INCLUIR"
Evento = "Alterar"

frEdit.Visible = True
End Sub

Private Sub cmdBackCep_Click()
Dim x As Integer

If lvCep.ListItems.Count > 0 Then
    txtCEP(Val(frCEP.Tag)).Text = lvCep.SelectedItem.Text
Else
    txtCEP(Val(frCEP.Tag)).Text = ""
End If
frCEP.Visible = False
frEdit.Caption = ""

For x = 0 To 3: lblTit(x).Visible = True: Next
cmdBackCep.Visible = False
cmdGravar.Enabled = True
cmdCancel.Enabled = True
txtCEP_LostFocus (Val(frCEP.Tag))

End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
Evento = ""
frEdit.Visible = False
End Sub

Private Sub cmdCep_Click(Index As Integer)
Dim x As Integer
frCEP.Tag = Index
frCEP.Visible = True
frCEP.ZOrder 0
frEdit.Caption = "BUSCA CEP"
For x = 0 To 3: lblTit(x).Visible = False: Next
cmdBackCep.Visible = True
cmdGravar.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub cmdConsultar_Click()
Dim Sql As String, RdoAux As rdoResultset, z As Long

z = SendMessage(lvCep.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT * FROM SO_CEP WHERE 1=1 "
If Trim(txtCepCrit.Text) <> "" Then
    Sql = Sql & " AND NU_CEP LIKE '" & txtCepCrit.Text & "%' "
End If
If Trim(txtLogCep.Text) <> "" Then
    Sql = Sql & " AND TE_DESCRICAO_CEP LIKE '" & txtLogCep.Text & "%' "
End If
Sql = Sql & "ORDER BY NU_CEP"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvCep.ListItems.Add(, "C" & !NU_CEP, !NU_CEP)
        itmX.SubItems(1) = !ID_TIPO_LOGR
        itmX.SubItems(2) = !TE_DESCRICAO_CEP
        itmX.SubItems(3) = !NM_BAIRRO
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdGravar_Click()
Eventos "INICIAR"
frEdit.Visible = False
End Sub

Private Sub cmdNovo_Click()
Eventos "INCLUIR"
Evento = "Novo"
Limpa
frEdit.Visible = True
End Sub

Private Sub Form_Load()
Dim x As Integer
For x = 2008 To Year(Now)
    cmbAno.AddItem x
Next
nMunicipio = 21244 'Jaboticabal
cmbTipoDoc.ListIndex = 0
cmbAno.Text = Year(Now)
bExec = False
cmbMes.ListIndex = Month(Now) - 1
bExec = True
Centraliza Me
GridHeader
TelaAtiva 0
Le
Eventos "INICIAR"
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
    .AddColumn "kCPFCNPJ", "Tipo", ecgHdrTextALignLeft, , 38
    .AddColumn "kIdProp", "Id.Proprietário", ecgHdrTextALignLeft, , 110
    .AddColumn "kNomeProp", "Nome do Proprietário", ecgHdrTextALignLeft, , 180
    .AddColumn "kNomeObra", "Nome da Obra", ecgHdrTextALignLeft, , 180
    .AddColumn "kDataAt", "Dt.Atual.", ecgHdrTextALignCentre, , 80
    .AddColumn "kNumAlv", "Nº Alvará", ecgHdrTextALignLeft, , 60
    .AddColumn "kDtAlv", "Dt Alvará", ecgHdrTextALignCentre, , 80
    .AddColumn "kNumHab", "Habite-se", ecgHdrTextALignLeft, , 60
    .AddColumn "kEndObr", "Endereço da Obra", ecgHdrTextALignLeft, , 180
    .AddColumn "kNumObr", "Num", ecgHdrTextALignRight, , 38
End With

End Sub

Private Sub Le()
Dim RdoAux As rdoResultset, Sql As String, sDoc As String
If Not bExec Then Exit Sub
grdMain.Clear
grdMain.Redraw = False
Sql = "SELECT * "
Sql = Sql & " FROM vwFULLOBRA WHERE YEAR(DT_ATUALIZACAO)=" & Val(cmbAno.Text) & " AND MONTH(DT_ATUALIZACAO)=" & cmbMes.ListIndex + 1
Sql = Sql & " ORDER BY ID_ALVARA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        
        If RdoAux!CS_TIP_RESP = 1 Then
            sDoc = Format(!ID_RESPONSAVEL, "00\.000\.000/0000-00")
        Else
            sDoc = Format(!ID_RESPONSAVEL, "000\.000\.000-00")
        End If
        
        With grdMain
            .AddRow
            .CellDetails .Rows, 1, IIf(RdoAux!CS_TIP_RESP = 1, "CNPJ", "CPF"), DT_DT_LEFT
            .CellDetails .Rows, 2, sDoc
            .CellDetails .Rows, 3, RdoAux!NM_RESPONSAVEL
            .CellDetails .Rows, 4, RdoAux!NM_OBRA
            .CellDetails .Rows, 5, Format(RdoAux!DT_ATUALIZACAO, "dd/mm/yyyy"), DT_CENTER
            .CellDetails .Rows, 6, RdoAux!ID_ALVARA
            .CellDetails .Rows, 7, RdoAux!DT_ALVARA, DT_CENTER
            .CellDetails .Rows, 8, RdoAux!ID_HABITESE
            .CellDetails .Rows, 9, RdoAux!NM_ENDERECO
            .CellDetails .Rows, 10, RdoAux!NU_ENDERECO, DT_RIGHT
        End With
       .MoveNext
    Loop
End With
grdMain.Redraw = True

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   grdMain.Enabled = True
   cmbAno.Enabled = True
   cmbMes.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdMain.Enabled = False
   cmbAno.Enabled = False
   cmbMes.Enabled = False
End If

End Sub

Private Sub TelaAtiva(nTela As Integer)
Dim x As Integer

For x = 0 To 3
    If x = nTela Then
        lblTit(x).ForeColor = &H0&
        frTela(x).Visible = True
        frTela(x).ZOrder 0
    Else
        lblTit(x).ForeColor = &H80&
        frTela(x).Visible = False
    End If
Next

End Sub

Private Sub grdMain_DblClick(ByVal lRow As Long, ByVal lCol As Long)
If grdMain.CellText(lRow, 1) <> "" Then
    cmdAlterar_Click
End If
End Sub

Private Sub lblTit_Click(Index As Integer)
TelaAtiva Index
End Sub

Private Sub Limpa()
cmbTipoDoc.ListIndex = 0
LimpaMascara mskCPF
LimpaMascara mskCNPJ
txtNome.Text = ""
End Sub

Private Sub txtCEP_LostFocus(Index As Integer)
Dim Sql As String, RdoAux As rdoResultset

txtUF.Text = ""
txtBairro.Text = ""
txtEndereco.Text = ""
If txtCEP(Index).Text <> "" Then
    Sql = "SELECT ID_TIPO_LOGR,TE_DESCRICAO_CEP,SG_UF,NM_BAIRRO FROM SO_CEP WHERE "
    Sql = Sql & "NU_CEP='" & txtCEP(Index).Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtUF.Text = !SG_UF
            txtBairro.Text = !NM_BAIRRO
            txtEndereco.Text = !ID_TIPO_LOGR & " " & !TE_DESCRICAO_CEP
        Else
            MsgBox "Cep não cadastrado.", vbCritical, "Erro"
            txtCEP(Index).SetFocus
        End If
       .Close
    End With
End If

End Sub

Private Sub txtCepCrit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdConsultar_Click
Else
    Tweak txtCepCrit, KeyAscii, IntegerPositive
End If
End Sub

Private Sub txtLogCep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdConsultar_Click
End If

End Sub
