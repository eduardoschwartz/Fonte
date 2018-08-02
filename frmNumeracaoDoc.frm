VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmNumeracaoDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de Documentos"
   ClientHeight    =   4755
   ClientLeft      =   3195
   ClientTop       =   3945
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   9090
   Begin Tributacao.jcFrames Tela2 
      Height          =   3660
      Left            =   45
      Top             =   585
      Width           =   8970
      _ExtentX        =   15822
      _ExtentY        =   6456
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
      Begin VB.Frame Frame2 
         Height          =   410
         Left            =   2925
         TabIndex        =   46
         Top             =   1170
         Width           =   3480
         Begin VB.OptionButton Opt 
            Caption         =   "Sim"
            Height          =   195
            Index           =   0
            Left            =   1890
            TabIndex        =   48
            Top             =   165
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Não"
            Height          =   195
            Index           =   1
            Left            =   2610
            TabIndex        =   47
            Top             =   165
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Abertura de Processo...:"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   49
            Top             =   135
            Width           =   1770
         End
      End
      Begin VB.Frame Frame1 
         Height          =   410
         Left            =   2925
         TabIndex        =   42
         Top             =   2745
         Width           =   3480
         Begin VB.OptionButton OptP 
            Caption         =   "Não"
            Height          =   195
            Index           =   1
            Left            =   2565
            TabIndex        =   44
            Top             =   165
            Value           =   -1  'True
            Width           =   600
         End
         Begin VB.OptionButton OptP 
            Caption         =   "Sim"
            Height          =   195
            Index           =   0
            Left            =   1845
            TabIndex        =   43
            Top             =   165
            Width           =   600
         End
         Begin VB.Label Label1 
            Caption         =   "Prorrogação de Prazo.:"
            Height          =   195
            Index           =   13
            Left            =   90
            TabIndex        =   45
            Top             =   135
            Width           =   1770
         End
      End
      Begin VB.Frame frDoc 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Documentos a serem apresentados"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   135
         TabIndex        =   35
         Top             =   2025
         Width           =   8745
         Begin VB.ListBox lstDoc 
            Appearance      =   0  'Flat
            Height          =   1155
            ItemData        =   "frmNumeracaoDoc.frx":0000
            Left            =   90
            List            =   "frmNumeracaoDoc.frx":0025
            Style           =   1  'Checkbox
            TabIndex        =   36
            Top             =   345
            Width           =   8565
         End
         Begin prjChameleon.chameleonButton cmdDoc 
            Height          =   240
            Left            =   3825
            TabIndex        =   11
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   690
            _ExtentX        =   1217
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
            MICON           =   "frmNumeracaoDoc.frx":01C4
            PICN            =   "frmNumeracaoDoc.frx":01E0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.TextBox txtFiscal 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   3195
         Width           =   1680
      End
      Begin VB.CheckBox chkCancel 
         Alignment       =   1  'Right Justify
         Caption         =   "CANCELADO...:"
         Height          =   195
         Left            =   7335
         TabIndex        =   14
         Top             =   2880
         Width           =   1500
      End
      Begin VB.ComboBox cmbTipoNot 
         Height          =   315
         ItemData        =   "frmNumeracaoDoc.frx":033A
         Left            =   1260
         List            =   "frmNumeracaoDoc.frx":0347
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2475
         Width           =   7620
      End
      Begin VB.ComboBox cmbPrazo 
         Height          =   315
         ItemData        =   "frmNumeracaoDoc.frx":03A4
         Left            =   1260
         List            =   "frmNumeracaoDoc.frx":03B1
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2835
         Width           =   1500
      End
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1260
         Width           =   1320
      End
      Begin VB.TextBox txtNome2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2295
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   495
         Width           =   6585
      End
      Begin VB.TextBox txtNome1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   2295
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         Width           =   6585
      End
      Begin VB.ComboBox cmbDesc 
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   855
         Width           =   7305
      End
      Begin VB.TextBox txtCod2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   5
         Top             =   495
         Width           =   960
      End
      Begin VB.TextBox txtCod1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         MaxLength       =   6
         TabIndex        =   4
         Top             =   135
         Width           =   960
      End
      Begin esMaskEdit.esMaskedEdit mskDataIni 
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Top             =   1260
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         MouseIcon       =   "frmNumeracaoDoc.frx":03D0
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
      Begin esMaskEdit.esMaskedEdit mskPeriodo1 
         Height          =   285
         Left            =   2925
         TabIndex        =   9
         Top             =   1665
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         MouseIcon       =   "frmNumeracaoDoc.frx":03EC
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
      Begin esMaskEdit.esMaskedEdit mskPeriodo2 
         Height          =   285
         Left            =   5220
         TabIndex        =   10
         Top             =   1665
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         MouseIcon       =   "frmNumeracaoDoc.frx":0408
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
      Begin esMaskEdit.esMaskedEdit mskRecebimento 
         Height          =   285
         Left            =   2925
         TabIndex        =   15
         Top             =   3195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         MouseIcon       =   "frmNumeracaoDoc.frx":0424
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
      Begin esMaskEdit.esMaskedEdit mskVencimento 
         Height          =   285
         Left            =   5220
         TabIndex        =   16
         Top             =   3195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   503
         MouseIcon       =   "frmNumeracaoDoc.frx":0440
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
      Begin prjChameleon.chameleonButton cmdRef1 
         Height          =   270
         Left            =   8595
         TabIndex        =   50
         ToolTipText     =   "Atualiza lista"
         Top             =   855
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   476
         BTYPE           =   14
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
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmNumeracaoDoc.frx":045C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Fiscal...:"
         Height          =   240
         Left            =   6480
         TabIndex        =   40
         Top             =   3255
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Vencimento:"
         Height          =   195
         Index           =   15
         Left            =   4275
         TabIndex        =   38
         Top             =   3255
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Data Recebimento do Documento.....:"
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   37
         Top             =   3255
         Width           =   2760
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Notifica..:"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   34
         Top             =   2520
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Prazo.............:"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   33
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Data Final..:"
         Height          =   195
         Index           =   10
         Left            =   4275
         TabIndex        =   32
         Top             =   1710
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Data Inicial..:"
         Height          =   195
         Index           =   9
         Left            =   1935
         TabIndex        =   31
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Período de Apuração...:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   30
         Top             =   1710
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Processo nº..:"
         Height          =   195
         Index           =   8
         Left            =   6480
         TabIndex        =   29
         Top             =   1305
         Width           =   1050
      End
      Begin VB.Label Label1 
         Caption         =   "Data Emissão.:"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   24
         Top             =   1326
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descrição......:"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   23
         Top             =   944
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Referência.....:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   22
         Top             =   562
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "Código Contr..:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   21
         Top             =   180
         Width           =   1140
      End
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   7890
      TabIndex        =   27
      ToolTipText     =   "Cancelar Edição"
      Top             =   4320
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
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
      MICON           =   "frmNumeracaoDoc.frx":0478
      PICN            =   "frmNumeracaoDoc.frx":0494
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
      Height          =   360
      Left            =   6735
      TabIndex        =   28
      ToolTipText     =   "Gravar os Dados"
      Top             =   4320
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmNumeracaoDoc.frx":05EE
      PICN            =   "frmNumeracaoDoc.frx":060A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   465
      Left            =   45
      Top             =   45
      Width           =   8970
      _ExtentX        =   15822
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
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmNumeracaoDoc.frx":09AF
         Left            =   1845
         List            =   "frmNumeracaoDoc.frx":09B1
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2985
      End
      Begin VB.TextBox txtAno 
         Height          =   330
         Left            =   5535
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   720
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   330
         Left            =   6255
         TabIndex        =   2
         Top             =   75
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2005
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtAno"
         BuddyDispid     =   196630
         OrigLeft        =   5790
         OrigTop         =   45
         OrigRight       =   6045
         OrigBottom      =   375
         Max             =   2020
         Min             =   2005
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Doc..: 123"
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
         Left            =   7290
         TabIndex        =   39
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Documento...:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   18
         Top             =   135
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Ano..:"
         Height          =   195
         Index           =   1
         Left            =   5040
         TabIndex        =   17
         Top             =   165
         Width           =   420
      End
   End
   Begin MSComctlLib.ListView lvDoc 
      Height          =   3660
      Left            =   45
      TabIndex        =   3
      Top             =   585
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   6456
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Núm."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Referên."
         Object.Width           =   1552
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Descrição"
         Object.Width           =   4482
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Dt.Emissão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Fiscal"
         Object.Width           =   4127
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   360
      Left            =   1230
      TabIndex        =   19
      ToolTipText     =   "Editar Registro"
      Top             =   4320
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
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
      MICON           =   "frmNumeracaoDoc.frx":09B3
      PICN            =   "frmNumeracaoDoc.frx":09CF
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
      Height          =   360
      Left            =   90
      TabIndex        =   20
      ToolTipText     =   "Novo Registro"
      Top             =   4320
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
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
      MICON           =   "frmNumeracaoDoc.frx":0B29
      PICN            =   "frmNumeracaoDoc.frx":0B45
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
Attribute VB_Name = "frmNumeracaoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec As Boolean
Dim sEvento As String

Private Sub cmbTipo_Click()

If Not bExec Then Exit Sub
If cmbTipo.ListIndex = -1 Then Exit Sub

CarregaDesc
CarregaLista
End Sub

Private Sub CarregaDesc()
Dim RdoAux As rdoResultset, Sql As String

cmbDesc.Clear
Sql = "select seq,descricao from tipodocumentodesc where codtipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " order by seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbDesc.AddItem !DESCRICAO
        cmbDesc.ItemData(cmbDesc.NewIndex) = !Seq
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdAlterar_Click()
sEvento = "Alterar"
Eventos "INCLUIR"
Limpa
Le
lblNum.Visible = True
lvDoc.Visible = False
Tela2.Visible = True
End Sub

Private Sub cmdCancel_Click()
lblNum.Visible = False
lvDoc.Visible = True
Tela2.Visible = False
Eventos "INICIAR"
End Sub

Private Sub cmdDoc_Click()
If cmdDoc.Value = True Then
    frDoc.Height = 1590
    frDoc.ZOrder 0
Else
    frDoc.Height = 330
End If

End Sub

Private Sub cmdGravar_Click()

If txtNome1.Text = "" Or txtNome2.Text = "" Then
    MsgBox "Digite os códigos de contribuinte e referência.", vbCritical, "Atenção"
    Exit Sub
End If

If mskDataIni.ClipText <> "" Then
    If Not IsDate(mskDataIni.Text) Then
        MsgBox "Data da emissão inválida.", vbCritical, "Atenção"
        Exit Sub
    End If
End If

If cmbDesc.ListIndex = -1 Then
    MsgBox "Selecione uma descrição.", vbCritical, "Atenção"
    Exit Sub
End If

If Opt(1).Value = True And Trim(txtNumProc.Text) = "" Then
    MsgBox "Digite o nº do processo.", vbCritical, "Atenção"
    Exit Sub
End If

Grava

lvDoc.Visible = True
Tela2.Visible = False
lblNum.Visible = False
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
sEvento = "Novo"
Eventos "INCLUIR"
txtNumProc.Enabled = False
txtNumProc.BackColor = Kde
Limpa
mskDataIni.Text = Format(Now, "dd/mm/yyyy")
txtFiscal.Text = NomeDeLogin
lblNum.Visible = True
lvDoc.Visible = False
Tela2.Visible = True
txtCod1.SetFocus
End Sub

Private Sub cmdRef1_Click()
CarregaDesc
End Sub

Private Sub Form_Load()
Dim RdoAux As rdoResultset, Sql As String
lblNum.Visible = False
lvDoc.Visible = True
Tela2.Visible = False
Eventos "INICIAR"
bExec = False
Sql = "select codigo,nome from tipolancdoc order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !nome
        cmbTipo.ItemData(cmbTipo.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
txtAno.Text = Year(Now)
cmbTipo.ListIndex = 0

bExec = True
cmbTipo_Click
Centraliza Me


End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
    cmdNovo.Visible = True
    cmdAlterar.Visible = True
    cmdRef1.Enabled = False
    cmdGravar.Visible = False
    cmdCancel.Visible = False
    For Each Ct In frmNumeracaoDoc
        If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Then
           Ct.BackColor = Kde
           Ct.Enabled = False
        End If
    Next
    cmbTipo.Enabled = True
    cmbTipo.BackColor = vbWhite
    txtAno.Enabled = True
    txtAno.BackColor = vbWhite
    UpDown1.Enabled = True
    txtNome1.Enabled = True
    txtNome2.Enabled = True
    txtNome1.Locked = True
    txtNome1.BackColor = Kde
    txtNome2.Locked = True
    txtNome2.BackColor = Kde
    txtFiscal.Locked = True
    txtFiscal.BackColor = Kde
ElseIf Tipo = "INCLUIR" Then
    cmdNovo.Visible = False
    cmdAlterar.Visible = False
    cmdRef1.Enabled = True
    cmdGravar.Visible = True
    cmdCancel.Visible = True
    For Each Ct In frmNumeracaoDoc
        If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Then
           Ct.BackColor = vbWhite
           Ct.Enabled = True
        End If
    Next
    cmbTipo.Enabled = False
    cmbTipo.BackColor = Kde
    txtAno.Enabled = False
    txtAno.BackColor = Kde
    UpDown1.Enabled = False
    txtNome1.Locked = True
    txtNome1.BackColor = Kde
    txtNome2.Locked = True
    txtNome2.BackColor = Kde
    txtFiscal.Locked = True
    txtFiscal.BackColor = Kde
End If

If cmbTipo.ListIndex = 0 Then
    cmbTipoNot.Enabled = False
    cmbTipoNot.BackColor = Kde
ElseIf cmbTipo.ListIndex = 1 Then
    cmdDoc.Enabled = False
    mskPeriodo1.Enabled = False
    mskPeriodo2.Enabled = False
    mskPeriodo1.BackColor = Kde
    mskPeriodo2.BackColor = Kde
End If

End Sub

Private Sub mskDataFim_GotFocus()
mskDataFim.SetFocus
mskDataFim.SelStart = 0
mskDataFim.SelLength = Len(mskDataFim.Text)
End Sub

Private Sub mskDataIni_GotFocus()
mskDataIni.SetFocus
mskDataIni.SelStart = 0
mskDataIni.SelLength = Len(mskDataIni.Text)
End Sub

Private Sub mskPeriodo1_GotFocus()
mskPeriodo1.SetFocus
mskPeriodo1.SelStart = 0
mskPeriodo1.SelLength = Len(mskPeriodo1.Text)
End Sub

Private Sub mskPeriodo2_GotFocus()
mskPeriodo2.SetFocus
mskPeriodo2.SelStart = 0
mskPeriodo2.SelLength = Len(mskPeriodo2.Text)
End Sub

Private Sub mskRecebimento_GotFocus()
mskRecebimento.SetFocus
mskRecebimento.SelStart = 0
mskRecebimento.SelLength = Len(mskRecebimento.Text)
End Sub

Private Sub mskVencimento_GotFocus()
mskVencimento.SetFocus
mskVencimento.SelStart = 0
mskVencimento.SelLength = Len(mskVencimento.Text)
End Sub

Private Sub Opt_Click(Index As Integer)
If Index = 0 Then
    txtNumProc.Text = ""
    txtNumProc.Enabled = False
    txtNumProc.BackColor = Kde
Else
    txtNumProc.Enabled = True
    txtNumProc.BackColor = Branco
End If
End Sub

Private Sub txtAno_Change()
CarregaLista
End Sub

Private Sub txtCod1_Change()
If txtNome1.Text <> "" Then txtNome1.Text = ""
End Sub

Private Sub txtCod1_KeyPress(KeyAscii As Integer)
Tweak txtCod1, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod1_LostFocus()
If Val(txtCod1.Text) > 0 Then txtNome1.Text = RetornaNome(CLng(txtCod1.Text))
If Val(txtCod2.Text) = 0 Then
    txtCod2.Text = txtCod1.Text
    txtCod2_LostFocus
End If
End Sub

Private Sub txtCod2_Change()
If txtNome2.Text <> "" Then txtNome2.Text = ""
End Sub

Private Sub txtCod2_KeyPress(KeyAscii As Integer)
Tweak txtCod2, KeyAscii, IntegerPositive
End Sub

Private Function RetornaNome(nCodReduz As Long) As String
Dim RdoAux As rdoResultset, Sql As String
RetornaNome = ""

If nCodReduz = 0 Then
    Exit Function
ElseIf nCodReduz < 100000 Then
    Sql = "select nomecidadao from vwfullimovel where codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            RetornaNome = !nomecidadao
        End If
       .Close
    End With
ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
    Sql = "select razaosocial from mobiliario where codigomob=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            RetornaNome = !razaosocial
        End If
       .Close
    End With
ElseIf nCodReduz > 500000 And nCodReduz < 700000 Then
    Sql = "select nomecidadao from cidadao where codcidadao=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            RetornaNome = !nomecidadao
        End If
       .Close
    End With
End If

If RetornaNome = "" And nCodReduz > 0 Then
    MsgBox "Código não cadastrado.", vbCritical, "Atenção"
End If

End Function

Private Sub txtCod2_LostFocus()
If Val(txtCod2.Text) > 0 Then txtNome2.Text = RetornaNome(CLng(txtCod2.Text))
End Sub

Private Sub Limpa()
Dim x As Integer

lblNum.Caption = "Nº Doc..: 000"
txtCod1.Text = ""
txtCod2.Text = ""
txtNome1.Text = ""
txtNome2.Text = ""
cmbDesc.ListIndex = -1
LimpaMascara mskDataIni
OptP(1).Value = True
Opt(0).Value = True
txtNumProc.Text = ""
LimpaMascara mskPeriodo1
LimpaMascara mskPeriodo2
cmbTipoNot.ListIndex = -1
cmbPrazo.ListIndex = -1

chkCancel.Value = vbUnchecked
LimpaMascara mskRecebimento
LimpaMascara mskVencimento
txtFiscal.Text = ""
frDoc.Height = 330

For x = 0 To lstDoc.ListCount - 1
    lstDoc.Selected(x) = False
Next

End Sub

Private Sub txtNumProc_LostFocus()
Dim Sql As String, RdoAux As rdoResultset, nNumProc As Long, nAnoProc As Integer
If txtNumProc.Text = "" Then Exit Sub

If ValidaProcesso2(txtNumProc.Text) = "Processo não Cadastrado." Then
    MsgBox "Nº de processo não cadastrado.", vbExclamation, "Atenção"
    txtNumProc.Text = ""
End If

End Sub

Private Sub Grava()
Dim Sql As String, RdoAux As rdoResultset, nCodTipo As Integer, nAno As Integer, nCod As Integer, nNumDoc As Integer
Dim nCod1 As Long, nCod2 As Long, nCodDesc As Integer, sDataEmissao As String, nProcNovo As Integer, sNumProc As String
Dim sDataApura1 As String, sDataApura2 As String, nPrazo As Integer, nProrroga As Integer, nCancel As Integer
Dim sDataRec As String, sDataVencto As String, sFiscal As String, aDoc() As Integer, x As Integer, nTipoNot As Integer

nCodTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
nAno = Val(txtAno.Text)

If sEvento = "Novo" Then
    Sql = "SELECT MAX(NUMDOC) AS MAXIMO FROM EMISSAODOCUMENTO WHERE TIPODOC=" & nCodTipo & " AND ANODOC=" & nAno
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!MAXIMO) Then
        nNumDoc = 1
    Else
        nNumDoc = RdoAux!MAXIMO + 1
    End If
    RdoAux.Close
Else
    nNumDoc = Mid(lblNum.Caption, InStr(1, lblNum.Caption, ":") + 2, Len(lblNum.Caption) - 10)
End If

nCod1 = Val(txtCod1.Text)
nCod2 = Val(txtCod2.Text)
nCodDesc = cmbDesc.ItemData(cmbDesc.ListIndex)
sDataEmissao = Format(Now, "dd/mm/yyyy")
nProcNovo = IIf(Opt(0).Value = True, 1, 0)
sNumProc = txtNumProc.Text
sDataApura1 = mskPeriodo1.Text
sDataApura2 = mskPeriodo2.Text
sDataRec = mskRecebimento.Text
sDataVencto = mskVencimento.Text
nPrazo = cmbPrazo.ItemData(cmbPrazo.ListIndex)
nProrroga = IIf(OptP(0).Value = True, 1, 0)
nCancel = IIf(chkCancel.Value = vbChecked, 1, 0)
sFiscal = txtFiscal.Text
If cmbTipoNot.ListIndex > -1 Then
    nTipoNot = cmbTipoNot.ItemData(cmbTipoNot.ListIndex)
Else
    nTipoNot = -1
End If


ReDim aDoc(0)
For x = 0 To lstDoc.ListCount - 1
    ReDim Preserve aDoc(UBound(aDoc) + 1)
    If (lstDoc.Selected(x)) Then
        aDoc(UBound(aDoc)) = 1
    Else
        aDoc(UBound(aDoc)) = 0
    End If
Next


If sEvento = "Novo" Then
    Sql = "INSERT EMISSAODOCUMENTO(TIPODOC,ANODOC,NUMDOC,COD1,COD2,CODDESC,DATAEMISSAO,PROCESSONOVO,PROCESSO,"
    Sql = Sql & "DATAAPURA1,DATAAPURA2,TIPONOTIF,TIPOPRAZO,PRORROGA,CANCEL,DATAREC,DATAVENCTO,FISCAL,DOC01,DOC02,"
    Sql = Sql & "DOC03,DOC04,DOC05,DOC06,DOC07,DOC08,DOC09,DOC10,DOC11) VALUES(" & nCodTipo & "," & nAno & ","
    Sql = Sql & nNumDoc & "," & nCod1 & "," & nCod2 & "," & nCodDesc & ",'" & Format(Now, "mm/dd/yyyy") & "',"
    Sql = Sql & nProcNovo & "," & IIf(txtNumProc.Text <> "", "'" & Mask(txtNumProc.Text) & "'", "Null") & ","
    Sql = Sql & IIf(IsDate(sDataApura1), "'" & Format(sDataApura1, "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & IIf(IsDate(sDataApura2), "'" & Format(sDataApura2, "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & IIf(nTipoNot > -1, nTipoNo, "Null") & "," & nPrazo & "," & nProrroga & "," & nCancel & ","
    Sql = Sql & IIf(IsDate(sDataRec), "'" & Format(sDataRec, "mm/dd/yyyy") & "'", "Null") & ","
    Sql = Sql & IIf(IsDate(sDataVencto), "'" & Format(sDataVencto, "mm/dd/yyyy") & "'", "Null") & ",'"
    Sql = Sql & sFiscal & "'," & aDoc(1) & "," & aDoc(2) & "," & aDoc(3) & "," & aDoc(4) & "," & aDoc(5) & ","
    Sql = Sql & aDoc(6) & "," & aDoc(7) & "," & aDoc(8) & "," & aDoc(9) & "," & aDoc(10) & "," & aDoc(11) & ")"
Else

End If

cn.Execute Sql, rdExecDirect

End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset, x As Integer

Sql = "SELECT * FROM EMISSAODOCUMENTO WHERE TIPODOC=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND ANODOC=" & txtAno.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblNum.Caption = "Nº Doc..: " & Format(!NumDoc, "0000")
    txtCod2.Text = !COD2
    txtCod2_LostFocus
    txtCod1.Text = !COD1
    txtCod1_LostFocus
    For x = 0 To cmbDesc.ListCount - 1
        If cmbDesc.ItemData(x) = !coddesc Then
            cmbDesc.ListIndex = x
            Exit For
        End If
    Next
    mskDataIni.Text = Format(!DataEmissao, "dd/mm/yyyy")
    If !processonovo = True Then
        Opt(0).Value = True
        txtNumProc.Enabled = False
        txtNumProc.BackColor = Kde
    Else
        Opt(1).Value = True
        txtNumProc.Enabled = True
        txtNumProc.BackColor = Branco
    End If
    txtNumProc.Text = SubNull(!Processo)
    If IsDate(!dataapura1) Then
        mskPeriodo1.Text = Format(!dataapura1, "dd/mm/yyyy")
    End If
    If IsDate(!dataapura2) Then
        mskPeriodo2.Text = Format(!dataapura2, "dd/mm/yyyy")
    End If
    If Not IsNull(!tiponotif) Then
        For x = 0 To cmbTipoNot.ListCount - 1
            If cmbTipoNot.ItemData(x) = !tiponotif Then
                cmbTipoNot.ListIndex = x
                Exit For
            End If
        Next
    End If
    cmbPrazo.ListIndex = !tipoprazo
    If !prorroga = True Then
        OptP(0).Value = True
    Else
        OptP(1).Value = True
    End If
    If !Cancel = 1 Then
        chkCancel.Value = vbChecked
    Else
        chkCancel.Value = vbUnchecked
    End If
    If Not IsNull(!datarec) Then
        mskRecebimento.Text = Format(!datarec, "dd/mm/yyyy")
    End If
    If Not IsNull(!DataVencto) Then
        mskVencimento.Text = Format(!DataVencto, "dd/mm/yyyy")
    End If
    txtFiscal.Text = SubNull(!fiscal)
    lstDoc.Selected(0) = !doc01
    lstDoc.Selected(1) = !doc02
    lstDoc.Selected(2) = !doc03
    lstDoc.Selected(3) = !doc04
    lstDoc.Selected(4) = !doc05
    lstDoc.Selected(5) = !doc06
    lstDoc.Selected(6) = !doc07
    lstDoc.Selected(7) = !doc08
    lstDoc.Selected(8) = !doc09
    lstDoc.Selected(9) = !doc10
    lstDoc.Selected(10) = !doc11
   .Close
End With

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, z As Long, itmX As ListItem
If bExec = False Then Exit Sub

z = SendMessage(lvDoc.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT emissaodocumento.tipodoc, emissaodocumento.anodoc, emissaodocumento.numdoc, emissaodocumento.cod1, emissaodocumento.cod2,"
Sql = Sql & "emissaodocumento.coddesc, emissaodocumento.dataemissao, emissaodocumento.processonovo, emissaodocumento.processo, emissaodocumento.dataapura1,"
Sql = Sql & "emissaodocumento.dataapura2, emissaodocumento.tiponotif, emissaodocumento.tipoprazo, emissaodocumento.prorroga, emissaodocumento.cancel,"
Sql = Sql & "emissaodocumento.datarec, emissaodocumento.datavencto, emissaodocumento.fiscal, emissaodocumento.doc01, emissaodocumento.doc02,"
Sql = Sql & "emissaodocumento.doc03, emissaodocumento.doc04, emissaodocumento.doc05, emissaodocumento.doc06, emissaodocumento.doc07, emissaodocumento.doc08,"
Sql = Sql & "emissaodocumento.doc09 , emissaodocumento.doc10, emissaodocumento.doc11, tipodocumentodesc.descricao FROM emissaodocumento INNER JOIN "
Sql = Sql & "tipodocumentodesc ON emissaodocumento.tipodoc = tipodocumentodesc.codtipo WHERE TIPODOC=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND ANODOC=" & txtAno.Text
Sql = Sql & " ORDER BY emissaodocumento.numdoc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Set itmX = lvDoc.ListItems.Add(, , Format(!NumDoc, "0000"))
        itmX.SubItems(1) = Format(!COD1, "000000")
        itmX.SubItems(2) = Format(!COD2, "000000")
        itmX.SubItems(3) = !DESCRICAO
        itmX.SubItems(4) = Format(!DataEmissao, "dd/mm/yyyy")
        itmX.SubItems(5) = !fiscal
       .MoveNext
    Loop
   .Close
End With

End Sub
