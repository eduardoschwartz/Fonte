VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{B60B1875-E5CA-11D2-BC3D-78A407C10000}#1.0#0"; "ksdpanel.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmReparcelamento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parcelamento de Divida Fiscal"
   ClientHeight    =   6270
   ClientLeft      =   1545
   ClientTop       =   3750
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   11265
   Begin KSDPanel.Panel frCod 
      Height          =   4935
      Left            =   1770
      TabIndex        =   47
      Top             =   480
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   8705
      Caption         =   "Localiza débitos a serem reparcelados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   5
      TextAlign       =   1
      ForeColor       =   12648447
      BackColor       =   4210752
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   930
         TabIndex        =   53
         Top             =   435
         Width           =   1635
      End
      Begin prjChameleon.chameleonButton cmdCancel 
         Height          =   375
         Left            =   5910
         TabIndex        =   48
         ToolTipText     =   "Cancelar"
         Top             =   4485
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReparcelamento.frx":0000
         PICN            =   "frmReparcelamento.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSairCod 
         Height          =   375
         Left            =   5430
         TabIndex        =   49
         ToolTipText     =   "Retornar Débitos Selecionados"
         Top             =   4485
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   ""
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReparcelamento.frx":0176
         PICN            =   "frmReparcelamento.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRemoveAll 
         Height          =   375
         Left            =   4470
         TabIndex        =   50
         ToolTipText     =   "Remover Todos"
         Top             =   4485
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReparcelamento.frx":027D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAddAll 
         Height          =   375
         Left            =   4950
         TabIndex        =   51
         ToolTipText     =   "Selecionar Todos"
         Top             =   4485
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReparcelamento.frx":0299
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvDeb 
         Height          =   3555
         Left            =   90
         TabIndex        =   52
         Top             =   855
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   6271
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ano"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Lanc"
            Object.Width           =   2998
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Seq"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Pc."
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Co."
            Object.Width           =   811
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Vencto."
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Data Base"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Ajuiz"
            Object.Width           =   776
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "DA"
            Object.Width           =   774
         EndProperty
      End
      Begin VB.Label lblNome 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   56
         Top             =   450
         Width           =   3675
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
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
         Left            =   180
         TabIndex        =   55
         Top             =   495
         Width           =   645
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecione os débitos a serem reparcelados."
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Left            =   150
         TabIndex        =   54
         Top             =   4545
         Width           =   3225
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Desconto de Juros e Multa"
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   7950
      TabIndex        =   73
      Top             =   3780
      Width           =   3285
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Left            =   2820
         TabIndex        =   78
         Top             =   270
         Width           =   255
      End
      Begin VB.Label lblValorPlano 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   2460
         TabIndex        =   77
         Top             =   270
         Width           =   315
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto de :"
         Height          =   225
         Left            =   1440
         TabIndex        =   76
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lblTipoPlano 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   75
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Plano:"
         Height          =   225
         Left            =   240
         TabIndex        =   74
         Top             =   270
         Width           =   435
      End
   End
   Begin VB.Frame fr5 
      BackColor       =   &H00EEEEEE&
      Height          =   1830
      Left            =   7950
      TabIndex        =   44
      Top             =   4410
      Width           =   3300
      Begin VB.ComboBox cmbResp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   225
         Width           =   1755
      End
      Begin prjChameleon.chameleonButton cmdAddDebito 
         Height          =   555
         Left            =   990
         TabIndex        =   10
         ToolTipText     =   "Adiciona Débitos a serem Reparcelados"
         Top             =   705
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "&Adicionar/Remover Débito"
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
         MICON           =   "frmReparcelamento.frx":02B5
         PICN            =   "frmReparcelamento.frx":02D1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdChangeUser 
         Height          =   345
         Left            =   1980
         TabIndex        =   46
         ToolTipText     =   "Altera o Responsável"
         Top             =   210
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Alterar"
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
         MICON           =   "frmReparcelamento.frx":07B4
         PICN            =   "frmReparcelamento.frx":07D0
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
         Left            =   2250
         TabIndex        =   65
         ToolTipText     =   "Fechar a Tela"
         Top             =   1395
         Width           =   930
         _ExtentX        =   1640
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
         MICON           =   "frmReparcelamento.frx":092A
         PICN            =   "frmReparcelamento.frx":0946
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
   Begin KSDPanel.Panel fr2 
      Height          =   3795
      Left            =   7950
      TabIndex        =   14
      Top             =   0
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   6694
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
      BorderStyle     =   6
      BackColor       =   15658734
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Diligências"
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   90
         TabIndex        =   70
         Top             =   3180
         Width           =   3105
         Begin VB.TextBox txtQtdeDil 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1155
            MaxLength       =   3
            TabIndex        =   9
            Text            =   "0"
            Top             =   195
            Width           =   465
         End
         Begin VB.Label lblSomaDil 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0,00"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   2115
            TabIndex        =   72
            Top             =   225
            Width           =   840
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Quantidade...:"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   71
            Top             =   240
            Width           =   1035
         End
      End
      Begin esMaskEdit.esMaskedEdit mskDataProc 
         Height          =   285
         Left            =   1815
         TabIndex        =   1
         Top             =   660
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   15658734
         ForeColor       =   12582912
         MouseIcon       =   "frmReparcelamento.frx":0AA0
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
         BorderStyle     =   0
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
         Locked          =   -1  'True
      End
      Begin VB.TextBox txtPercEntrada 
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   2220
         Width           =   1275
      End
      Begin VB.TextBox txtValorEntrada 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Top             =   1890
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cálculo de:"
         ForeColor       =   &H00800000&
         Height          =   585
         Left            =   90
         TabIndex        =   20
         Top             =   2580
         Width           =   3105
         Begin VB.CheckBox chkHon 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "Honorários"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1905
            TabIndex        =   8
            Top             =   270
            Width           =   1080
         End
         Begin VB.CheckBox chkJuros 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "Juros"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   930
            TabIndex        =   7
            Top             =   270
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.CheckBox chkMulta 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Caption         =   "Multa"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   150
            TabIndex        =   6
            Top             =   270
            Value           =   1  'Checked
            Width           =   735
         End
      End
      Begin VB.TextBox txtQtdeParc 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1260
         Width           =   1275
      End
      Begin VB.TextBox txtNumProc 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   330
         Width           =   1275
      End
      Begin esMaskEdit.esMaskedEdit mskVencto 
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   1575
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         MouseIcon       =   "frmReparcelamento.frx":0ABC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   10
         Mask            =   "99/99/9999"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3240
         Width           =   3045
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "% de Entrada..............:"
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   39
         Top             =   2310
         Width           =   1665
      End
      Begin VB.Label lblDataParc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   33
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor de Entrada.........:"
         Height          =   225
         Index           =   6
         Left            =   90
         TabIndex        =   26
         Top             =   1980
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do 1º Vencto......:"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   21
         Top             =   1650
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de Parcelas........:"
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   19
         Top             =   1320
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Parcelamento:"
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   18
         Top             =   990
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Processo.......:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo..........:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   16
         Top             =   390
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000080&
         Caption         =   " Dados do Processo"
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
         Index           =   1
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   3210
      End
   End
   Begin KSDPanel.Panel fr4 
      Height          =   3405
      Left            =   6300
      TabIndex        =   27
      Top             =   2850
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   6006
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
      BorderStyle     =   6
      BackColor       =   15658734
      Begin prjChameleon.chameleonButton cmdGrava 
         Height          =   1005
         Left            =   150
         TabIndex        =   42
         ToolTipText     =   "Grava o Reparcelamento e Imprime Carne"
         Top             =   1530
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1773
         BTYPE           =   3
         TX              =   "G&ravar o Reparcelamento"
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
         MICON           =   "frmReparcelamento.frx":0AD8
         PICN            =   "frmReparcelamento.frx":0AF4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGeraDebito 
         Height          =   1005
         Left            =   150
         TabIndex        =   43
         ToolTipText     =   "Gera os Débitos na Tela"
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1773
         BTYPE           =   3
         TX              =   "&Gerar Débitos"
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
         MICON           =   "frmReparcelamento.frx":1073
         PICN            =   "frmReparcelamento.frx":108F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblValorParcela 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   150
         TabIndex        =   80
         Top             =   2970
         Width           =   1275
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela R$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   150
         TabIndex        =   79
         Top             =   2700
         Width           =   1185
      End
   End
   Begin KSDPanel.Panel fr1 
      Height          =   2850
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   5027
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
      BorderStyle     =   6
      BackColor       =   15658734
      Begin MSFlexGridLib.MSFlexGrid grdOrigem 
         Height          =   2505
         Left            =   45
         TabIndex        =   13
         Top             =   285
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   4419
         _Version        =   393216
         Rows            =   1
         Cols            =   13
         FixedCols       =   0
         BackColorFixed  =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmReparcelamento.frx":121E
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000080&
         Caption         =   " Débitos a serem reparcelados"
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
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   30
         Width           =   7845
      End
   End
   Begin KSDPanel.Panel fr3 
      Height          =   2595
      Left            =   0
      TabIndex        =   22
      Top             =   3675
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4577
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
      BorderStyle     =   6
      Begin MSFlexGridLib.MSFlexGrid grdDestino 
         Height          =   1740
         Left            =   30
         TabIndex        =   23
         Top             =   270
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   3069
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedCols       =   0
         BackColorFixed  =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"frmReparcelamento.frx":12BC
      End
      Begin KSDPanel.Panel Panel2 
         Height          =   555
         Left            =   60
         TabIndex        =   34
         Top             =   2010
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   979
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
         BorderStyle     =   5
         BackColor       =   16777215
         Begin VB.Label lblValorJurosGer 
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
            Left            =   2145
            TabIndex        =   64
            Top             =   285
            Width           =   750
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor dos Juros..........:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   63
            Top             =   285
            Width           =   1980
         End
         Begin VB.Label lblValorMultaGer 
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
            Left            =   4995
            TabIndex        =   62
            Top             =   270
            Width           =   765
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor da Multa......:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3030
            TabIndex        =   61
            Top             =   285
            Width           =   1680
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Parcelas Geradas.......:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   135
            TabIndex        =   38
            Top             =   60
            Width           =   2055
         End
         Begin VB.Label lblNumParcGer 
            BackStyle       =   0  'Transparent
            Caption         =   "000"
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
            Left            =   2565
            TabIndex        =   37
            Top             =   60
            Width           =   435
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Val.Corr+Hon+Dil..:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   3030
            TabIndex        =   36
            Top             =   60
            Width           =   1680
         End
         Begin VB.Label lblValorParcGer 
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
            Left            =   4770
            TabIndex        =   35
            Top             =   60
            Width           =   990
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00000080&
         Caption         =   " Débitos a serem criados"
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
         Index           =   2
         Left            =   30
         TabIndex        =   24
         Top             =   30
         Width           =   6195
      End
   End
   Begin KSDPanel.Panel Panel1 
      Height          =   795
      Left            =   15
      TabIndex        =   28
      Top             =   2865
      Width           =   6240
      _ExtentX        =   11007
      _ExtentY        =   1402
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
      BorderStyle     =   5
      BackColor       =   16777215
      Begin VB.Label lblValorHon 
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
         Left            =   2145
         TabIndex        =   69
         Top             =   510
         Width           =   750
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor dos Honorários..:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   68
         Top             =   510
         Width           =   1980
      End
      Begin VB.Label lblValorDil 
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
         Left            =   4995
         TabIndex        =   67
         Top             =   495
         Width           =   765
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Diligências...:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3030
         TabIndex        =   66
         Top             =   510
         Width           =   1680
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor da Multa......:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3030
         TabIndex        =   60
         Top             =   270
         Width           =   1680
      End
      Begin VB.Label lblValorMulta 
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
         Left            =   4995
         TabIndex        =   59
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor dos Juros..........:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   135
         TabIndex        =   58
         Top             =   270
         Width           =   1980
      End
      Begin VB.Label lblValorJuros 
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
         Left            =   2145
         TabIndex        =   57
         Top             =   285
         Width           =   750
      End
      Begin VB.Label lblValorParc 
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
         Left            =   4770
         TabIndex        =   32
         Top             =   45
         Width           =   990
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Corrigido......:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3030
         TabIndex        =   31
         Top             =   45
         Width           =   1695
      End
      Begin VB.Label lblNumParc 
         BackStyle       =   0  'Transparent
         Caption         =   "000"
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
         Left            =   2565
         TabIndex        =   30
         Top             =   45
         Width           =   435
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcelas Selecionadas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   29
         Top             =   45
         Width           =   2025
      End
   End
   Begin VB.Label lblResp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   90
      TabIndex        =   41
      Top             =   840
      Width           =   3075
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código do Responsável pelo Parcelamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   150
      Width           =   3075
   End
End
Attribute VB_Name = "frmReparcelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nCod As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    sVencto As String
    sDataBase As String
    bAjuizado As Boolean
    bDA As Boolean
End Type

Dim sTipoCod As String
Dim aDebito() As Debito
Dim aDebitoTmp() As Debito
Dim xMark As Boolean
Dim nValorDil As Double
Dim RdoAux As rdoResultset, Sql As String
Dim itmX As ListItem
Dim nSomaAjuizado As Double
Dim z As Long, bGerado As Boolean

Private Sub chkHon_Click()

If grdOrigem.Rows > 1 Then
    If chkHon.Value = 1 Then
       lblValorHon.Caption = FormatNumber(nSomaAjuizado * 0.1, 2)
    Else
       lblValorHon.Caption = "0,00"
    End If
End If

End Sub

Private Sub chkJuros_Click()
AtualizaTotal
End Sub

Private Sub chkMulta_Click()
AtualizaTotal
End Sub

Private Sub cmbResp_Click()
Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & Val(cmbResp.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
         lblResp.Caption = !NOMECIDADAO
    End If
End With

End Sub

Private Sub cmdAddAll_Click()
Dim n As Integer
Dim nCod As Long
Dim nAno As Integer
Dim nLanc As Integer
Dim nSeq As Integer
Dim nParc As Integer
Dim nCompl As Integer
Dim sVencto As String
Dim sDataBase As String

For n = 1 To lvDeb.ListItems.Count
    lvDeb.ListItems(n).Checked = True
Next

'Remove da matriz
ReDim aDebitoTmp(0)
For n = 0 To UBound(aDebito)
    If aDebito(n).nCod <> Val(txtCod.text) Then
       If aDebito(n).nCod > 0 Then
          ReDim Preserve aDebitoTmp(UBound(aDebitoTmp) + 1)
          aDebitoTmp(UBound(aDebitoTmp)).nCod = aDebito(n).nCod
          aDebitoTmp(UBound(aDebitoTmp)).nAno = aDebito(n).nAno
          aDebitoTmp(UBound(aDebitoTmp)).nLanc = aDebito(n).nLanc
          aDebitoTmp(UBound(aDebitoTmp)).nSeq = aDebito(n).nSeq
          aDebitoTmp(UBound(aDebitoTmp)).nParc = aDebito(n).nParc
          aDebitoTmp(UBound(aDebitoTmp)).nCompl = aDebito(n).nCompl
          aDebitoTmp(UBound(aDebitoTmp)).sVencto = aDebito(n).sVencto
          aDebitoTmp(UBound(aDebitoTmp)).sDataBase = aDebito(n).sDataBase
       End If
    End If
Next
'Reverte a matriz
ReDim aDebito(UBound(aDebitoTmp))
For n = 0 To UBound(aDebitoTmp)
    aDebito(n).nCod = aDebitoTmp(n).nCod
    aDebito(n).nAno = aDebitoTmp(n).nAno
    aDebito(n).nLanc = aDebitoTmp(n).nLanc
    aDebito(n).nSeq = aDebitoTmp(n).nSeq
    aDebito(n).nParc = aDebitoTmp(n).nParc
    aDebito(n).nCompl = aDebitoTmp(n).nCompl
    aDebito(n).sVencto = aDebitoTmp(n).sVencto
    aDebito(n).sDataBase = aDebitoTmp(n).sDataBase
Next

'adiciona a matriz

For n = 1 To lvDeb.ListItems.Count
    nCod = Val(txtCod.text)
    nAno = Val(lvDeb.ListItems(n).text)
    nLanc = Left$(lvDeb.ListItems(n).ListSubItems(1).text, 2)
    nSeq = lvDeb.ListItems(n).ListSubItems(2)
    nParc = lvDeb.ListItems(n).ListSubItems(3)
    nCompl = lvDeb.ListItems(n).ListSubItems(4)
    sVencto = lvDeb.ListItems(n).ListSubItems(5)
    sDataBase = lvDeb.ListItems(n).ListSubItems(7)
     ReDim Preserve aDebito(UBound(aDebito) + 1)
     aDebito(UBound(aDebito)).nCod = nCod
     aDebito(UBound(aDebito)).nAno = nAno
     aDebito(UBound(aDebito)).nLanc = nLanc
     aDebito(UBound(aDebito)).nSeq = nSeq
     aDebito(UBound(aDebito)).nParc = nParc
     aDebito(UBound(aDebito)).nCompl = nCompl
     aDebito(UBound(aDebito)).sVencto = sVencto
     aDebito(UBound(aDebito)).sDataBase = sDataBase
Next

End Sub

Private Sub cmdAddDebito_Click()

If Val(txtQtdeParc.text) = 0 Then
    MsgBox "Digite a quantidade de parcelas.", vbExclamation, "Atenção"
    Exit Sub
ElseIf Val(txtQtdeParc.text) > 48 Then
    MsgBox "Parcelamento máximo em 48 vezes.", vbExclamation, "Atenção"
    Exit Sub
ElseIf Not IsDate(mskVencto.text) Then
    MsgBox "Digite a Data do 1º Vencimento.", vbExclamation, "Atenção"
    Exit Sub
End If

z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)
fr1.Enabled = False
fr3.Enabled = False
fr4.Enabled = False
fr5.Enabled = False
frCod.Visible = True
frCod.ZOrder 0

If grdOrigem.Rows = 1 Then
   xMark = False
   ReDim aDebito(0)
   txtCod.text = ""
   lblNome.Caption = ""
   lvDeb.SetFocus
Else
   xMark = True
   txtCod.text = Val(grdOrigem.TextMatrix(grdOrigem.Row, 0))
   txtCod_LostFocus
End If

End Sub

Private Sub cmdCancel_Click()
fr1.Enabled = True
fr3.Enabled = True
fr4.Enabled = True
fr5.Enabled = True
frCod.Visible = False

End Sub

Private Sub cmdChangeUser_Click()
Dim x As Long
'Altera Responsavel
If grdOrigem.Rows = 1 Then
   MsgBox "Selecione primeiro as parcelas.", vbExclamation, "Atenção"
   Exit Sub
End If

x = Val(InputBox("Digite o Código do Responsável pelo Reparcelamento.", "Alterar Responsável"))
If x > 0 Then
    If sTipoCod = "I" Then
        Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
        Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
        Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & x
    Else
        Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO Where CODIGOMOB =" & x
    End If
    
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
    With RdoAux
        If .RowCount = 0 Then
           MsgBox "Código não existe.", vbExclamation, "Atenção"
           Exit Sub
        Else
           If sTipoCod = "I" Then
              MsgBox "Responsável do Reparcelamento foi alterado para " & Format(x, "0000000") & " - " & !NOMECIDADAO
           ElseIf sTipoCod = "M" Then
              MsgBox "Responsável do Reparcelamento foi alterado para " & Format(x, "0000000") & " - " & !RAZAOSOCIAL
           End If
           cmbResp.AddItem Format(x, "0000000"), 0
           cmbResp.ListIndex = 0
        End If
    End With
End If

End Sub

Private Sub cmdGeraDebito_Click()
Dim df As Integer

'Gera Debitos
If grdDestino.Rows > 1 Then
   If MsgBox("Sobrescrever os Débitos Gerados ? ", vbQuestion + vbYesNo, "Atenção") = vbNo Then
      Exit Sub
   End If
End If

bAchou = False
For x = 1 To grdOrigem.Rows - 1
    If grdOrigem.TextMatrix(x, 7) = 0 Then
        bAchou = True
        Exit For
    End If
Next

If bAchou Then
    MsgBox "Não é possível reparcelar débitos zerados.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtQtdeParc.text) = 0 Then
    MsgBox "Digite a Qtde de Parcelas.", vbCritical, "Atenção"
    txtQtdeParc.SetFocus
    Exit Sub
End If

If Not IsDate(mskVencto.text) Then
    MsgBox "Data do 1º Vencimento inválido.", vbCritical, "Atenção"
    mskVencto.SetFocus
    Exit Sub
End If

df = ValidaFeriado(CDate(mskVencto.text))
If df = 1 Then
    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVencto.text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskVencto.text = Format(RetornaDiaUtil(CDate(mskVencto.text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 2 Then
    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVencto.text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskVencto.text = Format(RetornaDiaUtil(CDate(mskVencto.text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 3 Then
    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
    Sql = Sql & " Where DIA = " & Day(CDate(mskVencto.text))
    Sql = Sql & " AND MES=" & Month(CDate(mskVencto.text)) & " AND ANO=" & Year(CDate(mskVencto.text))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVencto.text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                mskVencto.text = RetornaDiaUtil(CDate(mskVencto.text))
            Else
                Exit Sub
            End If
          .Close
        End If
    End With
End If

'If Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) <> "ROSE" Then
    If CDate(mskVencto.text) < Format(Now, "dd/mm/yyyy") Then
        MsgBox "Data do 1º Vencimento tem que ser maior que a data atual.", vbCritical, "Atenção"
        mskVencto.SetFocus
        Exit Sub
    End If
'End If

If grdOrigem.Rows = 1 Then
   MsgBox "Selecione os débitos a serem reparcelados.", vbCritical, "Atenção"
   Exit Sub
End If

If CDbl(txtValorEntrada.text) >= (CDbl(lblValorParc.Caption) + CDbl(lblValorMulta.Caption) + CDbl(lblValorJuros.Caption)) Then
    MsgBox "O valor da entrada excede o valor do débito.", vbExclamation, "Atenção"
    Exit Sub
End If

If CDbl(txtPercEntrada.text) > 0 And (CDbl(txtPercEntrada.text) < 1 Or CDbl(txtPercEntrada.text) > 90) Then
    MsgBox "O Percentual de Entrada entrada deve estar entre 1 e 90.", vbExclamation, "Atenção"
    Exit Sub
End If

'Parse
grdDestino.Rows = 1
GeraDebito

End Sub

Private Sub cmdGrava_Click()
Dim sValidaProc As String

sValidaProc = ValidaProcesso(txtNumProc.text)
If sValidaProc <> "OK" Then
    MsgBox sValidaProc, vbCritical, "Atenção"
    Exit Sub
Else
    mskDataProc.text = Format(RetornaDataProcesso(Val(Left$(txtNumProc.text, Len(txtNumProc.text) - 5)), Val(Right$(txtNumProc.text, 4))), "dd/mm/yyyy")
End If

If grdDestino.Rows = 1 Then
    MsgBox "Favor gerar os débitos a serem criados.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataProc.text) Then
    MsgBox "Data do Processo inválido.", vbCritical, "Atenção"
    mskDataProc.SetFocus
    Exit Sub
End If

bGerado = True
GravaCarneTmp

If bGerado Then
    grdDestino.Rows = 1
    grdOrigem.Rows = 1
    txtNumProc.text = ""
    LimpaMascara mskDataProc
    lblDataParc.Caption = Format(Now, "dd/mm/yyyy")
    txtQtdeParc.text = ""
    LimpaMascara mskVencto
    txtValorEntrada.text = ""
    txtPercEntrada.text = ""
    cmbResp.Clear
    lblResp.Caption = ""
    lblNumParc.Caption = "000"
    lblValorParc.Caption = "0,00"
    lblValorJuros.Caption = "0,00"
    lblValorMulta.Caption = "0,00"
    lblValorHon.Caption = "0,00"
    lblValorDil.Caption = "0,00"
    lblNumParcGer.Caption = "000"
    lblValorJurosGer.Caption = "0,00"
    lblValorMultaGer.Caption = "0,00"
    lblValorParcGer.Caption = "0,00"
    txtQtdeDil.text = 0
    lblSomaDil.Caption = "0,00"
End If

End Sub


Private Sub cmdRemoveAll_Click()
Dim n As Integer

For n = 1 To lvDeb.ListItems.Count
    lvDeb.ListItems(n).Checked = False
Next

'Remove da matriz
ReDim aDebitoTmp(0)
For n = 0 To UBound(aDebito)
    If aDebito(n).nCod <> Val(txtCod.text) Then
       If aDebito(n).nCod > 0 Then
          ReDim Preserve aDebitoTmp(UBound(aDebitoTmp) + 1)
          aDebitoTmp(UBound(aDebitoTmp)).nCod = aDebito(n).nCod
          aDebitoTmp(UBound(aDebitoTmp)).nAno = aDebito(n).nAno
          aDebitoTmp(UBound(aDebitoTmp)).nLanc = aDebito(n).nLanc
          aDebitoTmp(UBound(aDebitoTmp)).nSeq = aDebito(n).nSeq
          aDebitoTmp(UBound(aDebitoTmp)).nParc = aDebito(n).nParc
          aDebitoTmp(UBound(aDebitoTmp)).nCompl = aDebito(n).nCompl
          aDebitoTmp(UBound(aDebitoTmp)).sVencto = aDebito(n).sVencto
          aDebitoTmp(UBound(aDebitoTmp)).sDataBase = aDebito(n).sDataBase
       End If
    End If
Next
'Reverte a matriz
ReDim aDebito(UBound(aDebitoTmp))
For n = 0 To UBound(aDebitoTmp)
    aDebito(n).nCod = aDebitoTmp(n).nCod
    aDebito(n).nAno = aDebitoTmp(n).nAno
    aDebito(n).nLanc = aDebitoTmp(n).nLanc
    aDebito(n).nSeq = aDebitoTmp(n).nSeq
    aDebito(n).nParc = aDebitoTmp(n).nParc
    aDebito(n).nCompl = aDebitoTmp(n).nCompl
    aDebito(n).sVencto = aDebitoTmp(n).sVencto
    aDebito(n).sDataBase = aDebitoTmp(n).sDataBase
Next

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSairCod_Click()

AdicionaDebitos
AtualizaTotal
AtualizaResp

End Sub

Private Sub Form_Load()
Ocupado

Centraliza Me
Liberado

'carrega valor da diligência
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & Year(Now) & " AND CODTRIBUTO=91"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
         nValorDil = FormatNumber(!VALORALIQ, 2)
     Else
         MsgBox "Taxa de Diligência não cadastrado para este ano.", vbExclamation, "Atenção"
     End If
    .Close
End With

lblDataParc.Caption = Mid$(frmMdi.Sbar.Panels(6).text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).text, 4)
txtValorEntrada.text = 0
txtPercEntrada.text = 0
frCod.Visible = False
frmMdi.AddWindow Me.Name, Me.Caption
End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True


End Sub

Public Sub CarregaDebito(nCodImovel As Long)
Dim Sql As String, RdoAux As rdoResultset, d As Integer
Dim nTotal As Double

z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,"
Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.CODCOMPLEMENTO,DATAAJUIZA,DATAINSCRICAO, "
Sql = Sql & "DEBITOPARCELA.STATUSLANC, LANCAMENTO.DESCREDUZ,SITUACAOLANCAMENTO.DescSituacao,DATAVENCIMENTO,DATADEBASE "
Sql = Sql & "FROM DEBITOPARCELA INNER JOIN LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
Sql = Sql & "Inner Join SITUACAOLANCAMENTO ON DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodImovel & " AND DEBITOPARCELA.CODLANCAMENTO<>20 AND STATUSLANC=3 AND NUMPARCELA>0 ORDER BY DATAVENCIMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
        Do Until .EOF
           nTotal = 0
           Sql = "SELECT TOTALLANCADO FROM VWCNSLANCAMENTO WHERE CODREDUZIDO=" & nCodImovel & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
           Sql = Sql & "CODLANCAMENTO =" & !CodLancamento & " And SEQLANCAMENTO = " & !SeqLancamento & " And NumParcela = " & !NumParcela & " And CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
           With RdoAux
               Do Until .EOF
                  nTotal = nTotal + !TOTALLANCADO
                 .MoveNext
               Loop
              .Close
           End With
           If !AnoExercicio = 2003 And !CodLancamento = 1 And !statuslanc > 2 Then GoTo PROXIMO
           If nTotal = 0 Then GoTo PROXIMO
           If !SeqLancamento > 99 Then GoTo PROXIMO
           Set itmX = lvDeb.ListItems.Add(, Format(!CODREDUZIDO, "000000") & !AnoExercicio & Format(!CodLancamento, "00") & Format(!SeqLancamento, "00") & Format(!NumParcela, "00") & Format(!CODCOMPLEMENTO, "00"), !AnoExercicio)
           itmX.SubItems(1) = Format(!CodLancamento, "00") & "-" & !DESCREDUZ
           itmX.SubItems(2) = Format(!SeqLancamento, "00")
           itmX.SubItems(3) = Format(!NumParcela, "00")
           itmX.SubItems(4) = Format(!CODCOMPLEMENTO, "00")
           itmX.SubItems(5) = Format(!DATAVENCIMENTO, "dd/mm/yyyy")
           itmX.SubItems(6) = !DescSituacao
           itmX.SubItems(7) = Format(!DATADEBASE, "dd/mm/yyyy")
           itmX.SubItems(8) = IIf(IsDate(!DATAAJUIZA), "S", "N")
           itmX.SubItems(9) = IIf(IsNull(!DATAINSCRICAO), "N", "S")
           If xMark Then
              For d = 1 To UBound(aDebito)
                 If !CODREDUZIDO = aDebito(d).nCod And !AnoExercicio = aDebito(d).nAno And !CodLancamento = aDebito(d).nLanc And _
                    !SeqLancamento = aDebito(d).nSeq And !NumParcela = aDebito(d).nParc And !CODCOMPLEMENTO = aDebito(d).nCompl Then
                    lvDeb.ListItems(lvDeb.ListItems.Count).Checked = True
                    Exit For
                 End If
              Next
           End If
PROXIMO:
          .MoveNext
        Loop
    Else
        MsgBox "Não existem Débitos a serem Reparcelados.", vbExclamation, "Atenção"
        txtCod.SetFocus
        txtCod.SelStart = 0
        txtCod.SelLength = Len(txtCod.text)
    End If
   .Close
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
End Sub


Private Sub lvDeb_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim nCod As Long
Dim nAno As Integer
Dim nLanc As Integer
Dim nSeq As Integer
Dim nParc As Integer
Dim nCompl As Integer
Dim sVencto As String
Dim n As Integer
Dim sDataBase As String
Dim bAjuiza As Boolean, bDA As Boolean

nCod = Val(txtCod.text)
nAno = Val(Item.text)
nLanc = Left$(Item.ListSubItems(1).text, 2)
nSeq = Item.ListSubItems(2)
nParc = Item.ListSubItems(3)
nCompl = Item.ListSubItems(4)
sVencto = Item.ListSubItems(5)
sDataBase = Item.ListSubItems(7)
bAjuiza = IIf(Item.ListSubItems(8) = "S", True, False)
bDA = IIf(Item.ListSubItems(9) = "S", True, False)

If Item.Checked = True Then
    'adiciona na matriz
     ReDim Preserve aDebito(UBound(aDebito) + 1)
     aDebito(UBound(aDebito)).nCod = nCod
     aDebito(UBound(aDebito)).nAno = nAno
     aDebito(UBound(aDebito)).nLanc = nLanc
     aDebito(UBound(aDebito)).nSeq = nSeq
     aDebito(UBound(aDebito)).nParc = nParc
     aDebito(UBound(aDebito)).nCompl = nCompl
     aDebito(UBound(aDebito)).sVencto = sVencto
     aDebito(UBound(aDebito)).sDataBase = sDataBase
     aDebito(UBound(aDebito)).bAjuizado = bAjuiza
     aDebito(UBound(aDebito)).bDA = bDA
Else
    'Remove da matriz
    ReDim aDebitoTmp(UBound(aDebito) - 1)
    For n = 0 To UBound(aDebito)
        If aDebito(n).nCod = nCod And aDebito(n).nAno = nAno And aDebito(n).nLanc = nLanc And _
           aDebito(n).nSeq = nSeq And aDebito(n).nParc = nParc And aDebito(n).nCompl = nCompl Then
        Else
            If aDebito(n).nCod > 0 Then
                ReDim Preserve aDebitoTmp(UBound(aDebitoTmp) + 1)
                aDebitoTmp(UBound(aDebitoTmp)).nCod = aDebito(n).nCod
                aDebitoTmp(UBound(aDebitoTmp)).nAno = aDebito(n).nAno
                aDebitoTmp(UBound(aDebitoTmp)).nLanc = aDebito(n).nLanc
                aDebitoTmp(UBound(aDebitoTmp)).nSeq = aDebito(n).nSeq
                aDebitoTmp(UBound(aDebitoTmp)).nParc = aDebito(n).nParc
                aDebitoTmp(UBound(aDebitoTmp)).nCompl = aDebito(n).nCompl
                aDebitoTmp(UBound(aDebitoTmp)).sVencto = aDebito(n).sVencto
                aDebitoTmp(UBound(aDebitoTmp)).sDataBase = aDebito(n).sDataBase
                aDebitoTmp(UBound(aDebitoTmp)).bAjuizado = aDebito(n).bAjuizado
                aDebitoTmp(UBound(aDebitoTmp)).bDA = aDebito(n).bDA
            End If
        End If
    Next
    'Reverte a matriz
    ReDim aDebito(UBound(aDebitoTmp))
    For n = 0 To UBound(aDebitoTmp)
        aDebito(n).nCod = aDebitoTmp(n).nCod
        aDebito(n).nAno = aDebitoTmp(n).nAno
        aDebito(n).nLanc = aDebitoTmp(n).nLanc
        aDebito(n).nSeq = aDebitoTmp(n).nSeq
        aDebito(n).nParc = aDebitoTmp(n).nParc
        aDebito(n).nCompl = aDebitoTmp(n).nCompl
        aDebito(n).sVencto = aDebitoTmp(n).sVencto
        aDebito(n).sDataBase = aDebitoTmp(n).sDataBase
        aDebito(n).bAjuizado = aDebitoTmp(n).bAjuizado
        aDebito(n).bDA = aDebitoTmp(n).bDA
    Next
End If

End Sub

Private Sub mskDataProc_GotFocus()
mskDataProc.SelStart = 0
mskDataProc.SelLength = Len(mskDataProc.text)

End Sub


Private Sub mskVencto_GotFocus()
mskVencto.SelStart = 0
mskVencto.SelLength = Len(mskVencto.text)

End Sub

Private Sub txtCod_GotFocus()

txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.text)

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    lvDeb.SetFocus
    Exit Sub
End If

Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_LostFocus()
Dim nCodReduz As Long
If Val(txtCod.text) = 0 Then Exit Sub
If Val(txtCod.text) < 100000 Then
    sTipoCod = "I"
ElseIf Val(txtCod.text) >= 100000 And Val(txtCod.text) < 500000 Then
    sTipoCod = "M"
ElseIf Val(txtCod.text) >= 500000 Then
    sTipoCod = "C"
End If
txtCod.text = Format(txtCod.text, "0000000")
nCodReduz = Val(txtCod.text)
lblNome.Caption = ""
If sTipoCod = "I" Then
    Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
    Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
    Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & nCodReduz
ElseIf sTipoCod = "M" Then
    Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO Where CODIGOMOB =" & nCodReduz
ElseIf sTipoCod = "C" Then
    Sql = "SELECT NOMECIDADAO FROM CIDADAO Where CODCIDADAO =" & nCodReduz
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux.RowCount > 0 Then
         If sTipoCod = "I" Or sTipoCod = "C" Then
            lblNome.Caption = !NOMECIDADAO
         ElseIf sTipoCod = "M" Then
            lblNome.Caption = !RAZAOSOCIAL
         End If
    Else
       MsgBox "Código não Cadastrado.", vbExclamation, "Atenção"
       z = SendMessage(lvDeb.hwnd, LVM_DELETEALLITEMS, 0, 0)
       txtCod.SetFocus
       Exit Sub
    End If
    .Close
End With
CarregaDebito (nCodReduz)
End Sub

Private Sub AdicionaDebitos()
Dim nAno As Integer
Dim nCod As Long
Dim nLanc As Integer
Dim nSeq As Integer
Dim nParc As Integer
Dim nCompl As Integer
Dim sVencto As String, RdoAux2 As rdoResultset
Dim nTotal As Double, nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double, nSomaTotal As Double
Dim n As Integer, bS As Boolean, bN As Boolean, bJuros As Boolean, bMulta As Boolean, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double

nSomaAjuizado = 0

grdOrigem.Rows = 1

For n = 1 To UBound(aDebito)
    nTotal = 0
    If aDebito(n).nCod > 0 Then
        grdOrigem.AddItem Format(aDebito(n).nCod, "000000") & Chr(9) & aDebito(n).nAno & Chr(9) & Format(aDebito(n).nLanc, "00") & Chr(9) & _
        Format(aDebito(n).nSeq, "00") & Chr(9) & Format(aDebito(n).nParc, "00") & Chr(9) & Format(aDebito(n).nCompl, "00") & Chr(9) & aDebito(n).sVencto
        Sql = "SELECT CODTRIBUTO,TOTALLANCADO FROM VWCNSLANCAMENTO WHERE CODREDUZIDO=" & aDebito(n).nCod & " AND ANOEXERCICIO=" & aDebito(n).nAno & " AND "
        Sql = Sql & "CODLANCAMENTO =" & aDebito(n).nLanc & " And SEQLANCAMENTO = " & aDebito(n).nSeq & " And NumParcela = " & aDebito(n).nParc & " And CODCOMPLEMENTO = " & aDebito(n).nCompl & " AND CODTRIBUTO<>3"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
        With RdoAux
            nSomaJuros = 0: nSomaMulta = 0: nSomaCorrecao = 0
            Do Until .EOF
                Sql = "SELECT MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    bJuros = !JUROS
                    bMulta = !Multa
                   .Close
                End With
                nTotal = nTotal + !TOTALLANCADO
                
                If nTotal > 0 Then
                    nValorCorrecao = CalculaCorrecaoPlano(!TOTALLANCADO, CDate(aDebito(n).sVencto))
                    'nValorCorrecao = 0
                    If bJuros Then
                        nValorJuros = CalculaJurosPlano(!TOTALLANCADO + nValorCorrecao, CDate(aDebito(n).sVencto))
                    End If
                    If bMulta Then
                        nValorMulta = CalculaMultaPlano(!TOTALLANCADO + nValorCorrecao, CDate(aDebito(n).sVencto))
                    End If
                End If
                nSomaTotal = nSomaTotal + nTotal + nValorJuros + nValorMulta + nValorCorrecao
                nSomaJuros = nSomaJuros + nValorJuros
                nSomaMulta = nSomaMulta + nValorMulta
                nSomaCorrecao = nSomaCorrecao + nValorCorrecao
                nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0
               
               .MoveNext
            Loop
           .Close
        End With
        If nSomaTotal > 0 Then
            
            
            grdOrigem.TextMatrix(grdOrigem.Rows - 1, 7) = FormatNumber(nTotal, 2)
            grdOrigem.TextMatrix(grdOrigem.Rows - 1, 8) = FormatNumber(nSomaJuros, 2)
            grdOrigem.TextMatrix(grdOrigem.Rows - 1, 9) = FormatNumber(nSomaMulta, 2)
            grdOrigem.TextMatrix(grdOrigem.Rows - 1, 10) = FormatNumber(nSomaCorrecao, 2)
            grdOrigem.TextMatrix(grdOrigem.Rows - 1, 11) = FormatNumber(nTotal + nSomaJuros + nSomaMulta + nSomaCorrecao, 2)
            
            If aDebito(n).bAjuizado = True Then
                 nSomaAjuizado = nSomaAjuizado + CDbl(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 11))
            End If
        End If
    End If
Next

If chkHon.Value = vbChecked Then
   lblValorHon.Caption = FormatNumber(nSomaAjuizado * 0.1, 2)
Else
   lblValorHon.Caption = "0,00"
End If

fr1.Enabled = True
'fr2.Enabled = True
fr3.Enabled = True
fr4.Enabled = True
fr5.Enabled = True
frCod.Visible = False

End Sub

Private Sub AtualizaTotal()
Dim n As Integer, nSoma As Double
Dim nSomaJ As Double
Dim nSomaM As Double
Dim nSomaC As Double

lblNumParc.Caption = Format(grdOrigem.Rows - 1, "000")

nSoma = 0: nSomaJ = 0: nSomaM = 0: nSomaC = 0
For n = 1 To grdOrigem.Rows - 1
    If grdOrigem.TextMatrix(n, 7) = "" Then grdOrigem.TextMatrix(n, 7) = 0
    nSoma = nSoma + CDbl(grdOrigem.TextMatrix(n, 7))
    If grdOrigem.TextMatrix(n, 8) = "" Then grdOrigem.TextMatrix(n, 8) = 0
    If chkJuros.Value = 1 Then
       nSomaJ = nSomaJ + CDbl(grdOrigem.TextMatrix(n, 8))
    End If
    If grdOrigem.TextMatrix(n, 9) = "" Then grdOrigem.TextMatrix(n, 9) = 0
    If chkMulta.Value = 1 Then
        nSomaM = nSomaM + CDbl(grdOrigem.TextMatrix(n, 9))
    End If
    If grdOrigem.TextMatrix(n, 10) = "" Then grdOrigem.TextMatrix(n, 10) = 0
    nSomaC = nSomaC + CDbl(grdOrigem.TextMatrix(n, 10))
Next
lblValorParc.Caption = FormatNumber(nSoma + nSomaC, 2)
lblValorJuros.Caption = FormatNumber(nSomaJ, 2)
lblValorMulta.Caption = FormatNumber(nSomaM, 2)

End Sub

Private Sub AtualizaTotalGerado()
Dim n As Integer, nSomaP As Double, nSomaJ As Double, nSomaM As Double
lblNumParcGer.Caption = Format(grdDestino.Rows - 1, "000")

nSomaP = 0
nSomaJ = 0
nSomaM = 0
For n = 1 To grdDestino.Rows - 1
    nSomaP = nSomaP + CDbl(grdDestino.TextMatrix(n, 7))
    nSomaJ = nSomaJ + CDbl(grdDestino.TextMatrix(n, 9))
    nSomaM = nSomaM + CDbl(grdDestino.TextMatrix(n, 10))
Next
lblValorParcGer.Caption = FormatNumber(nSomaP, 2)
lblValorJurosGer.Caption = FormatNumber(nSomaJ, 2)
lblValorMultaGer.Caption = FormatNumber(nSomaM, 2)

End Sub

Private Sub AtualizaResp()
Dim n As Integer, x As Integer, Achou As Boolean

cmbResp.Clear
For n = 1 To UBound(aDebito)
    Achou = False
    For x = 0 To cmbResp.ListCount - 1
        cmbResp.ListIndex = x
        If Val(cmbResp.text) = aDebito(n).nCod Then
           Achou = True
           Exit For
        End If
    Next
    If Not Achou Then
       If aDebito(n).nCod > 0 Then
            cmbResp.AddItem Format(aDebito(n).nCod, "0000000")
       End If
    End If
Next
If cmbResp.ListCount > 0 Then cmbResp.ListIndex = 0

End Sub

Private Sub txtNumProc_GotFocus()
txtNumProc.SelStart = 0
txtNumProc.SelLength = Len(txtNumProc.text)

End Sub

Private Sub txtNumProc_KeyPress(KeyAscii As Integer)

If KeyAscii <> 47 Then
   Tweak txtNumProc, KeyAscii, IntegerPositive
End If
End Sub

Private Sub txtPercEntrada_Change()
txtValorEntrada.text = 0
End Sub

Private Sub txtPercEntrada_KeyPress(KeyAscii As Integer)
Tweak txtNumProc, KeyAscii, IntegerPositive

End Sub

Private Sub txtQtdeDil_Change()

If Val(txtQtdeDil) = 0 Then
   lblValorDil.Caption = "0,00"
   lblSomaDil.Caption = "0,00"
Else
   lblSomaDil.Caption = FormatNumber(CDbl(txtQtdeDil.text) * nValorDil, 2)
   lblValorDil.Caption = FormatNumber(CDbl(txtQtdeDil.text) * nValorDil, 2)
End If

End Sub

Private Sub txtQtdeDil_GotFocus()
txtQtdeDil.SelStart = 0
txtQtdeDil.SelLength = Len(txtQtdeDil.text)
End Sub

Private Sub txtQtdeDil_KeyPress(KeyAscii As Integer)
Tweak txtQtdeDil, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtdeParc_Change()
Dim nQtde As Integer
If Val(txtQtdeParc.text) > 48 Then txtQtdeParc.text = 48
nQtde = Val(txtQtdeParc.text)
If nQtde > 0 Then
'    If Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) = "ROSE" Then
'        lblTipoPlano.Caption = TipoDePlano(nQtde)
'        lblValorPlano.Caption = aPlanoDesconto(lblTipoPlano.Caption).nValor
'    Else
        lblTipoPlano.Caption = "N/A"
        lblValorPlano.Caption = aPlanoDesconto(0).nValor
'    End If
Else
    lblTipoPlano.Caption = ""
    lblValorPlano.Caption = ""
End If
RecalculaDebitosPlano
End Sub

Private Sub txtQtdeParc_GotFocus()
txtQtdeParc.SelStart = 0
txtQtdeParc.SelLength = Len(txtQtdeParc.text)

End Sub

Private Sub txtQtdeParc_KeyPress(KeyAscii As Integer)
Tweak txtQtdeParc, KeyAscii, IntegerPositive

End Sub

Private Sub txtValorEntrada_Change()
txtPercEntrada.text = 0
End Sub

Private Sub txtValorEntrada_GotFocus()
txtValorEntrada.SelStart = 0
txtValorEntrada.SelLength = Len(txtValorEntrada.text)

End Sub

Private Sub txtValorEntrada_KeyPress(KeyAscii As Integer)

Tweak txtValorEntrada, KeyAscii, DecimalPositive

End Sub

Private Sub txtValorEntrada_LostFocus()
If Trim$(txtValorEntrada.text) = "" Then txtValorEntrada.text = 0
End Sub

Private Sub GeraDebito()
Dim nQtde As Integer, nValorParcela As Double, nCodResp As Long
Dim nSomaTotal As Double, a As Integer, sVencimento As String
Dim nDia As Integer, nMes As Integer, nAno As Integer, nSomaDil As Double
Dim nSeq As Integer, nValorPrimeira As Double, nValorGerado As Double
Dim nSomaJ As Double, nValorParcelaJ As Double, nValorPrimeiraJ As Double, nValorGeradoJ As Double
Dim nSomaM As Double, nValorParcelaM As Double, nValorPrimeiraM As Double, nValorGeradoM As Double
Dim nValorEntrada As Double, nPercEntrada As Double, nSomaParcela As Double

grdDestino.Rows = 1
nSomaParcela = 0
'PARAMETROS DE CÁLCULO
nQtde = Val(txtQtdeParc.text)
nSomaTotal = CDbl(lblValorParc.Caption) + CDbl(lblValorHon.Caption)
nSomaDil = CDbl(lblValorDil.Caption)
nSomaJ = CDbl(lblValorJuros.Caption)
nSomaM = CDbl(lblValorMulta.Caption)
nValorParcela = Round(nSomaTotal / nQtde, 2)
nValorParcelaJ = Round(nSomaJ / nQtde, 2)
nValorParcelaM = Round(nSomaM / nQtde, 2)
nCodResp = Val(cmbResp.text)
sVencimento = mskVencto.text

'OBTER % DE ENTRADA
nValorEntrada = CDbl(txtValorEntrada.text)
nPercEntrada = CDbl(txtPercEntrada.text)
If nValorEntrada > 0 Then
   nPercEntrada = (nValorEntrada * 100 / (nSomaTotal + nSomaJ + nSomaM)) / 100
Else
   If nPercEntrada > 0 Then
      nPercEntrada = nPercEntrada / 100
   End If
End If

'CALCULO PROPORCIONAL DE ENTRADA
If nPercEntrada > 0 Then
    nValorPrimeira = nSomaTotal * nPercEntrada
    nValorParcela = (nSomaTotal - nValorPrimeira) / (nQtde - 1)
    nValorPrimeiraJ = nSomaJ * nPercEntrada
    nValorParcelaJ = (nSomaJ - nValorPrimeiraJ) / (nQtde - 1)
    nValorPrimeiraM = nSomaM * nPercEntrada
    nValorParcelaM = (nSomaM - nValorPrimeiraM) / (nQtde - 1)
Else
    nValorPrimeira = nValorParcela
    nValorPrimeiraJ = nValorParcelaJ
    nValorPrimeiraM = nValorParcelaM
End If
nValorPrimeira = nValorPrimeira + nSomaDil

'BUSCA ULTIMA SEQUENCIA
Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodResp & " AND CODLANCAMENTO=20"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux!CONTADOR > 0 Then
    Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodResp & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO<100"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nSeq = RdoAux!SEQMAXIMA + 1
Else
    nSeq = 0
End If

For a = 1 To nQtde
    If a > 1 Then
       nDia = Val(Left$(mskVencto.text, 2))
       nMes = Val(Mid$(sVencimento, 4, 2)) + 1
       
       nAno = Val(Right$(sVencimento, 4))
       If nMes = 13 Then
          nMes = 1: nAno = nAno + 1
       End If
               
       sVencimento = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
       
       If Not IsDate(sVencimento) Then
           nDia = nDia - 3
           sVencimento = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
       End If
       
       nValorGerado = nValorParcela
       nValorGeradoJ = nValorParcelaJ
       nValorGeradoM = nValorParcelaM
    Else
       nAno = Val(Right$(sVencimento, 4))
       nValorGerado = nValorPrimeira
       nValorGeradoJ = nValorPrimeiraJ
       nValorGeradoM = nValorPrimeiraM
    End If
    
    grdDestino.AddItem Format(nCodResp, "0000000") & Chr(9) & nAno & Chr(9) & "20" & Chr(9) & _
    Format(nSeq, "00") & Chr(9) & Format(a, "00") & Chr(9) & "00" & Chr(9) & sVencimento & _
    Chr(9) & FormatNumber(nValorGerado, 2) & Chr(9) & "" & Chr(9) & FormatNumber(nValorGeradoJ, 2) & _
    Chr(9) & FormatNumber(nValorGeradoM, 2)
    nSomaParcela = FormatNumber(nValorGerado + nValorGeradoJ + nValorGeradoM + 1.13, 2)
    
Next

lblValorParcela.Caption = nSomaParcela

AtualizaTotalGerado

End Sub

Private Sub GravaCarneTmp()
On Error GoTo Erro

Dim x As Integer
Dim RdoAux2 As rdoResultset
Dim sNumInsc As String
Dim sCodReduz As String
Dim sNomeResp As String
Dim sTipoImposto As String
Dim sEndImovel As String
Dim nNumImovel As Integer
Dim sComplImovel As String
Dim sBairroImovel As String
Dim nCodLogr As Long
Dim sEndEntrega As String
Dim nNumEntrega As Integer
Dim sBairroEntrega As String
Dim sComplEntrega As String
Dim sCepEntrega As String
Dim sCidadeEntrega As String
Dim sUFEntrega As String
Dim sDescImposto As String
Dim nAno As Integer
Dim sNumProc As String
Dim dDataProc As Date
Dim nNumDoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim dDataVencto As Date
Dim sValorParc As String
Dim NumBarra1 As String
Dim StrBarra1 As String
Dim NumBarra2 As String
Dim NumBarra2a As String
Dim NumBarra2b As String
Dim NumBarra2c As String
Dim NumBarra2d As String
Dim StrBarra2 As String
Dim nLastCod As Long
Dim nValorTaxa As Double
Dim nValorHonDil As Double
Dim nCodLanc As Integer, nCodTrib As Integer

If MsgBox("Gravar o Reparcelamento ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

'BUSCA O VALOR DA TAXA DE EXPEDIENTE
Sql = "SELECT VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorTaxa = FormatNumber(!VALORPARCELA, 2)
   .Close
End With

Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & txtNumProc.text & " '"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este processo foi reparcelado, para continuar a usar este processo no reparcelamento voce deve cancela-lo primeiro.", vbExclamation, "Atenção"
    Exit Sub
End If

'GRAVA O PROCESSO
Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CODIGORESP,FUNCIONARIO,PLANO) VALUES('"
Sql = Sql & txtNumProc.text & "','" & Format(mskDataProc.text, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',"
Sql = Sql & txtQtdeParc.text & "," & Virg2Ponto(txtValorEntrada.text) & "," & Virg2Ponto(txtPercEntrada.text) & ","
Sql = Sql & IIf(chkMulta.Value = vbChecked, 1, 0) & "," & IIf(chkJuros.Value = vbChecked, 1, 0) & ","
Sql = Sql & Val(cmbResp.text) & ",'" & NomeDeLogin & "',"
'Sql = Sql & Val(cmbResp.text) & ",'" & Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) & "',"
Sql = Sql & Val(lblTipoPlano.Caption) & ")"
cn.Execute Sql, rdExecDirect

'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!MAXIMO) Then
   nLastCod = 0
Else
   nLastCod = RdoAux!MAXIMO + 10
End If
RdoAux.Close

nValorHonDil = (CDbl(lblValorHon.Caption) + CDbl(lblValorDil.Caption)) / Val(txtQtdeParc.text)

'GRAVA AS PARCELAS DE DESTINO
With grdDestino
    For x = 1 To grdDestino.Rows - 1
          'GRAVA DESTINOREPARC
          Sql = "INSERT DESTINOREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO) VALUES('" & txtNumProc.text & "'," & .TextMatrix(x, 0) & ","
          Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
          Sql = Sql & .TextMatrix(x, 5) & ")"
          cn.Execute Sql, rdExecDirect
          'GRAVA DEBITOPARCELA    // (STATUS 3 - NAO PAGO)
          Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
          Sql = Sql & "NUMPROCESSO) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & ","
          Sql = Sql & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
          Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',"
          Sql = Sql & 1 & ",'" & txtNumProc.text & "')"
          cn.Execute Sql, rdExecDirect
          
          nCodLanc = grdOrigem.TextMatrix(1, 2)
          Select Case nCodLanc
            Case 1, 29
                nCodTrib = 200 'REPARC IPTU
            Case 2, 3, 5
                nCodTrib = 201  'REPARC ISS
            Case 13
                nCodTrib = 202 'REPARC VIGILANCIA
            Case Else
                nCodTrib = 203 'REPARC DIVERSO
          End Select
          
          'GRAVA DEBITOTRIBUTO   // (TRIBUTOS REPARCELADOS)
          Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
          Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
          Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nCodTrib & "," & Virg2Ponto(CStr(CDbl(.TextMatrix(x, 7)) - CDbl(nValorHonDil))) & ")"
          cn.Execute Sql, rdExecDirect
                                
          'GRAVA DEBITOTRIBUTO   // (TRIBUTO 3 - TX.EXP.DOC)
          Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
          Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
          Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 3 & "," & Virg2Ponto(CStr(nValorTaxa)) & ")"
          cn.Execute Sql, rdExecDirect
          
          'GRAVA DEBITOTRIBUTO   // (TRIBUTO 113 - JUROS)
          If CDbl(.TextMatrix(x, 9)) > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
            Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 113 & "," & Virg2Ponto(.TextMatrix(x, 9)) & ")"
            cn.Execute Sql, rdExecDirect
          End If
          'GRAVA DEBITOTRIBUTO   // (TRIBUTO 112 - MULTA)
          If CDbl(.TextMatrix(x, 10)) > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
            Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 112 & "," & Virg2Ponto(.TextMatrix(x, 10)) & ")"
            cn.Execute Sql, rdExecDirect
          End If
          If CDbl(lblValorHon.Caption) > 0 Then
                'GRAVA DEBITOTRIBUTO   // (TRIBUTO 90 - HONORÁRIOS)
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
                Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 90 & "," & Virg2Ponto(CDbl(lblValorHon.Caption) / Val(txtQtdeParc.text)) & ")"
                cn.Execute Sql, rdExecDirect
          End If
          If CDbl(lblValorDil.Caption) > 0 And x = 1 Then
                'GRAVA DEBITOTRIBUTO   // (TRIBUTO 91 - DILIGÊNCIAS)
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
                Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 91 & "," & Virg2Ponto(CDbl(lblValorDil.Caption)) & ")"
                cn.Execute Sql, rdExecDirect
          End If
          'GRAVA NUMDOCUMENTO
          Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
          Sql = Sql & nLastCod + x & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTaxa)) & ")"
          cn.Execute Sql, rdExecDirect
          grdDestino.TextMatrix(x, 8) = nLastCod + x
          'GRAVA PARCELADOCUMENTO
          Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(cmbResp.text) & ","
          Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
          Sql = Sql & .TextMatrix(x, 5) & "," & nLastCod + x & ")"
          cn.Execute Sql, rdExecDirect
    Next
End With

'GRAVA AS PARCELAS DE ORIGEM
With grdOrigem
    For x = 1 To grdOrigem.Rows - 1
          'GRAVA ORIGEMREPARC
          Sql = "INSERT ORIGEMREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO) VALUES('" & txtNumProc.text & "'," & .TextMatrix(x, 0) & ","
          Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
          Sql = Sql & .TextMatrix(x, 5) & ")"
          cn.Execute Sql, rdExecDirect
         'ATUALIZA O STATUS DE ORIGEM   // (4 - REPARCELADO)
          Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=4 WHERE CODREDUZIDO=" & .TextMatrix(x, 0)
          Sql = Sql & " AND ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND CODLANCAMENTO=" & .TextMatrix(x, 2)
          Sql = Sql & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4)
          Sql = Sql & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
          cn.Execute Sql, rdExecDirect
    Next
End With

'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

'DADOS CABEÇALHO
sNumProc = txtNumProc.text
dDataProc = CDate(mskDataProc.text)
sDescImposto = "REPARCELAMENTO"
NumBarra1 = Format(ExtraiNumero(sNumProc), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)


'****************************************************************************************
Select Case Val(txtCod.text)
    Case 1 To 99999
        'DADOS DO IMOVEL
        Sql = "SELECT * FROM vwCnsImovel WHERE CODREDUZIDO=" & Val(cmbResp.text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
               If .RowCount > 0 Then
                     sNumInsc = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
                     sCodReduz = cmbResp.text & "-" & RetornaDVCodReduzido(Val(cmbResp.text))
                     Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
                     Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
                     Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & Val(cmbResp.text)
                     Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset)
                     sNomeResp = RdoAux3!NOMECIDADAO
                     RdoAux3.Close
                     sTipoImposto = "REPARCEL."
                     sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                     nNumImovel = !Li_Num
                     sComplImovel = IIf(IsNull(!Li_Compl) Or !Li_Compl = "", " ", !Li_Compl)
                     If !CODBAIRRO <> 999 Then
                         sBairroImovel = !DescBairro
                     End If
                     nCodLogr = !CodLogr
                     sQuadra = !Li_Quadras
                     sLote = !Li_Lotes
                    .Close
                    'ENDERECO DE ENTREGA
                    Sql = "SELECT ENDENTREGA.*, BAIRRO.DESCBAIRRO, Cidade.DESCCIDADE FROM ENDENTREGA LEFT OUTER JOIN "
                    Sql = Sql & "CIDADE ON ENDENTREGA.EE_UF = CIDADE.SIGLAUF AND ENDENTREGA.EE_CIDADE = CIDADE.CODCIDADE "
                    Sql = Sql & "LEFT OUTER JOIN  BAIRRO ON ENDENTREGA.EE_UF = BAIRRO.SIGLAUF AND ENDENTREGA.EE_CIDADE = BAIRRO.CODCIDADE "
                    Sql = Sql & "AND  ENDENTREGA.EE_BAIRRO = BAIRRO.CODBAIRRO "
                    Sql = Sql & "WHERE CODREDUZIDO=" & Val(cmbResp.text)
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                           If .RowCount > 0 Then
                               sEndEntrega = !Ee_NomeLog
                               nNumEntrega = !Ee_NumImovel
                               sBairroEntrega = IIf(!DescBairro = 999, " ", !DescBairro)
                               sComplEntrega = IIf(IsNull(!Ee_Complemento) Or !Ee_Complemento = "", " ", !Ee_Complemento)
                               sCepEntrega = !Ee_Cep
                               sCidadeEntrega = !desccidade
                               sUFEntrega = !Ee_Uf
                           Else
                               sEndEntrega = sEndImovel
                               nNumEntrega = nNumImovel
                               sBairroEntrega = sBairroImovel
                               sComplEntrega = sComplImovel
                               sCepEntrega = RetornaCEP(nCodLogr, nNumEntrega)
                               sCidadeEntrega = "Jaboticabal"
                               sUFEntrega = "SP"
                           End If
                          .Close
                    End With
               Else
                    MsgBox "Código não cadastrado.", vbCritical, "Atenção"
               End If
        End With
     Case 100000 To 500000
        'DADOS DA EMPRESA
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & Val(cmbResp.text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               sNumInsc = SubNull(!INSCESTADUAL)
               sCodReduz = cmbResp.text
               sNomeResp = !RAZAOSOCIAL
               sTipoImposto = "REPARCEL."
               sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
               nNumImovel = Val(SubNull(!Numero))
               sComplImovel = SubNull(!COMPLEMENTO)
               If !CODBAIRRO <> 999 Then
                    Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO=" & !CODBAIRRO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                             sBairroImovel = !DescBairro
                        Else
                             sBairroImovel = ""
                        End If
                       .Close
                    End With
               End If
               nCodLogr = 0
               sQuadra = ""
               sLote = ""
        
               Sql = "SELECT NOMELOGRADOURO,NUMIMOVEL,COMPLEMENTO,UF,CIDADE.DESCCIDADE AS DESCCIDADE1,"
               Sql = Sql & "BAIRRO.DESCBAIRRO AS DESCBAIRRO1,CEP,MOBILIARIOENDENTREGA.DESCBAIRRO,"
               Sql = Sql & "MOBILIARIOENDENTREGA.DESCCIDADE FROM CIDADE INNER JOIN BAIRRO ON "
               Sql = Sql & "CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE RIGHT OUTER Join "
               Sql = Sql & "MOBILIARIOENDENTREGA ON BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND "
               Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                    If .RowCount > 0 Then
                        sEndEntrega = SubNull(!NomeLogradouro)
                        nNumEntrega = SubNull(!NUMIMOVEL)
                        sBairroEntrega = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        sComplEntrega = SubNull(!COMPLEMENTO)
                        sCepEntrega = SubNull(!cep)
                        sCidadeEntrega = IIf(IsNull(!desccidade), SubNull(!DESCCIDADE1), SubNull(!desccidade))
                        sUFEntrega = SubNull(!UF)
                    Else
                        sEndEntrega = sEndImovel
                        nNumEntrega = nNumImovel
                        sBairroEntrega = sBairroImovel
                        sComplEntrega = sComplImovel
                        sCepEntrega = ""
                        sCidadeEntrega = "JABOTICABAL"
                        sUFEntrega = "SP"
                    End If
                   .Close
               End With
            Else
            End If
           .Close
        End With
     Case 500000 To 800000
        Sql = "SELECT cidadao.codcidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, "
        Sql = Sql & "cidadao.siglauf, cidade.desccidade, bairro.descbairro, cidadao.nomelogradouro AS nomerua, cidadao.nomebairro, cidadao.nomecidade,"
        Sql = Sql & "cidadao.codlogradouro , vwLOGRADOURO.AbrevTipoLog, vwLOGRADOURO.AbrevTitLog, vwLOGRADOURO.NomeLogradouro "
        Sql = Sql & "FROM  vwLOGRADOURO INNER JOIN  cidadao ON vwLOGRADOURO.CODLOGRADOURO = cidadao.codlogradouro LEFT OUTER JOIN "
        Sql = Sql & "cidade INNER JOIN  bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade And cidadao.codbairro = bairro.codbairro "
        Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & Val(txtCod.text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sTipoImposto = "REPARCEL."
                If Not IsNull(!NomeLogradouro) Then
                    sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                Else
                    sEndImovel = SubNull(!nomerua)
                End If
                nNumImovel = SubNull(!NUMIMOVEL)
                sComplImovel = SubNull(!COMPLEMENTO)
                nCodBairro = Val(SubNull(!CODBAIRRO))
                sCidadeEntrega = SubNull(!desccidade)
                sUFEntrega = SubNull(!SiglaUF)
                If Not IsNull(!CPF) Then
                   sCPF = !CPF
                ElseIf Not IsNull(!CNPJ) Then
                   sCPF = !CNPJ
                ElseIf Not IsNull(!RG) Then
                   sCPF = !RG
                Else
                   sCPF = ""
                End If
             Else
                sCPF = ""
             End If
             If nCodBairro = 0 Then
                sBairroImovel = SubNull(!NOMEBAIRRO)
                sCidadeEntrega = SubNull(!NomeCidade)
             End If
        End With
End Select

'****************************************************************************************


    
'GRAVA TEMPORARIO
For x = 1 To txtQtdeParc.text
    nAno = grdDestino.TextMatrix(x, 1)
    nNumDoc = grdDestino.TextMatrix(x, 8)
    nNumParc = grdDestino.TextMatrix(x, 4)
    dDataVencto = grdDestino.TextMatrix(x, 6)
    sValorParc = CDbl(grdDestino.TextMatrix(x, 7)) + CDbl(grdDestino.TextMatrix(x, 9)) + CDbl(grdDestino.TextMatrix(x, 10)) + CDbl(nValorTaxa)
    NumBarra2 = Gera2of5Cod(sValorParc, grdDestino.TextMatrix(x, 6), nNumDoc, nNumParc, grdDestino.TextMatrix(x, 2), grdDestino.TextMatrix(x, 3), grdDestino.TextMatrix(x, 5))
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)
    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))

    Sql = "INSERT CARNETMP(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,ENDIMOVEL,NUMIMOVEL,COMPLIMOVEL,"
    Sql = Sql & "BAIRROIMOVEL,ENDENTREGA,NUMENTREGA,COMPLENTREGA,BAIRROENTREGA,CEPENTREGA,CIDADEENTREGA,UFENTREGA,"
    Sql = Sql & "DESCIMPOSTO,EXERCICIO,NUMPROCESSO,DATAPROCESSO,NUMDOCUMENTO,DV,QUADRA,LOTE,DATAVENCTO,NUMPARCELA,"
    Sql = Sql & "NUMTOTPARCELA,VALORPARCELA,STRBARRA1,STRBARRA2,NUMBARRA1,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
    Sql = Sql & "DADOSLANCAMENTO,TAXAEXP,SAIR) VALUES('" & NomeDoUsuario & "'," & x & ",'" & sNumInsc & "','" & sCodReduz & "','"
    Sql = Sql & sTipoImposto & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left$(sEndImovel, 40) & "'," & nNumImovel & ",'" & Left$(sComplImovel, 30) & "','"
    Sql = Sql & Left$(sBairroImovel, 25) & "','" & Left$(sEndEntrega, 40) & "'," & nNumEntrega & ",'" & Left$(sComplEntrega, 30) & "','" & Left$(sBairroEntrega, 25) & "','"
    Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Left$(sDescImposto, 30) & "'," & nAno & ",'" & Left$(sNumProc, 25) & "','"
    Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "'," & nNumDoc & "," & RetornaDVNumDoc(nNumDoc) & ",'" & sQuadra & "','"
    Sql = Sql & sLote & "','" & Format(dDataVencto, "mm/dd/yyyy") & "'," & IIf(nNumParc = 0, 1, nNumParc) & "," & Val(txtQtdeParc.text) & ","
    Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
    Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & "REPARCELAMENTO" & "'," & "0" & "," & "0" & ")"
    cn.Execute Sql, rdExecDirect

Next

frmConfissaoDivida.txtNumProc.text = txtNumProc.text
frmConfissaoDivida.txtNumProc.Locked = True
frmConfissaoDivida.CarregaProcesso
frmConfissaoDivida.show


Exit Sub

Erro:
For z = 0 To rdoErrors.Count - 1
Next
Resume Next
End Sub

Private Function CalculaJurosPlano(nValorDebito As Double, dDataVencto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String, nValorJuros As Double
Dim nValorPlano As Double
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= CDate(mskVencto.text) Then
    CalculaJurosPlano = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(CDate(mskVencto.text)) And Year(dDataVencto) = Year(CDate(mskVencto.text)) Then
    CalculaJurosPlano = 0
    Exit Function
End If

If Not dcJuros.Exists(Year(CDate(mskVencto.text))) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJurosPlano = 0
   Exit Function
End If

'MONTA O NOVO VENCIMENTO A PARTIR DO DIA 1 DO MES SUBSEQUENTE
nDia = Day(dDataVencto)
nMes = Month(dDataVencto)
nAno = Year(dDataVencto)
nDia = 1
If nMes = 12 Then
    nMes = 1
    nAno = nAno + 1
Else
    nMes = nMes + 1
End If

sDataVencto = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
dDataVencto = Format(sDataVencto, "dd/mm/yyyy")
nNumMes = Int(DateDiff("d", dDataVencto, CDate(mskVencto.text)) / 30) + 1



'If CDate(dDataVencto) >= CDate(mskVencto.text) Then
'    CalculaJurosPlano = 0
'    Exit Function
'End If
If lblValorPlano.Caption = "" Then lblValorPlano.Caption = "0"
nValorPlano = CDbl(lblValorPlano.Caption) / 100

'nNumMes = Int((DateDiff("d", dDataVencto, CDate(mskVencto.text))) / 30)
'Sql = "SELECT PERCJUROS FROM JUROS WHERE ANOJUROS=" & Year(CDate(mskVencto.text))
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If .RowCount = 0 Then
'        MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
'        CalculaJurosPlano = 0
'        Exit Function
'    Else
'        nValorPerc = !PERCJUROS
'    End If
'   .Close
'End With

nValorPerc = nValorPerc / 100
nValorJuros = nValorDebito * nValorPerc * nNumMes
nValorJuros = nValorJuros - (nValorJuros * nValorPlano)

If nValorJuros > 0 Then
   CalculaJurosPlano = FormatNumber(nValorJuros, 3)
End If

End Function

Private Function CalculaMultaPlano(nValorDebito As Double, dDataVencto As Date) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String, nValorMulta As Double
Dim nValorPlano As Double

If CDate(dDataVencto) >= CDate(mskVencto.text) Then
    CalculaMultaPlano = 0
    Exit Function
End If
On Error Resume Next

nValorPlano = CDbl(lblValorPlano.Caption) / 100

nNumDia = Abs(DateDiff("d", CDate(mskVencto.text), dDataVencto))

If nNumDia = 0 Then
   CalculaMultaPlano = 0
   Exit Function
End If

Sql = "SELECT MINDIA,MAXDIA,PERCDIA FROM MULTA WHERE ANOMULTA=" & Year(dDataVencto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         If nNumDia >= !MINDIA And nNumDia <= !MAXDIA Then
             nValorPerc = !PERCDIA
             Exit Do
         ElseIf nNumDia >= !MINDIA And !MAXDIA = 0 Then
             nValorPerc = !PERCDIA
             Exit Do
         End If
        .MoveNext
    Loop
End With

nValorPerc = nValorPerc / 100
nValorMulta = nValorDebito * nValorPerc
nValorMulta = nValorMulta - (nValorMulta * nValorPlano)

If nValorMulta > 0 Then
   CalculaMultaPlano = FormatNumber(nValorMulta, 3)
End If

End Function

Private Sub RecalculaDebitosPlano()
Dim sVencto As String
Dim nTotal As Double, nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double, nSomaTotal As Double
Dim n As Integer
Exit Sub
For n = 1 To grdOrigem.Rows - 1
    nTotal = grdOrigem.TextMatrix(n, 7)
    sVencto = grdOrigem.TextMatrix(n, 6)
    nValorCorrecao = CalculaCorrecao(nTotal, CDate(sVencto))
    nValorJuros = CalculaJurosPlano(nTotal + nValorCorrecao, CDate(sVencto))
    nValorMulta = CalculaMultaPlano(nTotal + nValorCorrecao, CDate(sVencto))
            
    grdOrigem.TextMatrix(n, 8) = FormatNumber(nValorJuros, 2)
    grdOrigem.TextMatrix(n, 9) = FormatNumber(nValorMulta, 2)
    grdOrigem.TextMatrix(n, 10) = FormatNumber(nValorCorrecao, 2)
    grdOrigem.TextMatrix(n, 11) = FormatNumber(nSomaTotal, 2)
Next

AtualizaTotal
End Sub


Public Function CalculaCorrecaoPlano(nValorDebito As Double, dDataBase As Date) As Double

Dim RdoAux As rdoResultset, Sql As String
Dim UfirAtual As Double
Dim UfirBase As Double

If Year(dDataBase) > Year(mskVencto.text) Then
    CalculaCorrecaoPlano = 0
    Exit Function
End If

UfirAtual = RetornaUFIR(Year(mskVencto.text))
If UfirAtual = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano atual.", vbCritical, "Alerta !!!"
    CalculaCorrecaoPlano = 0
    Exit Function
End If

UfirBase = RetornaUFIR(Year(dDataBase))
If UfirBase = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano base.", vbCritical, "Alerta !!!"
    CalculaCorrecaoPlano = 0
    Exit Function
End If

CalculaCorrecaoPlano = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecaoPlano > 0 Then
   CalculaCorrecaoPlano = FormatNumber(CalculaCorrecaoPlano, 2)
End If
End Function

