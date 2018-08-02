VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmParcelamento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parcelamento de Divida Fiscal"
   ClientHeight    =   5595
   ClientLeft      =   4260
   ClientTop       =   2370
   ClientWidth     =   9480
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   9480
   Begin VB.CheckBox chkMulta 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Multa"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6510
      TabIndex        =   4
      Top             =   150
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkJuros 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Juros"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7380
      TabIndex        =   5
      Top             =   150
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.CheckBox chkHon 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Honorário"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8340
      TabIndex        =   8
      Top             =   480
      Value           =   1  'Checked
      Width           =   990
   End
   Begin VB.TextBox txtQtdeDil 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   7395
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "0"
      Top             =   840
      Width           =   465
   End
   Begin VB.TextBox txtValorEntrada 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7380
      TabIndex        =   7
      Text            =   "0"
      Top             =   450
      Width           =   855
   End
   Begin VB.CheckBox chkCorrecao 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Correção"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8340
      TabIndex        =   6
      Top             =   150
      Value           =   1  'Checked
      Width           =   990
   End
   Begin VB.Frame fr2 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   60
      TabIndex        =   43
      Top             =   60
      Width           =   6285
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1590
         TabIndex        =   1
         Top             =   660
         Width           =   1275
      End
      Begin VB.TextBox txtQtdeParc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1590
         TabIndex        =   2
         Top             =   990
         Width           =   1275
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1590
         MaxLength       =   6
         TabIndex        =   0
         Top             =   330
         Width           =   1275
      End
      Begin esMaskEdit.esMaskedEdit mskDataProc 
         Height          =   285
         Left            =   4695
         TabIndex        =   44
         Top             =   690
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   15658734
         ForeColor       =   12582912
         MouseIcon       =   "frmParcelamento.frx":0000
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
      Begin esMaskEdit.esMaskedEdit mskVencto 
         Height          =   285
         Left            =   4680
         TabIndex        =   3
         Top             =   1005
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         MouseIcon       =   "frmParcelamento.frx":001C
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
         Left            =   0
         TabIndex        =   53
         Top             =   30
         Width           =   2910
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo.....:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Processo.......:"
         Height          =   225
         Index           =   1
         Left            =   2970
         TabIndex        =   51
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Parcelamento:"
         Height          =   225
         Index           =   2
         Left            =   2970
         TabIndex        =   50
         Top             =   60
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de Parcelas...:"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do 1º Vencto......:"
         Height          =   225
         Index           =   4
         Left            =   2970
         TabIndex        =   48
         Top             =   1050
         Width           =   1665
      End
      Begin VB.Label lblDataParc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4680
         TabIndex        =   47
         Top             =   60
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido...:"
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   46
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblNome 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2970
         TabIndex        =   45
         Top             =   330
         Width           =   3195
      End
   End
   Begin VB.Frame fr4 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   60
      TabIndex        =   22
      Top             =   5040
      Width           =   9345
      Begin prjChameleon.chameleonButton cmdCancel 
         Height          =   345
         Left            =   1500
         TabIndex        =   12
         ToolTipText     =   "Cancelar Operação"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Cancelar"
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
         MICON           =   "frmParcelamento.frx":0038
         PICN            =   "frmParcelamento.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDel 
         Height          =   345
         Left            =   5850
         TabIndex        =   15
         ToolTipText     =   "Desmarca todos os lançamentos"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Remover"
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
         MICON           =   "frmParcelamento.frx":02DD
         PICN            =   "frmParcelamento.frx":02F9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGeraDebito 
         Height          =   345
         Left            =   60
         TabIndex        =   11
         ToolTipText     =   "Gerar os Débitos na Tela"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Gerar"
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
         MICON           =   "frmParcelamento.frx":0453
         PICN            =   "frmParcelamento.frx":046F
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
         Left            =   7980
         TabIndex        =   16
         ToolTipText     =   "Sair da Tela"
         Top             =   60
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   609
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
         MICON           =   "frmParcelamento.frx":06B3
         PICN            =   "frmParcelamento.frx":06CF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdSimulado 
         Height          =   345
         Left            =   2955
         TabIndex        =   13
         ToolTipText     =   "Simulado de Parcelamento"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Simulado"
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
         MICON           =   "frmParcelamento.frx":073D
         PICN            =   "frmParcelamento.frx":0759
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAdd 
         Height          =   345
         Left            =   4395
         TabIndex        =   14
         ToolTipText     =   "Seleciona todos os lançamentos"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Marcar"
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
         MICON           =   "frmParcelamento.frx":0B21
         PICN            =   "frmParcelamento.frx":0B3D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdGrava 
         Height          =   345
         Left            =   1500
         TabIndex        =   23
         ToolTipText     =   "Gravar o parcelamento e imprimir os documentos"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "G&ravar"
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
         MICON           =   "frmParcelamento.frx":0C97
         PICN            =   "frmParcelamento.frx":0CB3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdVoltar 
         Height          =   345
         Left            =   60
         TabIndex        =   24
         ToolTipText     =   "Gerar os Débitos na Tela"
         Top             =   60
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmParcelamento.frx":0E5B
         PICN            =   "frmParcelamento.frx":0E77
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
   Begin MSComctlLib.ListView lvDestino 
      Height          =   2355
      Left            =   30
      TabIndex        =   10
      Top             =   1710
      Visible         =   0   'False
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   1340
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Sq"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Pc"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Vencto."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Principal"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Juros"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Multa"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Correção"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Honor."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Dilig."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Total"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Documento"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvOrigem 
      Height          =   2355
      Left            =   30
      TabIndex        =   17
      Top             =   1710
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   4154
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   1340
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lc"
         Object.Width           =   2472
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Sq"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Pc"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Cp"
         Object.Width           =   707
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Vencto."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Aj"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Principal"
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Juros"
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Multa"
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Correção"
         Object.Width           =   1766
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Total"
         Object.Width           =   1766
      EndProperty
   End
   Begin VB.Label lblSomaDil 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   8430
      TabIndex        =   61
      Top             =   870
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Diligências.:"
      Height          =   195
      Index           =   4
      Left            =   6480
      TabIndex        =   60
      Top             =   870
      Width           =   885
   End
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
      Left            =   8910
      TabIndex        =   59
      Top             =   1170
      Width           =   255
   End
   Begin VB.Label lblValorPlano 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   8550
      TabIndex        =   58
      Top             =   1170
      Width           =   315
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto de :"
      Height          =   225
      Left            =   7470
      TabIndex        =   57
      Top             =   1170
      Width           =   1005
   End
   Begin VB.Label lblTipoPlano 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
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
      Left            =   6990
      TabIndex        =   56
      Top             =   1170
      Width           =   375
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Plano:"
      Height          =   225
      Left            =   6480
      TabIndex        =   55
      Top             =   1170
      Width           =   435
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada.....:"
      Height          =   225
      Index           =   6
      Left            =   6480
      TabIndex        =   54
      Top             =   510
      Width           =   825
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
      TabIndex        =   42
      Top             =   4170
      Width           =   2115
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor da Correção......:"
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
      Index           =   0
      Left            =   6195
      TabIndex        =   41
      Top             =   4410
      Width           =   2070
   End
   Begin VB.Label lblValorCorrecao 
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
      Left            =   8205
      TabIndex        =   40
      Top             =   4425
      Width           =   1020
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
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   4425
      Width           =   2070
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
      Left            =   5100
      TabIndex        =   38
      Top             =   4410
      Width           =   960
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
      Left            =   3360
      TabIndex        =   37
      Top             =   4425
      Width           =   1710
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
      Left            =   3360
      TabIndex        =   36
      Top             =   4170
      Width           =   1710
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
      Left            =   5100
      TabIndex        =   35
      Top             =   4170
      Width           =   960
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
      Left            =   6195
      TabIndex        =   34
      Top             =   4170
      Width           =   1980
   End
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
      Left            =   8205
      TabIndex        =   33
      Top             =   4170
      Width           =   1020
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Principal...........:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   4680
      Width           =   2010
   End
   Begin VB.Label lblValorPrincipal 
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
      Left            =   2160
      TabIndex        =   31
      Top             =   4665
      Width           =   960
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
      Left            =   2160
      TabIndex        =   30
      Top             =   4410
      Width           =   960
   End
   Begin VB.Label lblNumParc 
      Alignment       =   1  'Right Justify
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
      Left            =   2160
      TabIndex        =   29
      Top             =   4170
      Width           =   960
   End
   Begin VB.Label lblValorTotal 
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
      Left            =   5100
      TabIndex        =   28
      Top             =   4650
      Width           =   960
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Totalizado...:"
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
      Index           =   2
      Left            =   3360
      TabIndex        =   27
      Top             =   4665
      Width           =   1710
   End
   Begin VB.Label lblValorParcela 
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
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   8205
      TabIndex        =   26
      Top             =   4665
      Width           =   1020
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor da Parcela........:"
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
      Height          =   225
      Index           =   1
      Left            =   6210
      TabIndex        =   25
      Top             =   4650
      Width           =   2070
   End
   Begin VB.Label lblAnoProc 
      Height          =   315
      Left            =   2130
      TabIndex        =   21
      Top             =   6000
      Width           =   1635
   End
   Begin VB.Label lblNumProc 
      Height          =   315
      Left            =   210
      TabIndex        =   20
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblResp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   19
      Top             =   2550
      Width           =   3075
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00000080&
      Caption         =   " Débitos disponíveis para parcelamento"
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
      Left            =   60
      TabIndex        =   18
      Top             =   1470
      Width           =   9375
   End
End
Attribute VB_Name = "frmParcelamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nCodReduzido As Long
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
    sDataPago As String
    nValorPago As Double
    nCodBanco As Integer
    dDataPag As Date
End Type

Private Type TRIBUTO
    nCodTributo  As Integer
    sNomeTributo As String
    nValorTributo As Double
    nPercentual As Double
End Type

Dim Rdoaux As rdoResultset, Sql As String, aTributo() As TRIBUTO, aTributos() As TRIBUTO, nValorDil As Double, dDataBase As Date
Dim sNumProc As String, nNumProc As Long, nAnoProc As Integer, sTipoReparc As String, aLancamento()
Dim bIPTU As Boolean, bISS As Boolean, bVS As Boolean, bTLic As Boolean, bDIV As Boolean, bTCD As Boolean, nValorExp As Double

Private Sub chkCorrecao_Click()
CarregaDebito Val(txtCod.text)
AtualizaTotal
End Sub

Private Sub chkHon_Click()
AtualizaTotal
End Sub

Private Sub chkJuros_Click()
CarregaDebito Val(txtCod.text)
AtualizaTotal
End Sub

Private Sub chkMulta_Click()
CarregaDebito Val(txtCod.text)
AtualizaTotal
End Sub

Private Sub cmdAdd_Click()
Dim x As Integer

For x = 1 To lvOrigem.ListItems.Count
    lvOrigem.ListItems(x).Checked = True
Next
AtualizaTotal
End Sub

Private Sub cmdCancel_Click()
If Val(txtCod.text) = 0 Then Exit Sub
If MsgBox("Cancelar a operação ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
Limpar
End Sub

Private Sub cmdDel_Click()
Dim x As Integer

For x = 1 To lvOrigem.ListItems.Count
    lvOrigem.ListItems(x).Checked = False
Next
AtualizaTotal
End Sub

Private Sub cmdGeraDebito_Click()
Dim df As Integer, bAchou As Boolean, x As Integer, bS As Boolean, bN As Boolean, sTCD As Boolean, nPos As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nQtde As Integer, nQtdePago As Integer
Dim bVigS As Boolean, bVigN As Boolean

If lblNome.Caption = "" Then
    MsgBox "Selecione o contribuinte", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataProc.text) Then
    MsgBox "Data do Processo inválido.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtQtdeParc.text) < 2 Or Val(txtQtdeParc) > 60 Then
    MsgBox "A quantidade de parcelas deve ser no mínimo 2 e no máximo 60.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskVencto.text) Then
    MsgBox "Data de vencimento inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If CDate(Format(mskVencto.text, "dd/mm/yyyy")) < CDate(Format(dDataBase, "dd/mm/yyyy")) Then
    MsgBox "Data de vencimento menor que a data base", vbExclamation, "Atenção"
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
    Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With Rdoaux
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

bAchou = False
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Nenhuma parcela foi selecionada.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtValorEntrada.text) > 0 Then
    If Val(txtValorEntrada.text) <= lblValorParcela Or Val(txtValorEntrada.text) >= lblValorTotal Then
        MsgBox "O valor da parcela de entrada não pode ser menor ou igual " & vbCrLf & " ao valor de uma parcela, e não pode ser maior que o " & vbCrLf & " valor total do parcelamento", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

If CDbl(lblValorTotal.Caption) < 60 Then
    MsgBox "Parcelamento mínimo deve ser de R$60,00 reais.", vbExclamation, "Atenção"
    Exit Sub
End If

bVigS = False: bVigN = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            If Val(Left$(.ListItems(x).SubItems(1), 2)) = 13 Then
                bVigS = True
            Else
                bVigN = True
            End If
        End If
    Next
End With
If bVigN And bVigS Then
    MsgBox "Não é possivel parcelar débitos de Vigilância Sanitária com outros débitos " & vbCrLf & "no mesmo parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

bS = False: bN = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            If .ListItems(x).SubItems(6) = "S" Then
                bS = True
            Else
                bN = True
            End If
        End If
    Next
End With
If bN And bS Then
    MsgBox "Não é possivel parcelar débitos ajuizados e não ajuizados " & vbCrLf & "no mesmo parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

CarregaLancamento
bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If nLanc = 1 Or nLanc = 29 Then
                bIPTU = True
            ElseIf nLanc = 2 Or nLanc = 3 Or nLanc = 5 Then
                bISS = True
            ElseIf nLanc = 13 Then
                bVS = True
            ElseIf nLanc = 8 Then
                bTCD = True
                nPos = x 'guarda a posicao do TCD
            ElseIf nLanc = 6 Then
                bTLic = True
            Else
                If nLanc <> 7 Then 'CM pode ser parcelado junto com iPTU
                    bDIV = True
                End If
            End If
        End If
    Next
End With

If bTCD And (bIPTU Or bISS Or bVS Or bDIV) Then
    MsgBox "TCD não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If

If bitpu And (bISS Or bVS Or bDIV) Then
    MsgBox "IPTU não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If
If bISS And (bIPTU Or bVS Or bDIV) Then
    MsgBox "ISS não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If
If bVS And (bISS Or bIPTU Or bDIV) Then
    MsgBox "Vigilância Sanitária não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If
If bDIV And (bISS Or bVS Or bIPTU) Then
    MsgBox "IPTU,ISS e VS não podem ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If

If bTCD Then
    With lvOrigem
        nAno = .ListItems(nPos).text
        nLanc = Val(Left$(.ListItems(nPos).SubItems(1), 2))
        nSeq = .ListItems(nPos).SubItems(2)
        nParc = .ListItems(nPos).SubItems(3)
        nCompl = .ListItems(nPos).SubItems(4)
        Sql = "SELECT NUMPROCESSO FROM DESTINOREPARC WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND ANOEXERCICIO=" & nAno & " AND "
        Sql = Sql & "CODLANCAMENTO=" & 20 & " AND NUMSEQUENCIA=" & 0 & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With Rdoaux
            If .RowCount > 0 Then
                sNumProc = !NUMPROCESSO
            Else
                sNumProc = ""
            End If
           .Close
        End With
        
        If sNumProc <> "" Then
            Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND ANOEXERCICIO=" & nAno & " AND "
            Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
            Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With Rdoaux
                If Not IsNull(!CONTADOR) Then
                    nQtdePago = !CONTADOR
                End If
               .Close
            End With
            
            Sql = "SELECT QTDEPARCELA FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
            Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With Rdoaux
                If .RowCount > 0 Then
                    nQtde = !QTDEPARCela
                Else
                    nQtde = 0
                End If
               .Close
            End With
            If nQtde > 0 Then
                If Val(txtQtdeParc.text) > nQtde Then
                    MsgBox "O máximo de parcelas para este parcelamento é de " & nQtde - nQtdePago & ".", vbExclamation, "Atenção"
                    Exit Sub
                End If
            End If
        End If
    End With
End If

Ocupado
DefineParcelas
Liberado

If CDbl(lblValorParcela.Caption) < 20 Then
    MsgBox "Valor da Parcela mínima deve ser de R$20,00 reais.", vbExclamation, "Atenção"
    Exit Sub
End If

TrocaTela
End Sub

Private Sub cmdGrava_Click()
On Error GoTo Erro
Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long

If MsgBox("Deseja Gravar este Parcelamento ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

nNumProc = Left$(txtNumProc.text, InStr(1, txtNumProc.text, "/", vbBinaryCompare) - 2)
nAnoProc = Right$(txtNumProc.text, 4)
lblNumProc.Caption = nNumProc
lblAnoProc.Caption = nAnoProc
sNumProc = CStr(nNumProc) & "/" & CStr(nAnoProc)

'VERIFICA SE O PROCESSO JA FOI UTILIZADO
Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & " '"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Rdoaux.RowCount > 0 Then
    MsgBox "Este processo já foi utilizado em um parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

'GRAVA O PROCESSO
Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,NUMPROC,ANOPROC,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CODIGORESP,FUNCIONARIO,PLANO,NOVO) VALUES('"
Sql = Sql & sNumProc & "'," & nNumProc & "," & nAnoProc & ",'" & Format(mskDataProc.text, "mm/dd/yyyy") & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "',"
Sql = Sql & Val(txtQtdeParc.text) & "," & Virg2Ponto(txtValorEntrada.text) & "," & 0 & ","
Sql = Sql & IIf(chkMulta.Value = vbChecked, 1, 0) & "," & IIf(chkJuros.Value = vbChecked, 1, 0) & ","
Sql = Sql & Val(txtCod.text) & ",'" & NomeDeLogin & "'," & Val(lblTipoPlano.Caption) & "," & 1 & ")"
cn.Execute Sql, rdExecDirect

'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
    nLastCod = !MAXIMO + 10
   .Close
End With
    
bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If nLanc = 1 Or nLanc = 29 Then
                bIPTU = True
            ElseIf nLanc = 2 Or nLanc = 3 Or nLanc = 5 Or nLanc = 6 Then
                bISS = True
            ElseIf nLanc = 13 Then
                bVS = True
            ElseIf nLanc = 8 Then
                bIPTU = True
            Else
                bDIV = True
            End If
        End If
    Next
End With
    
'GRAVA AS PARCELAS DE DESTINO
With lvDestino
    For x = 1 To .ListItems.Count
        nAno = .ListItems(x).text
        nLanc = 20
        nSeq = .ListItems(x).SubItems(1)
        nParc = .ListItems(x).SubItems(2)
        nCompl = 0
        sDataVencto = .ListItems(x).SubItems(3)
       'GRAVA DESTINOREPARC
        Sql = "INSERT DESTINOREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO) VALUES('" & sNumProc & "'," & Val(txtCod.text) & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA DEBITOPARCELA
        If nAno = Year(dDataBase) Then
            nSituacao = 3
        Else
            nSituacao = 18
        End If
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMPROCESSO) VALUES(" & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
        Sql = Sql & nSituacao & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "')"
        cn.Execute Sql, rdExecDirect
       'DEFINE TRIBUTO PRINCIPAL
        If bIPTU Then
            nCodTrib = 200
        ElseIf bISS Then
            nCodTrib = 201
        ElseIf bVS Then
            nCodTrib = 202
        ElseIf bDIV Then
            nCodTrib = 203
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTOS PRINCIPAL)
        nValorTributo = CDbl(.ListItems(x).SubItems(4))
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
        Sql = Sql & nParc & "," & nCompl & "," & nCodTrib & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 3 - TX.EXP.DOC)
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
        Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
        Sql = Sql & nParc & "," & nCompl & "," & 3 & "," & Virg2Ponto(CStr(nValorExp)) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 113 - JUROS)
        nValorTributo = CDbl(.ListItems(x).SubItems(5))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 113 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 112 - MULTA)
        nValorTributo = CDbl(.ListItems(x).SubItems(6))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 112 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 26 - CORREÇÃO)
        nValorTributo = CDbl(.ListItems(x).SubItems(7))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 26 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
        If chkHon.Value = 1 And CDbl(.ListItems(x).SubItems(8)) > 0 Then
            'GRAVA DEBITOTRIBUTO   // (TRIBUTO 90 - HONORÁRIOS)
             nValorTributo = CDbl(.ListItems(x).SubItems(8))
             Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
             Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
             Sql = Sql & nParc & "," & nCompl & "," & 90 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
             cn.Execute Sql, rdExecDirect
        End If
        If Val(txtQtdeDil.text) > 0 Then
            'GRAVA DEBITOTRIBUTO   // (TRIBUTO 91 - DILIGÊNCIAS)
             nValorTributo = CDbl(.ListItems(x).SubItems(9))
             Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
             Sql = Sql & Val(txtCod.text) & "," & nAno & "," & nLanc & "," & nSeq & ","
             Sql = Sql & nParc & "," & nCompl & "," & 91 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
             cn.Execute Sql, rdExecDirect
        End If
       'GRAVA NUMDOCUMENTO
        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
        Sql = Sql & nLastCod + x & ",'" & Format(dDataBase, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorExp)) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.text) & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
        Sql = Sql & nCompl & "," & nLastCod + x & ")"
        cn.Execute Sql, rdExecDirect
       'PREENCHE DOCUMENTO NO LVDESTINO
        .ListItems(x).SubItems(11) = nLastCod + x
    Next
End With

'GRAVA AS PARCELAS DE ORIGEM
With lvOrigem
    For x = 1 To .ListItems.Count
        If lvOrigem.ListItems(x).Checked = True Then
             nAno = .ListItems(x).text
             nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
             nSeq = .ListItems(x).SubItems(2)
             nParc = .ListItems(x).SubItems(3)
             nCompl = .ListItems(x).SubItems(4)
            'GRAVA ORIGEMREPARC
             Sql = "INSERT ORIGEMREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO) VALUES('" & sNumProc & "'," & Val(txtCod.text) & ","
             Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ")"
             cn.Execute Sql, rdExecDirect
            'ATUALIZA O STATUS DE ORIGEM   // (4 - REPARCELADO)
             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=4 WHERE CODREDUZIDO=" & Val(txtCod.text)
             Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
             Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
             Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
             cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'LIMPA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

'DADOS CABEÇALHO
dDataProc = CDate(mskDataProc.text)
sDescImposto = "REPARCELAMENTO"
NumBarra1 = Format(ExtraiNumero(txtNumProc.text), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)

'ENDEREÇO DO CONTRIBUINTE
Select Case Val(txtCod.text)
    Case 1 To 99999
        'DADOS DO IMOVEL
        Sql = "SELECT * FROM vwCnsImovel WHERE CODREDUZIDO=" & Val(txtCod.text)
        Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With Rdoaux
            If .RowCount > 0 Then
                sNumInsc = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
                sCodReduz = Format(Val(txtCod.text), "000000") & "-" & RetornaDVCodReduzido(Val(txtCod.text))
                Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
                Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
                Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & Val(txtCod.text)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                sNomeResp = RdoAux2!NOMECIDADAO
                RdoAux2.Close
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
               Sql = Sql & "WHERE CODREDUZIDO=" & Val(txtCod.text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                      If .RowCount > 0 Then
                          sEndEntrega = !Ee_NomeLog
                          nNumEntrega = !Ee_NumImovel
                          If IsNull(!DescBairro) Then
                            sBairroEntrega = ""
                          Else
                            sBairroEntrega = IIf(!DescBairro = 999, " ", !DescBairro)
                          End If
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
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & Val(txtCod.text)
        Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With Rdoaux
            If .RowCount > 0 Then
                sNumInsc = SubNull(!INSCESTADUAL)
                sCodReduz = !CODIGOMOB
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
        Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao,cidadao.cpf, cidadao.cnpj, cidadao.rg, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, "
        Sql = Sql & "cidadao.siglauf, cidade.desccidade, bairro.descbairro, cidadao.nomelogradouro AS nomerua, cidadao.nomebairro, cidadao.nomecidade,"
        Sql = Sql & "cidadao.codlogradouro , vwLOGRADOURO.AbrevTipoLog, vwLOGRADOURO.AbrevTitLog, vwLOGRADOURO.NomeLogradouro "
        Sql = Sql & "FROM  vwLOGRADOURO RIGHT JOIN  cidadao ON vwLOGRADOURO.CODLOGRADOURO = cidadao.codlogradouro LEFT OUTER JOIN "
        Sql = Sql & "cidade INNER JOIN  bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade And cidadao.codbairro = bairro.codbairro "
        Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & Val(txtCod.text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sTipoImposto = "REPARCEL."
                sCodReduz = !CodCidadao
                sNomeResp = !NOMECIDADAO
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

'GRAVA TEMPORARIO
With lvDestino
    For x = 1 To .ListItems.Count
'    For X = 1 To 1
        nAno = .ListItems(x).text
        nSeq = .ListItems(x).SubItems(1)
        nParc = .ListItems(x).SubItems(2)
        nNumDoc = .ListItems(x).SubItems(11)
        sDataVencto = .ListItems(x).SubItems(3)
        sValorParc = .ListItems(x).SubItems(10)
        NumBarra2 = Gera2of5Cod(CStr(CDbl(sValorParc) + nValorExp), CDate(sDataVencto), nNumDoc, nParc, 20, nSeq, 0)
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
        Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Left$(sDescImposto, 30) & "'," & nAno & ",'" & Left$(txtNumProc.text, 25) & "','"
        Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "'," & nNumDoc & "," & RetornaDVNumDoc(nNumDoc) & ",'" & sQuadra & "','"
        Sql = Sql & sLote & "','" & Format(sDataVencto, "mm/dd/yyyy") & "'," & IIf(nParc = 0, 1, nParc) & "," & Val(txtQtdeParc.text) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
        Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & "REPARCELAMENTO" & "'," & Virg2Ponto(CDbl(nValorExp)) & "," & "0" & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

frmConfissaoDivida.txtNumProc.text = sNumProc
frmConfissaoDivida.txtNumProc.Locked = True
frmConfissaoDivida.CarregaProcesso
cmdVoltar_Click
Limpar
Unload frmParcelamento
frmConfissaoDivida.show

Exit Sub

Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdTributos_Click()
Dim sTributo As String, x As Integer
CarregaTributos

For x = 1 To UBound(aTributos)
    sTributo = sTributo & aTributos(x).sNomeTributo & vbCrLf
Next
MsgBox sTributo, vbOKOnly, "Lista dos tributos contidos nas parcelas selecionadas."

End Sub

Private Sub cmdSimulado_Click()
Simulado
End Sub

Private Sub cmdVoltar_Click()
TrocaTela
lblValorParcela.Caption = "0,00"
End Sub

Private Sub Form_Load()
Centraliza Me
frmMdi.AddWindow Me.Name, Me.Caption
dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).text, 4))
lblDataParc.Caption = Format(dDataBase, "dd/mm/yyyy")
bExec = True

'carrega valor da diligência
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & Year(Now) & " AND CODTRIBUTO=91"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
     If .RowCount > 0 Then
         nValorDil = FormatNumber(!VALORALIQ, 2)
     Else
         MsgBox "Taxa de Diligência não cadastrado para este ano.", vbExclamation, "Atenção"
     End If
    .Close
End With

'BUSCA O VALOR DA TAXA DE EXPEDIENTE
Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 1"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
    nValorExp = FormatNumber(!VALORPARCELA, 2)
   .Close
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cmdSair.Enabled = False Then
    Cancel = 1
Else
    frmMdi.RemoveWindow Me.Name
End If
End Sub


Private Sub lvDestino_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Item.Checked = True
End Sub

Private Sub lvOrigem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
AtualizaTotal
End Sub

Private Sub mskVencto_GotFocus()
mskVencto.SetFocus
mskVencto.SelStart = 0
mskVencto.SelLength = Len(mskVencto.text)
End Sub

Private Sub txtCod_Change()
lvOrigem.ListItems.Clear
lblNome.Caption = ""
End Sub

Private Sub txtCod_GotFocus()
txtCod.SetFocus
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_LostFocus()
Dim nCodReduz As Long, sTipoCod As String

If Val(txtCod.text) = 0 Then Exit Sub
If Val(txtCod.text) = 0 Then
    lblNome.Caption = ""
    Exit Sub
End If
If Val(txtCod.text) < 100000 Then
    sTipoCod = "I"
ElseIf Val(txtCod.text) >= 100000 And Val(txtCod.text) < 500000 Then
    sTipoCod = "M"
ElseIf Val(txtCod.text) >= 500000 Then
    sTipoCod = "C"
End If
txtCod.text = Format(txtCod.text, "000000")
nCodReduz = Val(txtCod.text)
lblNome.Caption = ""
If sTipoCod = "I" Then
    Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
    Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
    Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & nCodReduz & " AND TIPOPROP='P'"
ElseIf sTipoCod = "M" Then
    Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO Where CODIGOMOB =" & nCodReduz
ElseIf sTipoCod = "C" Then
    Sql = "SELECT NOMECIDADAO FROM CIDADAO Where CODCIDADAO =" & nCodReduz
End If
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
    If Rdoaux.RowCount > 0 Then
         If sTipoCod = "I" Or sTipoCod = "C" Then
            lblNome.Caption = !NOMECIDADAO
         ElseIf sTipoCod = "M" Then
            lblNome.Caption = !RAZAOSOCIAL
         End If
    Else
       MsgBox "Código não Cadastrado.", vbExclamation, "Atenção"
       lvOrigem.ListItems.Clear
       txtCod.SetFocus
       Exit Sub
    End If
    .Close
End With

CarregaDebito (nCodReduz)
If lvOrigem.ListItems.Count = 0 Then
    MsgBox "O contribuinte não possue débitos a serem parcelados.", vbExclamation, "Atenção"
    txtCod.text = 0
    txtCod.SetFocus
Else
    txtCod.Locked = True
    txtCod.BackColor = Kde
End If

End Sub

Private Sub CarregaDebito(nCodReduz As Long)

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date
Dim nSomaValorTributo As Double, sAj As String
Dim x As Integer
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, Achou As Boolean, sDescLanc As String
Dim itmX As ListItem
Ocupado
lvOrigem.ListItems.Clear
ReDim aDebito(0)
Set qd.ActiveConnection = cn
On Error Resume Next
RdoAux3.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = nCodReduz
qd(1) = nCodReduz 'codigo
qd(2) = 1970
qd(3) = 2020 'ano
qd(4) = 1
qd(5) = 99 'lancamento
qd(6) = 0
qd(7) = 99 'sequencia
qd(8) = 1
qd(9) = 12 'numparcela
qd(10) = 0
qd(11) = 9 'complemento
qd(12) = 3
qd(13) = 3 'statuslanc
qd(14) = Format(Now, "mm/dd/yyyy") 'data atual
qd(15) = NomeDoUsuario
Set RdoAux3 = qd.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
With RdoAux3
    Do Until .EOF
        'VALIDAÇÃO
        'If !CodLancamento = 7 Then MsgBox "TESTE"
        If !CodLancamento <> 8 And IsNull(!DATAINSCRICAO) Then GoTo proximo
        If !CodLancamento = 20 Then GoTo proximo
        'sDescLanc = !desclancamento
        'CARREGA MATRIZ DE DÉBITO
        nEval = UBound(aDebito)
        Achou = False
        For x = 1 To nEval
            If aDebito(x).nCodReduzido = !CODREDUZIDO And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
               aDebito(x).nSeq = !SeqLancamento And aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
               Achou = True
               Exit For
            End If
        Next
        'SE NÃO ENCONTRAR O LANCAMENTO NA MATRIZ, ADICIONAR ELE
        If Not Achou Then
           ReDim Preserve aDebito(UBound(aDebito) + 1)
           nEval = UBound(aDebito)
           aDebito(nEval).nCodReduzido = !CODREDUZIDO
           aDebito(nEval).nAno = !AnoExercicio
           aDebito(nEval).sLanc = !DESCLANCAMENTO
           aDebito(nEval).nLanc = !CodLancamento
           aDebito(nEval).nSeq = !SeqLancamento
           aDebito(nEval).nParc = !NumParcela
           aDebito(nEval).nCompl = !CODCOMPLEMENTO
           aDebito(nEval).nSituacao = !STATUSLANC
           aDebito(nEval).sSituacao = !SITUACAO
           aDebito(nEval).sVencto = Format(!DATAVENCIMENTO, "dd/mm/yyyy")
           aDebito(nEval).nValorTributo = FormatNumber(!VALORTRIBUTO, 2)
           If chkJuros.Value = 1 Then
                aDebito(nEval).nValorJuros = FormatNumber(!VALORJUROS, 2)
           Else
                aDebito(nEval).nValorJuros = FormatNumber(0, 2)
           End If
           If chkMulta.Value = 1 Then
                aDebito(nEval).nValorMulta = FormatNumber(!VALORMULTA, 2)
           Else
                aDebito(nEval).nValorMulta = FormatNumber(0, 2)
           End If
           If chkCorrecao.Value = 1 Then
                aDebito(nEval).nValorCorrecao = FormatNumber(!VALORCORRECAO, 2)
           Else
                aDebito(nEval).nValorCorrecao = FormatNumber(0, 2)
           End If
           aDebito(nEval).nValorAtual = !ValorTotal
           If IsNull(!DATAAJUIZA) Then
                aDebito(nEval).sAj = "N"
           Else
                aDebito(nEval).sAj = "S"
           End If
        Else
          'SE ENCONTRAR ADICIONAR O VALOR AO JA EXISTENTE
           aDebito(x).nValorAtual = aDebito(x).nValorAtual + !ValorTotal
           If chkJuros.Value = 1 Then
                aDebito(x).nValorJuros = aDebito(x).nValorJuros + !VALORJUROS
           End If
           If chkMulta.Value = 1 Then
                aDebito(x).nValorMulta = aDebito(x).nValorMulta + !VALORMULTA
           End If
           If chkCorrecao.Value = 1 Then
                aDebito(x).nValorCorrecao = aDebito(x).nValorCorrecao + !VALORCORRECAO
           End If
           aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !VALORTRIBUTO, 2)
        End If
proximo:
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(aDebito)
    With aDebito(x)
        If .nValorTributo > 0 Then
            Set itmX = lvOrigem.ListItems.Add(, , .nAno)
            itmX.SubItems(1) = Format(.nLanc, "00") & "-" & .sLanc
            itmX.SubItems(2) = Format(.nSeq, "00")
            itmX.SubItems(3) = Format(.nParc, "00")
            itmX.SubItems(4) = Format(.nCompl, "00")
            itmX.SubItems(5) = Format(.sVencto, "dd/mm/yyyy")
            itmX.SubItems(6) = .sAj
            itmX.SubItems(7) = FormatNumber(.nValorTributo, 2)
            itmX.SubItems(8) = FormatNumber(.nValorJuros, 2)
            itmX.SubItems(9) = FormatNumber(.nValorMulta, 2)
            itmX.SubItems(10) = FormatNumber(.nValorCorrecao, 2)
            itmX.SubItems(11) = FormatNumber(.nValorTributo + .nValorJuros + .nValorMulta + .nValorCorrecao, 2)
         End If
    End With
Next
Liberado
End Sub

Private Sub txtNumProc_LostFocus()
Dim sValidaProc As String
On Error Resume Next
If Trim$(txtNumProc.text) <> "" Then
    If InStr(1, txtNumProc.text, "/", vbBinaryCompare) > 0 Then
        nNumProc = Left$(txtNumProc.text, InStr(1, txtNumProc.text, "/", vbBinaryCompare) - 2)
        nAnoProc = Right$(txtNumProc.text, 4)
        lblNumProc.Caption = nNumProc
        lblAnoProc.Caption = nAnoProc
        sNumProc = CStr(nNumProc) & "/" & CStr(nAnoProc)
        sValidaProc = ValidaProcesso(sNumProc)
        If sValidaProc <> "OK" Then
            MsgBox sValidaProc, vbCritical, "Atenção"
            txtNumProc.SetFocus
            Exit Sub
        Else
            mskDataProc.text = Format(RetornaDataProcesso(nNumProc, nAnoProc), "dd/mm/yyyy")
        End If
    Else
        MsgBox "Processo inválido.", vbExclamation, "Atenção"
        txtNumProc.SetFocus
    End If
End If
End Sub

Private Sub txtQtdeDil_Change()
If Val(txtQtdeDil) = 0 Then
   lblSomaDil.Caption = "0,00"
Else
   lblSomaDil.Caption = FormatNumber(CDbl(txtQtdeDil.text) * nValorDil, 2)
End If
AtualizaTotal
End Sub

Private Sub txtQtdeDil_GotFocus()
txtQtdeDil.SetFocus
txtQtdeDil.SelStart = 0
txtQtdeDil.SelLength = Len(txtQtdeDil.text)
End Sub

Private Sub txtQtdeDil_KeyPress(KeyAscii As Integer)
Tweak txtQtdeDil, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtdeDil_LostFocus()
If Trim$(txtQtdeDil.text) = "" Then txtQtdeDil.text = "0"
AtualizaTotal
End Sub

Private Sub txtQtdeParc_GotFocus()
txtQtdeParc.SetFocus
txtQtdeParc.SelStart = 0
txtQtdeParc.SelLength = Len(txtQtdeParc.text)
End Sub

Private Sub txtQtdeParc_KeyPress(KeyAscii As Integer)
Tweak txtQtdeParc, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtdeParc_LostFocus()
If Trim$(txtQtdeParc.text) = "" Then txtQtdeParc.text = "0"
End Sub

Private Sub txtValorEntrada_GotFocus()
txtValorEntrada.SetFocus
txtValorEntrada.SelStart = 0
txtValorEntrada.SelLength = Len(txtValorEntrada.text)
End Sub

Private Sub txtValorEntrada_KeyPress(KeyAscii As Integer)
Tweak txtValorEntrada, KeyAscii, DecimalPositive
End Sub

Private Sub AtualizaTotal()

Dim x As Integer, nContaParcela As Integer, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double
Dim nValorCorrecao As Double, nValortotal As Double, nValorAjuizado As Double

LimpaContador
If lvOrigem.ListItems.Count = 0 Then Exit Sub

nContaParcela = 0: nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValortotal = 0: nValorAjuizado = 0

With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nContaParcela = nContaParcela + 1
            nValorPrincipal = nValorPrincipal + CDbl(.ListItems(x).ListSubItems(7))
            If chkJuros.Value = 1 Then
                nValorJuros = nValorJuros + CDbl(.ListItems(x).ListSubItems(8))
            End If
            If chkMulta.Value = 1 Then
                nValorMulta = nValorMulta + CDbl(.ListItems(x).ListSubItems(9))
            End If
            If chkCorrecao.Value = 1 Then
                nValorCorrecao = nValorCorrecao + CDbl(.ListItems(x).ListSubItems(10))
            End If
            If chkHon.Value = 1 Then
                If .ListItems(x).SubItems(6) = "S" Then
                    nValorAjuizado = nValorAjuizado + CDbl(.ListItems(x).SubItems(11))
                End If
            End If
            nValortotal = nValortotal + CDbl(.ListItems(x).SubItems(11))
        End If
    Next
End With
If txtQtdeDil.text = "" Then txtQtdeDil.text = 0
lblNumParc.Caption = Format(nContaParcela, "000")
lblValorCorrecao.Caption = FormatNumber(nValorCorrecao, 2)
lblValorJuros.Caption = FormatNumber(nValorJuros, 2)
lblValorMulta.Caption = FormatNumber(nValorMulta, 2)
lblValorPrincipal.Caption = FormatNumber(nValorPrincipal, 2)
If nContaParcela > 0 Then
    lblValorDil.Caption = FormatNumber(Val(txtQtdeDil.text) * CDbl(lblSomaDil.Caption), 2)
End If
lblValorHon.Caption = FormatNumber(nValorAjuizado * 0.1, 2)
lblValorTotal.Caption = FormatNumber(nValortotal + CDbl(lblValorHon.Caption) + CDbl(lblValorDil.Caption), 2)

End Sub

Private Sub LimpaContador()
lblValorParcela.Caption = "0,00"
lblNumParc.Caption = "000"
lblValorCorrecao.Caption = "0,00"
lblValorMulta.Caption = "0,00"
lblValorJuros.Caption = "0,00"
lblValorTotal.Caption = "0,00"
lblValorPrincipal.Caption = "0,00"
lblValorDil.Caption = "0,00"
lblValorHon.Caption = "0,00"
End Sub

Private Sub TrocaTela()
If cmdGeraDebito.Visible = True Then
    cmdGeraDebito.Visible = False
    cmdVoltar.Visible = True
    cmdAdd.Enabled = False
    cmdDel.Enabled = False
    cmdSair.Enabled = False
    txtNumProc.Locked = True
    txtNumProc.BackColor = Kde
    txtQtdeParc.Locked = True
    txtQtdeParc.BackColor = Kde
    mskVencto.Locked = True
    mskVencto.BackColor = Kde
    txtValorEntrada.Locked = True
    txtValorEntrada.BackColor = Kde
    txtQtdeDil.Locked = True
    txtQtdeDil.BackColor = Kde
    chkMulta.Enabled = False
    chkJuros.Enabled = False
    chkCorrecao.Enabled = False
    chkHon.Enabled = False
    lvOrigem.Visible = False
    lvDestino.Visible = True
    lblTitulo.Caption = " Débitos que serão gerados no parcelamento"
    cmdGrava.Visible = True
    cmdCancel.Visible = False
    'cmdTributos.Enabled = True
    lvDestino.SetFocus
Else
    cmdGeraDebito.Visible = True
    cmdVoltar.Visible = False
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    cmdSair.Enabled = True
    txtNumProc.Locked = False
    txtNumProc.BackColor = Branco
    txtQtdeParc.Locked = False
    txtQtdeParc.BackColor = Branco
    mskVencto.Locked = False
    mskVencto.BackColor = Branco
    txtValorEntrada.Locked = False
    txtValorEntrada.BackColor = Branco
    txtQtdeDil.Locked = False
    txtQtdeDil.BackColor = Branco
    chkMulta.Enabled = True
    chkJuros.Enabled = True
    chkCorrecao.Enabled = True
    chkHon.Enabled = True
    lvOrigem.Visible = True
    lvDestino.Visible = False
    lblTitulo.Caption = " Débitos disponíveis para parcelamento"
    cmdGrava.Visible = False
    cmdCancel.Visible = True
    'cmdTributos.Enabled = False
    cmdGeraDebito.SetFocus
End If

End Sub

Private Sub DefineParcelas()
Dim x As Integer, nSeq As Integer, a As Integer, sVencimento As String, sVencimento2 As String, y As Integer, nValorEntrada As Double
Dim nValorPrimeira As Double, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValortotal As Double, nValorHon As Double, nValorDil As Double
Dim nValorPrincipal1 As Double, nValorJuros1 As Double, nValorMulta1 As Double, nValorCorrecao1 As Double, nValorTotal1 As Double, nValorHon1 As Double
Dim nPrincipal As Double, nJuros As Double, nMulta As Double, nCorrecao As Double, nTotal As Double, nHonorario As Double, nDiligencia As Double, nItem As Integer
Dim nDia As Integer, nMes As Integer, nAno As Integer, itmX As ListItem, nQtdeParc As Integer, nDif As Double, nPerc As Double
'LIMPA TELA
lvDestino.ListItems.Clear

'BUSCA ULTIMA SEQUENCIA
Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND CODLANCAMENTO=20"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
    If !CONTADOR > 0 Then
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO<100"
        Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nSeq = Rdoaux!SEQMAXIMA + 1
    Else
        nSeq = 0
    End If
   .Close
End With

'ADICIONA ITENS
nQtdeParc = Val(txtQtdeParc.text)
sVencimento = mskVencto.text
If txtValorEntrada.text = "" Then txtValorEntrada.text = "0"
'CALCULA VALORES
nValorEntrada = CDbl(txtValorEntrada.text)
nPrincipal = CDbl(lblValorPrincipal.Caption)
nJuros = CDbl(lblValorJuros.Caption)
nMulta = CDbl(lblValorMulta.Caption)
nCorrecao = CDbl(lblValorCorrecao.Caption)
nHonorario = CDbl(lblValorHon.Caption)
nDiligencia = CDbl(lblValorDil.Caption)
nTotal = CDbl(lblValorTotal.Caption)

If nValorEntrada > 0 Then
    nItem = 2
    nPerc = nValorEntrada / nTotal
    nValorPrincipal1 = FormatNumber(nPrincipal * nPerc, 2)
    nValorJuros1 = FormatNumber(nJuros * nPerc, 2)
    nValorMulta1 = FormatNumber(nMulta * nPerc, 2)
    nValorCorrecao1 = FormatNumber(nCorrecao * nPerc, 2)
    nValorHon1 = FormatNumber(nHonorario * nPerc, 2)
    nValorDil1 = FormatNumber(nDiligencia * nPerc, 2)
    nValorPrincipal = (nPrincipal - nValorPrincipal1) / (nQtdeParc - 1)
    nValorJuros = (nJuros - nValorJuros1) / (nQtdeParc - 1)
    nValorMulta = (nMulta - nValorMulta1) / (nQtdeParc - 1)
    nValorCorrecao = (nCorrecao - nValorCorrecao1) / (nQtdeParc - 1)
    nValorHon = (nHonorario - nValorHon1) / (nQtdeParc - 1)
    nValorDil = (nDiligencia - nValorDil1) / (nQtdeParc - 1)
Else
    nItem = 1
    nValorPrincipal = nPrincipal / nQtdeParc
    nValorJuros = nJuros / nQtdeParc
    nValorMulta = nMulta / nQtdeParc
    nValorCorrecao = nCorrecao / nQtdeParc
    nValorHon = nHonorario / nQtdeParc
    nValorDil = nDiligencia / nQtdeParc
    nValorPrincipal1 = nValorPrincipal
    nValorJuros1 = nValorJuros
    nValorMulta1 = nValorMulta
    nValorCorrecao1 = nValorCorrecao
    nValorHon1 = nValorHon
    nValorDil1 = nValorDil
End If
nValortotal = nValorPrincipal + nValorHon + nValorDil + nValorJuros + nValorMulta + nValorCorrecao
nValorTotal1 = nValorPrincipal1 + nValorHon1 + nValorDil1 + nValorJuros1 + nValorMulta1 + nValorCorrecao1
For x = 1 To nQtdeParc
    'CALCULA VENCIMENTO
     If x > 1 Then
       nDia = Val(Left$(mskVencto.text, 2))
       nMes = Val(Mid$(Format(sVencimento, "dd/mm/yyyy"), 4, 2)) + 1
       
       nAno = Val(Right$(sVencimento, 4))
       If nMes = 13 Then
          nMes = 1: nAno = nAno + 1
       End If
               
       sVencimento = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
       
       If Not IsDate(sVencimento) Then
           If nMes = 2 Then
                nDia = nDia - 3
           Else
                nDia = nDia - 1
           End If
           sVencimento = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
       End If
    Else
       nAno = Val(Right$(sVencimento, 4))
    End If
    'sVencimento2 = RetornaDiaUtil(CDate(sVencimento))
    sVencimento2 = sVencimento
    'PREENCHE A LISTA
    Set itmX = lvDestino.ListItems.Add(, , nAno)
    itmX.SubItems(1) = Format(nSeq, "00")
    itmX.SubItems(2) = Format(x, "00")
    itmX.SubItems(3) = Format(sVencimento2, "dd/mm/yyyy")
    If x = 1 Then
        itmX.SubItems(4) = FormatNumber(nValorPrincipal1, 2)
        itmX.SubItems(5) = FormatNumber(nValorJuros1, 2)
        itmX.SubItems(6) = FormatNumber(nValorMulta1, 2)
        itmX.SubItems(7) = FormatNumber(nValorCorrecao1, 2)
        itmX.SubItems(8) = FormatNumber(nValorHon1, 2)
        itmX.SubItems(9) = FormatNumber(nValorDil1, 2)
        itmX.SubItems(10) = FormatNumber(nValorTotal1, 2)
    Else
        itmX.SubItems(4) = FormatNumber(nValorPrincipal, 2)
        itmX.SubItems(5) = FormatNumber(nValorJuros, 2)
        itmX.SubItems(6) = FormatNumber(nValorMulta, 2)
        itmX.SubItems(7) = FormatNumber(nValorCorrecao, 2)
        itmX.SubItems(8) = FormatNumber(nValorHon, 2)
        itmX.SubItems(9) = FormatNumber(nValorDil, 2)
        itmX.SubItems(10) = FormatNumber(nValortotal, 2)
    End If
    If nAno > Year(dDataBase) Then
        For y = 1 To 10
            itmX.ForeColor = vbRed
            itmX.ListSubItems(y).ForeColor = vbRed
        Next y
    End If
    itmX.Checked = True
Next
' Exit Sub
'CORRIGE ARREDONDAMENTO
nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValortotal = 0: nValorHon = 0: nValorDil = 0
With lvDestino
    For x = 1 To .ListItems.Count
        nValorPrincipal = nValorPrincipal + CDbl(.ListItems(x).SubItems(4))
        nValorJuros = nValorJuros + CDbl(.ListItems(x).SubItems(5))
        nValorMulta = nValorMulta + CDbl(.ListItems(x).SubItems(6))
        nValorCorrecao = nValorCorrecao + CDbl(.ListItems(x).SubItems(7))
        nValorHon = nValorHon + CDbl(.ListItems(x).SubItems(8))
        nValorDil = nValorDil + CDbl(.ListItems(x).SubItems(9))
        nValortotal = nValortotal + CDbl(.ListItems(x).SubItems(10))
    Next
    
    nDif = CDbl(lblValorPrincipal.Caption) - nValorPrincipal
    .ListItems(nItem).SubItems(4) = FormatNumber(CDbl(.ListItems(nItem).SubItems(4)) + nDif, 2)
    nDif = CDbl(lblValorJuros.Caption) - nValorJuros
    .ListItems(nItem).SubItems(5) = FormatNumber(CDbl(.ListItems(nItem).SubItems(5)) + nDif, 2)
    nDif = CDbl(lblValorMulta.Caption) - nValorMulta
    .ListItems(nItem).SubItems(6) = FormatNumber(CDbl(.ListItems(nItem).SubItems(6)) + nDif, 2)
    nDif = CDbl(lblValorCorrecao.Caption) - nValorCorrecao
    .ListItems(nItem).SubItems(7) = FormatNumber(CDbl(.ListItems(nItem).SubItems(7)) + nDif, 2)
    nDif = CDbl(lblValorHon.Caption) - nValorHon
    .ListItems(nItem).SubItems(8) = FormatNumber(CDbl(.ListItems(nItem).SubItems(8)) + nDif, 2)
    nDif = CDbl(lblValorDil.Caption) - nValorDil
    .ListItems(nItem).SubItems(9) = FormatNumber(CDbl(.ListItems(nItem).SubItems(9)) + nDif, 2)
    nDif = CDbl(lblValorTotal.Caption) - nValortotal
    .ListItems(nItem).SubItems(10) = FormatNumber(CDbl(.ListItems(nItem).SubItems(10)) + nDif, 2)
End With

'VALOR DA PARCELA
If lvDestino.ListItems.Count < 3 Then
    lblValorParcela.Caption = FormatNumber(CDbl(lvDestino.ListItems(2).SubItems(10)) + nValorExp, 2)
Else
    lblValorParcela.Caption = FormatNumber(CDbl(lvDestino.ListItems(3).SubItems(10)) + nValorExp, 2)
End If

End Sub

Private Sub CarregaTributos()
Dim x As Integer, y As Integer, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim nCodTributo As Integer, sDescTributo As String, bAchou As Boolean

ReDim aTributos(0)
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nAno = .ListItems(x).text
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            nSeq = .ListItems(x).SubItems(2)
            nParc = .ListItems(x).SubItems(3)
            nCompl = .ListItems(x).SubItems(4)
            Sql = "SELECT debitotributo.codtributo,tributo.desctributo FROM debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
            Sql = Sql & "Where debitotributo.CODREDUZIDO =  " & Val(txtCod.text) & " And debitotributo.AnoExercicio = " & Val(nAno) & " And debitotributo.CodLancamento = " & Val(nLanc) & " And "
            Sql = Sql & "debitotributo.SeqLancamento = " & Val(nSeq) & " AND debitotributo.numparcela = " & nParc & " AND debitotributo.codcomplemento = " & nCompl
            Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With Rdoaux
                Do Until .EOF
                    nCodTributo = !CodTributo
                    sDescTributo = !DESCTRIBUTO
                    bAchou = False
                    For y = 1 To UBound(aTributos)
                        If aTributos(y).nCodTributo = nCodTributo Then
                           bAchou = True
                           Exit For
                        End If
                    Next
                    If Not bAchou Then
                        ReDim Preserve aTributos(UBound(aTributos) + 1)
                        aTributos(UBound(aTributos)).nCodTributo = nCodTributo
                        aTributos(UBound(aTributos)).sNomeTributo = sDescTributo
                    End If
                   .MoveNext
                Loop
               .Close
            End With
        End If
    Next
End With

If chkHon.Value = 1 And Val(lblValorHon.Caption) > 0 Then
    ReDim Preserve aTributos(UBound(aTributos) + 1)
    aTributos(UBound(aTributos)).nCodTributo = 90
    aTributos(UBound(aTributos)).sNomeTributo = "HONORARIOS ADVOCATÍCIOS"
End If
If Val(txtQtdeDil.text) > 0 Then
    ReDim Preserve aTributos(UBound(aTributos) + 1)
    aTributos(UBound(aTributos)).nCodTributo = 91
    aTributos(UBound(aTributos)).sNomeTributo = "DILIGÊNCIAS"
End If

End Sub

Private Sub CarregaLancamento()
Dim x As Integer, bAchou As Boolean, nLanc As Integer

ReDim aLancamento(0)
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            bAchou = False
            For y = 1 To UBound(aLancamento)
                If aLancamento(y) = nLanc Then
                   bAchou = True
                   Exit For
                End If
            Next
            If Not bAchou Then
                ReDim Preserve aLancamento(UBound(aLancamento) + 1)
                aLancamento(UBound(aLancamento)) = nLanc
            End If
        End If
    Next
End With

End Sub

Private Sub Limpar()
txtCod.text = ""
txtNumProc.text = ""
lblNome.Caption = ""
LimpaMascara mskDataProc
LimpaMascara mskVencto
txtQtdeParc.text = ""
txtQtdeDil.text = ""
lblValorDil.Caption = ""
chkMulta.Value = 1
chkJuros.Value = 1
chkHon.Value = 1
txtValorEntrada.text = ""
lblTipoPlano.Caption = ""
lblValorPlano.Caption = ""
lvOrigem.ListItems.Clear
txtCod.Locked = False
txtCod.BackColor = Branco
LimpaContador

End Sub

Private Sub Simulado()
Dim df As Integer, bAchou As Boolean, bS As Boolean, bN As Boolean, nValorParcela As Double, nLanc As Integer
Dim x As Integer, nSeq As Integer, a As Integer, sVencimento As String, sVencimento2 As String, y As Integer, nValorEntrada As Double
Dim nValorPrimeira As Double, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValortotal As Double, nValorHon As Double, nValorDil As Double
Dim nValorPrincipal1 As Double, nValorJuros1 As Double, nValorMulta1 As Double, nValorCorrecao1 As Double, nValorTotal1 As Double, nValorHon1 As Double
Dim nPrincipal As Double, nJuros As Double, nMulta As Double, nCorrecao As Double, nTotal As Double, nHonorario As Double, nDiligencia As Double, nItem As Integer
Dim nDia As Integer, nMes As Integer, nAno As Integer, itmX As ListItem, nQtdeParc As Integer, nDif As Double, nPerc As Double

bAchou = False
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Nenhuma parcela foi selecionada.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtValorEntrada.text) > 0 Then
    MsgBox "Simulado não calcula parcelas com valor de entrada.", vbExclamation, "Atenção"
    Exit Sub
End If

If CDbl(lblValorTotal.Caption) < 60 Then
    MsgBox "Parcelamento mínimo deve ser de R$60,00 reais.", vbExclamation, "Atenção"
    Exit Sub
End If

bS = False: bN = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            If .ListItems(x).SubItems(6) = "S" Then
                bS = True
            Else
                bN = True
            End If
        End If
    Next
End With
If bN And bS Then
    MsgBox "Não é possivel parcelar débitos ajuizados e não ajuizados " & vbCrLf & "no mesmo parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

CarregaLancamento
bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If Val(nLanc) = 1 Or Val(nLanc) = 29 Or Val(nLanc) = 7 Then
                bIPTU = True
            ElseIf Val(nLanc) = 2 Or Val(nLanc) = 3 Or Val(nLanc) = 5 Then
                bISS = True
            ElseIf Val(nLanc) = 13 Then
                bVS = True
            ElseIf Val(nLanc) = 8 Then
                bTCD = True
            ElseIf Val(nLanc) = 6 Then
                bTLic = True
            Else
                bDIV = True
            End If
        End If
    Next
End With

If bTCD And (bIPTU Or bISS Or bVS Or bDIV) Then
    MsgBox "TCD não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If

If bitpu And (bISS Or bVS Or bDIV) Then
    MsgBox "IPTU não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If
If bISS And (bIPTU Or bVS Or bDIV) Then
    MsgBox "ISS não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If
If bVS And (bISS Or bIPTU Or bDIV) Then
    MsgBox "Vigilância Sanitária não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If
If bDIV And (bISS Or bVS Or bIPTU) Then
    MsgBox "IPTU,ISS e VS não podem ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
    Exit Sub
End If

Ocupado

Sql = "DELETE FROM SIMULADOREPARC WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

'ADICIONA ITENS
For nQtdeParc = 2 To 60
    'LIMPA TELA
    lvDestino.ListItems.Clear
    
    sVencimento = Format(dDataBase, "mm/dd/yyyy")
    'CALCULA VALORES
    nPrincipal = CDbl(lblValorPrincipal.Caption)
    nJuros = CDbl(lblValorJuros.Caption)
    nMulta = CDbl(lblValorMulta.Caption)
    nCorrecao = CDbl(lblValorCorrecao.Caption)
    nHonorario = CDbl(lblValorHon.Caption)
    nDiligencia = CDbl(lblValorDil.Caption)
    nTotal = CDbl(lblValorTotal.Caption)
    
    nItem = 1
    nValorPrincipal = nPrincipal / nQtdeParc
    nValorJuros = nJuros / nQtdeParc
    nValorMulta = nMulta / nQtdeParc
    nValorCorrecao = nCorrecao / nQtdeParc
    nValorHon = nHonorario / nQtdeParc
    nValorDil = nDiligencia / nQtdeParc
    nValorPrincipal1 = nValorPrincipal
    nValorJuros1 = nValorJuros
    nValorMulta1 = nValorMulta
    nValorCorrecao1 = nValorCorrecao
    nValorHon1 = nValorHon
    nValorDil1 = nValorDil
    nValortotal = nValorPrincipal + nValorHon + nValorDil + nValorJuros + nValorMulta + nValorCorrecao
    
    For x = 1 To nQtdeParc
        'PREENCHE A LISTA
        Set itmX = lvDestino.ListItems.Add(, , nAno)
        itmX.SubItems(1) = Format(nSeq, "00")
        itmX.SubItems(2) = Format(x, "00")
        itmX.SubItems(3) = Format(sVencimento2, "dd/mm/yyyy")
        itmX.SubItems(4) = FormatNumber(nValorPrincipal, 2)
        itmX.SubItems(5) = FormatNumber(nValorJuros, 2)
        itmX.SubItems(6) = FormatNumber(nValorMulta, 2)
        itmX.SubItems(7) = FormatNumber(nValorCorrecao, 2)
        itmX.SubItems(8) = FormatNumber(nValorHon, 2)
        itmX.SubItems(9) = FormatNumber(nValorDil, 2)
        itmX.SubItems(10) = FormatNumber(nValortotal, 2)
    Next
    
    'VALOR DA PARCELA
    If lvDestino.ListItems.Count < 3 Then
        nValorParcela = FormatNumber(CDbl(lvDestino.ListItems(2).SubItems(10)) + nValorExp, 2)
    Else
        nValorParcela = FormatNumber(CDbl(lvDestino.ListItems(3).SubItems(10)) + nValorExp, 2)
    End If
    If nValorParcela < 20 Then Exit For
    Sql = "INSERT SIMULADOREPARC(COMPUTER,QUANTIDADE,VALOR) VALUES('" & NomeDoUsuario & "'," & nQtdeParc & "," & Virg2Ponto(CStr(nValorParcela)) & ")"
    cn.Execute Sql, rdExecDirect
    
Next

Liberado

If frmMdi.frTeste.Visible = True Then
    frmReport.ShowReport "SIMULADOTMP", frmMdi.hwnd, Me.hwnd
Else
    frmReport.ShowReport "SIMULADO", frmMdi.hwnd, Me.hwnd
End If


Sql = "DELETE FROM SIMULADOREPARC WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub
