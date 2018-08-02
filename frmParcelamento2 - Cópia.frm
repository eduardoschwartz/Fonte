VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmParcelamento2 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parcelamento de Divida Fiscal  (Decreto Nº 4.860 de 13 de Novembro de 2006)"
   ClientHeight    =   5415
   ClientLeft      =   2010
   ClientTop       =   1965
   ClientWidth     =   11505
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   11505
   Begin VB.Frame frDetalhe 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2205
      Left            =   2220
      TabIndex        =   59
      Top             =   2160
      Visible         =   0   'False
      Width           =   7575
      Begin MSFlexGridLib.MSFlexGrid grdTributo 
         Height          =   1935
         Left            =   30
         TabIndex        =   60
         Top             =   30
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   3413
         _Version        =   393216
         Rows            =   10
         Cols            =   8
         FixedCols       =   0
         BackColor       =   12582912
         ForeColor       =   65535
         BackColorFixed  =   0
         ForeColorFixed  =   16777215
         BackColorSel    =   65535
         ForeColorSel    =   192
         BackColorBkg    =   12582912
         GridColor       =   16777215
         GridColorFixed  =   12582912
         FocusRect       =   0
         GridLinesFixed  =   0
         SelectionMode   =   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   ">Código |<Nome do Tributo           |>Principal     |>Juros        |>Multa        |>Correção     |>Total       |>%       "
      End
      Begin VB.Label lblPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6720
         TabIndex        =   67
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Geral ---->"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   66
         Top             =   1980
         Width           =   2385
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   5940
         TabIndex        =   65
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4980
         TabIndex        =   64
         Top             =   1980
         Width           =   945
      End
      Begin VB.Label lblM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4140
         TabIndex        =   63
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label lblJ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3300
         TabIndex        =   62
         Top             =   1980
         Width           =   825
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2400
         TabIndex        =   61
         Top             =   1980
         Width           =   885
      End
   End
   Begin VB.OptionButton optGuia 
      Caption         =   "Normal"
      Height          =   195
      Index           =   0
      Left            =   7965
      TabIndex        =   84
      Top             =   855
      Width           =   825
   End
   Begin VB.OptionButton optGuia 
      Caption         =   "Boleto"
      Height          =   195
      Index           =   1
      Left            =   8865
      TabIndex        =   83
      Top             =   855
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.Frame frDDList 
      BackColor       =   &H00EEEEEE&
      Height          =   375
      Left            =   10125
      TabIndex        =   79
      Top             =   90
      Width           =   1230
      Begin VB.ListBox lstAno 
         Height          =   1635
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   81
         Top             =   405
         Width           =   1140
      End
      Begin prjChameleon.chameleonButton cmdDDList 
         Height          =   240
         Left            =   270
         TabIndex        =   80
         ToolTipText     =   "Exibir Lista"
         Top             =   45
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
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":0000
         PICN            =   "frmParcelamento2.frx":001C
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
   Begin VB.CheckBox chkAnistia 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "REFIS-2013"
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
      Height          =   255
      Left            =   10035
      TabIndex        =   77
      Top             =   810
      Width           =   1395
   End
   Begin prjChameleon.chameleonButton cmdAnistia 
      Height          =   345
      Left            =   10920
      TabIndex        =   76
      ToolTipText     =   "Informação do REFIS-I"
      Top             =   1080
      Visible         =   0   'False
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   ""
      ENAB            =   0   'False
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
      MICON           =   "frmParcelamento2.frx":0176
      PICN            =   "frmParcelamento2.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame pnlContaParcela 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   2880
      TabIndex        =   68
      Top             =   1305
      Visible         =   0   'False
      Width           =   6915
      Begin MSFlexGridLib.MSFlexGrid grdTemp 
         Height          =   2730
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   4815
         _Version        =   393216
         Rows            =   20
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   8388608
         ForeColorFixed  =   16777215
         BackColorSel    =   128
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColorFixed  =   16777215
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
         FocusRect       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "^Ano      |^Lanc  |^Seq   |^Parc    |^Compl  |<Processos                                    "
      End
      Begin prjChameleon.chameleonButton cmdOK 
         Height          =   345
         Left            =   5850
         TabIndex        =   70
         ToolTipText     =   "Fechar Aviso"
         Top             =   3360
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&OK"
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
         MICON           =   "frmParcelamento2.frx":02EC
         PICN            =   "frmParcelamento2.frx":0308
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   3795
         Left            =   0
         Top             =   0
         Width           =   6915
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "VEZES E NÃO PODEM MAIS SER PARCELADOS ANTES DE AJUIZAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   72
         Top             =   570
         Width           =   6075
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "O(S) SEGUINTE(S) LANÇAMENTO(S) JÁ FORAM PARCELADOS DUAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   71
         Top             =   270
         Width           =   6195
      End
   End
   Begin MSComctlLib.ListView lvDestino 
      Height          =   3045
      Left            =   30
      TabIndex        =   9
      Top             =   1710
      Visible         =   0   'False
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5371
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Pc"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Vencto."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Vl.Liquido"
         Object.Width           =   1763
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Juros"
         Object.Width           =   1483
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Multa"
         Object.Width           =   1483
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Correção"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Principal"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Saldo"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Jr%"
         Object.Width           =   899
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Jur/Mes"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Jur Apl."
         Object.Width           =   1498
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Hon."
         Object.Width           =   1130
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Total"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Documento"
         Object.Width           =   0
      EndProperty
   End
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
      Left            =   7290
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
   Begin VB.TextBox txtValorEntrada 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
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
      Left            =   8070
      TabIndex        =   6
      Top             =   150
      Value           =   1  'Checked
      Width           =   945
   End
   Begin VB.Frame fr2 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   60
      TabIndex        =   41
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
         TabIndex        =   42
         Top             =   690
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   15658734
         ForeColor       =   12582912
         MouseIcon       =   "frmParcelamento2.frx":0462
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
         MouseIcon       =   "frmParcelamento2.frx":047E
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
         TabIndex        =   51
         Top             =   30
         Width           =   2910
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo.....:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Processo.......:"
         Height          =   225
         Index           =   1
         Left            =   2970
         TabIndex        =   49
         Top             =   690
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do Parcelamento:"
         Height          =   225
         Index           =   2
         Left            =   2970
         TabIndex        =   48
         Top             =   60
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de Parcelas...:"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   1050
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Data do 1º Vencto......:"
         Height          =   225
         Index           =   4
         Left            =   2970
         TabIndex        =   46
         Top             =   1050
         Width           =   1665
      End
      Begin VB.Label lblDataParc 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   4680
         TabIndex        =   45
         Top             =   60
         Width           =   1185
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido...:"
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   44
         Top             =   390
         Width           =   1485
      End
      Begin VB.Label lblNome 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2970
         TabIndex        =   43
         Top             =   330
         Width           =   3195
      End
   End
   Begin MSComctlLib.ListView lvOrigem 
      Height          =   3105
      Left            =   30
      TabIndex        =   15
      Top             =   1710
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   5477
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
         Text            =   "Lançamento"
         Object.Width           =   4586
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
   Begin VB.Frame fr4 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   495
      Left            =   45
      TabIndex        =   20
      Top             =   4860
      Width           =   11355
      Begin prjChameleon.chameleonButton cmdSimulado 
         Height          =   345
         Left            =   2955
         TabIndex        =   12
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":049A
         PICN            =   "frmParcelamento2.frx":04B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDetalhe 
         Height          =   345
         Left            =   7290
         TabIndex        =   58
         ToolTipText     =   "Desmarca todos os lançamentos"
         Top             =   60
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Detalhes"
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
         MICON           =   "frmParcelamento2.frx":087E
         PICN            =   "frmParcelamento2.frx":089A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPrint 
         Height          =   345
         Left            =   2955
         TabIndex        =   55
         ToolTipText     =   "Imprimir o cálculo de parcelamento"
         Top             =   60
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   609
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":09F4
         PICN            =   "frmParcelamento2.frx":0A10
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
         Height          =   345
         Left            =   1500
         TabIndex        =   11
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":0B6A
         PICN            =   "frmParcelamento2.frx":0B86
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
         TabIndex        =   14
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":0E0F
         PICN            =   "frmParcelamento2.frx":0E2B
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
         Left            =   45
         TabIndex        =   10
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":0F85
         PICN            =   "frmParcelamento2.frx":0FA1
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
         TabIndex        =   13
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":11E5
         PICN            =   "frmParcelamento2.frx":1201
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
         TabIndex        =   21
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
         MICON           =   "frmParcelamento2.frx":135B
         PICN            =   "frmParcelamento2.frx":1377
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
         Left            =   45
         TabIndex        =   22
         ToolTipText     =   "Retorna a tela anterior"
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmParcelamento2.frx":151F
         PICN            =   "frmParcelamento2.frx":153B
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
         Left            =   10125
         TabIndex        =   82
         ToolTipText     =   "Sair da Tela"
         Top             =   45
         Width           =   1125
         _ExtentX        =   1984
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
         MICON           =   "frmParcelamento2.frx":1695
         PICN            =   "frmParcelamento2.frx":16B1
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Anos..:"
      Height          =   240
      Index           =   8
      Left            =   9585
      TabIndex        =   78
      Top             =   180
      Width           =   465
   End
   Begin VB.Label lblAnistia 
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
      Left            =   9990
      TabIndex        =   75
      Top             =   1170
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   5
      Left            =   10620
      TabIndex        =   74
      Top             =   1170
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Isenção dos juros e multa conforme REFIS-IV em :"
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   1
      Left            =   6480
      TabIndex        =   73
      Top             =   1170
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Label lblAno 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1650
      TabIndex        =   57
      Top             =   5580
      Width           =   6525
   End
   Begin VB.Label Label1 
      Caption         =   "Exercício(s).:"
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
      Index           =   0
      Left            =   30
      TabIndex        =   56
      Top             =   5580
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sequencia..:"
      Height          =   225
      Left            =   6450
      TabIndex        =   54
      Top             =   870
      Width           =   1065
   End
   Begin VB.Label lblSeq 
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Left            =   7470
      TabIndex        =   53
      Top             =   870
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada.....:"
      Height          =   225
      Index           =   6
      Left            =   6480
      TabIndex        =   52
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
      TabIndex        =   40
      Top             =   6630
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
      Height          =   195
      Index           =   0
      Left            =   6195
      TabIndex        =   39
      Top             =   6870
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
      Height          =   195
      Left            =   8205
      TabIndex        =   38
      Top             =   6855
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
      TabIndex        =   37
      Top             =   6885
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
      TabIndex        =   36
      Top             =   6870
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
      TabIndex        =   35
      Top             =   6885
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
      TabIndex        =   34
      Top             =   6630
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
      TabIndex        =   33
      Top             =   6630
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
      Height          =   195
      Left            =   6195
      TabIndex        =   32
      Top             =   6630
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
      Height          =   195
      Left            =   8205
      TabIndex        =   31
      Top             =   6630
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
      TabIndex        =   30
      Top             =   7140
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
      TabIndex        =   29
      Top             =   7125
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
      TabIndex        =   28
      Top             =   6870
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
      TabIndex        =   27
      Top             =   6630
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
      TabIndex        =   26
      Top             =   7110
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
      TabIndex        =   25
      Top             =   7125
      Width           =   1710
   End
   Begin VB.Label lblValorParcela 
      BackStyle       =   0  'Transparent
      Caption         =   "VARIÁVEL"
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
      Left            =   8355
      TabIndex        =   24
      Top             =   7170
      Width           =   1050
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
      Height          =   195
      Index           =   1
      Left            =   6210
      TabIndex        =   23
      Top             =   7110
      Width           =   2070
   End
   Begin VB.Label lblAnoProc 
      Height          =   315
      Left            =   2130
      TabIndex        =   19
      Top             =   6000
      Width           =   1635
   End
   Begin VB.Label lblNumProc 
      Height          =   315
      Left            =   210
      TabIndex        =   18
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Label lblResp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   60
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   1470
      Width           =   11385
   End
End
Attribute VB_Name = "frmParcelamento2"
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

Private Type Multa
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    bAchou As Boolean
End Type

Private Type Parcela
    nNumProc1 As Long
    nAnoProc1 As Integer
    nNumProc2 As Long
    nAnoProc2 As Integer
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nContador As Integer
End Type


Dim RdoAux As rdoResultset, Sql As String, aTributo() As TRIBUTO, aTributos() As TRIBUTO, nValorDil As Double
Dim sNumProc As String, nNumProc As Long, nAnoProc As Integer, sTipoReparc As String, aLancamento(), nPlano As Integer
Dim bIPTU As Boolean, bISS As Boolean, bVS As Boolean, bTLic As Boolean, bDIV As Boolean, bTCD As Boolean, nValorExp As Double, bMove As Boolean, X1 As Integer, Y1 As Integer
Dim bAjuizado As Boolean

Private Sub chkAnistia_Click()
txtQtdeParc_LostFocus
CarregaDebito Val(txtCod.Text)
End Sub

Private Sub chkCorrecao_Click()
CarregaDebito Val(txtCod.Text)
AtualizaTotal
End Sub

Private Sub chkHon_Click()
AtualizaTotal
End Sub

Private Sub chkJuros_Click()
CarregaDebito Val(txtCod.Text)
AtualizaTotal
End Sub

Private Sub chkMulta_Click()
CarregaDebito Val(txtCod.Text)
AtualizaTotal
End Sub

Private Sub cmdAdd_Click()
Dim x As Integer

For x = 1 To lvOrigem.ListItems.Count
    lvOrigem.ListItems(x).Checked = True
Next
AtualizaTotal
End Sub

Private Sub cmdAnistia_Click()
Dim sTexto As String
sTexto = "- Somente para parcelamentos efetuados até 27/12/2013" & vbCrLf
sTexto = sTexto & "- Todos os débitos parcelados devem estar vencidos até 31/12/2012" & vbCrLf
sTexto = sTexto & "- O valor da parcela mínima é de R$250,00 " & vbCrLf
sTexto = sTexto & "- Até 6 parcelas -> anistia de 95% de juros e multa " & vbCrLf
sTexto = sTexto & "- Até 12 parcelas -> anistia de 90% de juros e multa " & vbCrLf
sTexto = sTexto & "- Até 24 parcelas -> anistia de 80% de juros e multa " & vbCrLf
sTexto = sTexto & "- Até 36 parcelas -> anistia de 70% de juros e multa " & vbCrLf
sTexto = sTexto & "- Até 48 parcelas -> anistia de 60% de juros e multa " & vbCrLf
sTexto = sTexto & "- De 49 à 120 parcelas -> anistia de 50% de juros e multa"

MsgBox sTexto, vbInformation, "Regras para o REFIS-2013"
End Sub

Private Sub cmdCancel_Click()
If Val(txtCod.Text) = 0 Then Exit Sub
If MsgBox("Cancelar a operação ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
Limpar
End Sub

Private Sub cmdDDList_Click()
Dim nAno As Integer, x As Integer, y As Integer

If cmdDDList.Value = True Then
    frDDList.Height = 2130
Else
    For x = 0 To lstAno.ListCount - 1
        nAno = lstAno.List(x)
        If lstAno.Selected(x) = True Then
            For y = 1 To lvOrigem.ListItems.Count
                If lvOrigem.ListItems(y).Text = nAno Then
                    lvOrigem.ListItems(y).Checked = True
                End If
            Next
        Else
            For y = 1 To lvOrigem.ListItems.Count
                If lvOrigem.ListItems(y).Text = nAno Then
                    lvOrigem.ListItems(y).Checked = False
                End If
            Next
        End If
    Next
    AtualizaTotal
    frDDList.Height = 375
End If

End Sub

Private Sub cmdDel_Click()
Dim x As Integer

For x = 1 To lvOrigem.ListItems.Count
    lvOrigem.ListItems(x).Checked = False
Next
AtualizaTotal
End Sub

Private Sub cmdDetalhe_Click()
If frDetalhe.Visible = True Then
    frDetalhe.Visible = False
Else
    AtualizaTributo
    frDetalhe.Visible = True
End If
End Sub

Private Sub cmdGeraDebito_Click()
Dim df As Integer, bAchou As Boolean, x As Integer, bS As Boolean, bN As Boolean, sTCD As Boolean, nPos As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nQtde As Integer, nQtdePago As Integer
Dim bVigS As Boolean, bVigN As Boolean, dDataBase As Date

dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
cmdDetalhe.Value = False
frDetalhe.Visible = False
If lblNome.Caption = "" Then
    MsgBox "Selecione o contribuinte", vbExclamation, "Atenção"
    Exit Sub
End If

'If Not IsDate(mskDataProc.text) Then
'    MsgBox "Data do Processo inválido.", vbCritical, "Atenção"
'    Exit Sub
'End If



If Val(txtQtdeParc.Text) < 2 Or Val(txtQtdeParc) > 120 Then
    MsgBox "A quantidade de parcelas deve ser no mínimo 2 e no máximo 120.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskVencto.Text) Then
    MsgBox "Data de vencimento inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If CDate(Format(mskVencto.Text, "dd/mm/yyyy")) < CDate(Format(dDataBase, "dd/mm/yyyy")) Then
    MsgBox "Data de vencimento menor que a data base", vbExclamation, "Atenção"
    Exit Sub
End If

If Not bAnistia Then
'    If Not ValidaMI Then Exit Sub
End If

'df = ValidaFeriado(CDate(mskVencto.Text))
df = 0 'foi solicitado para que todos os vencimentos do parcelamento sejam no mesmo dia, então não vamos mais verificar feriados

If df = 1 Then
    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVencto.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskVencto.Text = Format(RetornaDiaUtil(CDate(mskVencto.Text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 2 Then
    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVencto.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskVencto.Text = Format(RetornaDiaUtil(CDate(mskVencto.Text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 3 Then
    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
    Sql = Sql & " Where DIA = " & Day(CDate(mskVencto.Text))
    Sql = Sql & " AND MES=" & Month(CDate(mskVencto.Text)) & " AND ANO=" & Year(CDate(mskVencto.Text))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVencto.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                mskVencto.Text = RetornaDiaUtil(CDate(mskVencto.Text))
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

If Not ContaParcela Then Exit Sub

CONTINUA:
If Val(txtValorEntrada.Text) > 0 Then
    MsgBox "Parcelamento com valor de entrada bloqueado.", vbCritical, "Módulo não disponível"
'    If Val(txtValorEntrada.text) <= lblValorParcela Or Val(txtValorEntrada.text) >= lblValorTotal Then
'        MsgBox "O valor da parcela de entrada não pode ser menor ou igual " & vbCrLf & " ao valor de uma parcela, e não pode ser maior que o " & vbCrLf & " valor total do parcelamento", vbExclamation, "Atenção"
        Exit Sub
 '   End If
End If

If CDbl(lblValorTotal.Caption) < 60 Then
    MsgBox "Parcelamento mínimo deve ser de R$60,00 reais.", vbExclamation, "Atenção"
    Exit Sub
End If

bVigS = False: bVigN = False: bAjuizado = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            If Val(Left$(.ListItems(x).SubItems(1), 2)) = 13 Then
                bVigS = True
            Else
                bVigN = True
            End If
            If .ListItems(x).SubItems(6) = "S" Then
                bAjuizado = True
            End If
        End If
    Next
End With
'If bVigN And bVigS Then
'    MsgBox "Não é possivel parcelar débitos de Vigilância Sanitária com outros débitos " & vbCrLf & "no mesmo parcelamento.", vbExclamation, "Atenção"
'    Exit Sub
'End If

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


'*********ANISTIA**********
If bAnistia And chkAnistia.Value = 1 Then
    With lvOrigem
        For x = 1 To .ListItems.Count
            If .ListItems(x).Checked = True And Left(.ListItems(x).SubItems(1), 2) <> 69 Then
                If CDate(.ListItems(x).SubItems(5)) > CDate("31/12/2012") Then
                    MsgBox "Não é possível incluir débitos com vencimento posterior a 31/12/2012 no projeto REFIS-2013", vbCritical, "Atenção"
                    Exit Sub
                End If
            End If
        Next
    End With
End If
'**************************


CarregaLancamento
bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If nLanc = 1 Or nLanc = 29 Then
                bIPTU = True
            ElseIf nLanc = 2 Or nLanc = 3 Or nLanc = 5 Or nLanc = 14 Or nLanc = 69 Or nLanc = 49 Then
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

'If bTCD And (bIPTU Or bISS Or bVS Or bDIV) Then
'    MsgBox "TCD não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If

'If bitpu And (bISS Or bVS Or bDIV) Then
'    MsgBox "IPTU não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If
'If bISS And (bIPTU Or bVS Or bDIV) Then
'    MsgBox "ISS não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If
'If bVS And (bISS Or bIPTU Or bDIV) Then
'    MsgBox "Vigilância Sanitária não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If
'If bDIV And (bISS Or bVS Or bIPTU) Then
'If bDIV And (bISS Or bVS) Then
'    MsgBox "IPTU,ISS e VS não podem ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If

If bTCD Then
    With lvOrigem
        nAno = .ListItems(nPos).Text
        nLanc = Val(Left$(.ListItems(nPos).SubItems(1), 2))
        nSeq = .ListItems(nPos).SubItems(2)
        nParc = .ListItems(nPos).SubItems(3)
        nCompl = .ListItems(nPos).SubItems(4)
        Sql = "SELECT NUMPROCESSO FROM DESTINOREPARC WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND "
        Sql = Sql & "CODLANCAMENTO=" & 20 & " AND NUMSEQUENCIA=" & 0 & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNumProc = !NUMPROCESSO
            Else
                sNumProc = ""
            End If
           .Close
        End With
        
        If sNumProc <> "" Then
            Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND "
            Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If Not IsNull(!CONTADOR) Then
                    nQtdePago = !CONTADOR
                End If
               .Close
            End With
            
            Sql = "SELECT QTDEPARCELA FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    nQtde = !qtdeparcela
                Else
                    nQtde = 0
                End If
               .Close
            End With
            If nQtde > 0 Then
                If Val(txtQtdeParc.Text) > nQtde Then
                    MsgBox "O máximo de parcelas para este parcelamento é de " & nQtde - nQtdePago & ".", vbExclamation, "Atenção"
                    Exit Sub
                End If
            End If
        End If
    End With
End If

Ocupado
DefineParcelas Val(txtQtdeParc.Text)
Liberado

For x = 1 To lvDestino.ListItems.Count - 1
'    If Year(CDate(lvDestino.ListItems(x).SubItems(1))) = 2007 Then
        If CDbl(lvDestino.ListItems(x).ListSubItems(12)) < 20 Then
            MsgBox "Valor da Parcela mínima deve ser de R$25,00 reais." & vbCrLf & "O Valor da parcela é de R$ " & lvDestino.ListItems(x).ListSubItems(12), vbExclamation, "Atenção"
            Exit Sub
        End If
'    End If
Next

TrocaTela
End Sub

Private Sub cmdGrava_Click()
On Error GoTo Erro
Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, dDataBase As Date

dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))

If optGuia(1).Value = True Then
    EmiteBoleto
    Exit Sub
End If

If txtNumProc.Text = "" Then
    MsgBox "Digite o número do Processo.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataProc.Text) Then
    MsgBox "Data do Processo inválido.", vbCritical, "Atenção"
    Exit Sub
End If

'******ANISTIA*********
If bAnistia And chkAnistia.Value = 1 Then
    With lvOrigem
        For x = 1 To .ListItems.Count
            If .ListItems(x).Checked Then
                If CDate(.ListItems(x).SubItems(5)) > CDate("31/12/2012") And Val(Left(.ListItems(x).SubItems(1), 2)) <> 69 Then
                    MsgBox "Existem lancamentos posteriores a 31/12/2012 e não podem ser anistiados.", vbCritical, "Atenção"
                    Exit Sub
                End If
            End If
        Next
    End With

    If CDbl(lvDestino.ListItems(1).SubItems(12)) < 245 Then
        MsgBox "Valor mínimo da parcela é de R$250,00 (R$245,00 + R$5,00 de serviços)", vbCritical, "Atenção"
        Exit Sub
    End If
Else
    nPlano = 0
End If

'*********************

If MsgBox("Deseja Gravar este Parcelamento ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
'Exit Sub
nNumProc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
nAnoProc = Right$(txtNumProc.Text, 4)
lblNumProc.Caption = nNumProc
lblAnoProc.Caption = nAnoProc
sNumProc = CStr(nNumProc) & "/" & CStr(nAnoProc)

'VERIFICA SE O PROCESSO JA FOI UTILIZADO
Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & " '"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este processo já foi utilizado em um parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

'GRAVA O PROCESSO
Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,NUMPROC,ANOPROC,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CODIGORESP,FUNCIONARIO,PLANO,NOVO) VALUES('"
Sql = Sql & sNumProc & "'," & nNumProc & "," & nAnoProc & ",'" & Format(mskDataProc.Text, "mm/dd/yyyy") & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "',"
Sql = Sql & Val(txtQtdeParc.Text) & "," & Virg2Ponto(txtValorEntrada.Text) & "," & 0 & ","
Sql = Sql & IIf(chkMulta.Value = vbChecked, 1, 0) & "," & IIf(chkJuros.Value = vbChecked, 1, 0) & ","
Sql = Sql & Val(txtCod.Text) & ",'" & NomeDeLogin & "'," & nPlano & "," & 1 & ")"
cn.Execute Sql, rdExecDirect

bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If nLanc = 1 Or nLanc = 29 Then
                bIPTU = True
            ElseIf nLanc = 2 Or nLanc = 3 Or nLanc = 5 Or nLanc = 6 Or nLanc = 14 Then
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
    
AtualizaTributo
    
'GRAVA AS PARCELAS DE DESTINO
With lvDestino
    For x = 1 To .ListItems.Count - 1
        sDataVencto = .ListItems(x).SubItems(1)
        nAno = Val(Right$(sDataVencto, 4))
        nLanc = 20
        nSeq = Val(lblSeq.Caption)
        nParc = .ListItems(x).Text
        nCompl = 0
       'GRAVA DESTINOREPARC
        Sql = "INSERT DESTINOREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,VALORLIQUIDO,"
        Sql = Sql & "JUROS,MULTA,CORRECAO,VALORPRINCIPAL,SALDO,JUROSPERC,JUROSVALOR,JUROSAPL,HONORARIO,TOTAL) VALUES('"
        Sql = Sql & sNumProc & "'," & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(2)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(3)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(4)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(5)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(6)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(7)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(8)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(9)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(10)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(11)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(12)))) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA DEBITOPARCELA
        If nAno = Year(dDataBase) Then
            nSituacao = 3
        Else
            nSituacao = 18
        End If
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
        Sql = Sql & nSituacao & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
        cn.Execute Sql, rdExecDirect
       
       'DEFINE TRIBUTO PRINCIPAL
        With grdTributo
            For y = 1 To .Rows - 1
                nValorTributo = CDbl(lvDestino.ListItems(x).SubItems(2))
                nCodTrib = Val(.TextMatrix(y, 0))
                nPerc = CDbl(.TextMatrix(y, 7))
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
                Sql = Sql & nParc & "," & nCompl & "," & nCodTrib & "," & Virg2Ponto(CStr(nValorTributo * nPerc / 100)) & ")"
                cn.Execute Sql, rdExecDirect
            Next
        End With
       
       '***********ANISTIA*********
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 609 - TX.PARCEL.REFIS)
        If Val(lblAnistia.Caption) > 0 And bAnistia Then
            nValorExp = 5
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 609 & "," & Virg2Ponto(CStr(nValorExp)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       '**************************
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 113 - JUROS)
        nValorTributo = CDbl(.ListItems(x).SubItems(3))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 113 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 112 - MULTA)
        nValorTributo = CDbl(.ListItems(x).SubItems(4))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 112 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 26 - CORREÇÃO)
        nValorTributo = CDbl(.ListItems(x).SubItems(5))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 26 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 585 - JUROS APLICADO)
        nValorTributo = CDbl(.ListItems(x).SubItems(10))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 585 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
        If chkHon.Value = 1 And CDbl(.ListItems(x).SubItems(11)) > 0 Then
            'GRAVA DEBITOTRIBUTO   // (TRIBUTO 90 - HONORÁRIOS)
             nValorTributo = CDbl(.ListItems(x).SubItems(11))
             Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
             Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
             Sql = Sql & nParc & "," & nCompl & "," & 90 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
             cn.Execute Sql, rdExecDirect
        End If
        'RETORNA ULTIMO DOCUMENTO
        Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nLastCod = !MAXIMO + 1
           .Close
        End With
       'GRAVA NUMDOCUMENTO
        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO) VALUES("
        Sql = Sql & nLastCod & ",'" & Format(dDataBase, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorExp)) & "," & IIf(Val(lblAnistia.Caption) > 0, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.Text) & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
        Sql = Sql & nCompl & "," & nLastCod & ")"
        cn.Execute Sql, rdExecDirect
       'PREENCHE DOCUMENTO NO LVDESTINO
        .ListItems(x).SubItems(13) = nLastCod
    Next
End With

'GRAVA AS PARCELAS DE ORIGEM
With lvOrigem
    For x = 1 To .ListItems.Count
        If lvOrigem.ListItems(x).Checked = True Then
             nAno = .ListItems(x).Text
             nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
             nSeq = .ListItems(x).SubItems(2)
             nParc = .ListItems(x).SubItems(3)
             nCompl = .ListItems(x).SubItems(4)
            'GRAVA ORIGEMREPARC
             Sql = "INSERT ORIGEMREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,PRINCIPAL,JUROS,MULTA,CORRECAO) VALUES('" & sNumProc & "'," & Val(txtCod.Text) & ","
             Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(7)))) & ","
             Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(8)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(9)))) & ","
             Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(10)))) & ")"
             cn.Execute Sql, rdExecDirect
            'ATUALIZA O STATUS DE ORIGEM   // (4 - REPARCELADO)
             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=4 WHERE CODREDUZIDO=" & Val(txtCod.Text)
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
dDataProc = CDate(mskDataProc.Text)
sDescImposto = "REPARCELAMENTO"
NumBarra1 = Format(ExtraiNumero(txtNumProc.Text), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)

'ENDEREÇO DO CONTRIBUINTE
Select Case Val(txtCod.Text)
    Case 1 To 99999
        'DADOS DO IMOVEL
        Sql = "SELECT * FROM vwCnsImovel WHERE CODREDUZIDO=" & Val(txtCod.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNumInsc = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
                sCodReduz = Format(Val(txtCod.Text), "000000") & "-" & RetornaDVCodReduzido(Val(txtCod.Text))
                Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
                Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
                Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & Val(txtCod.Text) & " AND TIPOPROP='P' AND PRINCIPAL=1"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                sNomeResp = RdoAux2!nomecidadao
                RdoAux2.Close
                sTipoImposto = "REPARCEL."
                sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                nNumImovel = !Li_Num
                sComplImovel = IIf(IsNull(!Li_Compl) Or !Li_Compl = "", " ", !Li_Compl)
                If !CodBairro <> 999 Then
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
               Sql = Sql & "WHERE CODREDUZIDO=" & Val(txtCod.Text)
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
        Sql = "SELECT * "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & Val(txtCod.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNumInsc = SubNull(!INSCESTADUAL)
                sCodReduz = !codigomob
                sNomeResp = !razaosocial
                sTipoImposto = "REPARCEL."
                If IsNull(!NomeLogradouro) Then
                    sEndImovel = !NomeLogr
                Else
                    sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                End If
                nNumImovel = Val(SubNull(!Numero))
                sComplImovel = SubNull(!Complemento)
                If !CodCidade <> 413 Then
                    sBairroImovel = !DescBairro
                    sCidadeEntrega = !desccidade
                End If
                GoTo fim
                If !CodBairro <> 999 Then
                     Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO=" & !CodBairro
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
                Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.Text)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sEndEntrega = SubNull(!NomeLogradouro)
                        nNumEntrega = SubNull(!numimovel)
                        sBairroEntrega = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        sComplEntrega = SubNull(!Complemento)
                        sCepEntrega = SubNull(!Cep)
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
fim:
    Case 500000 To 800000
        Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao,cidadao.cpf, cidadao.cnpj, cidadao.rg, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, "
        Sql = Sql & "cidadao.siglauf, cidade.desccidade, bairro.descbairro, cidadao.nomelogradouro AS nomerua "
        Sql = Sql & "cidadao.codlogradouro , vwLOGRADOURO.AbrevTipoLog, vwLOGRADOURO.AbrevTitLog, vwLOGRADOURO.NomeLogradouro "
        Sql = Sql & "FROM  vwLOGRADOURO RIGHT JOIN  cidadao ON vwLOGRADOURO.CODLOGRADOURO = cidadao.codlogradouro LEFT OUTER JOIN "
        Sql = Sql & "cidade INNER JOIN  bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade And cidadao.codbairro = bairro.codbairro "
        Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & Val(txtCod.Text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sTipoImposto = "REPARCEL."
                sCodReduz = !CodCidadao
                sNomeResp = !nomecidadao
                If Not IsNull(!NomeLogradouro) Then
                    sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                Else
                    sEndImovel = SubNull(!nomerua)
                End If
                nNumImovel = SubNull(!numimovel)
                sComplImovel = SubNull(!Complemento)
                nCodBairro = Val(SubNull(!CodBairro))
                sCidadeEntrega = SubNull(!desccidade)
                sUFEntrega = SubNull(!SiglaUF)
                If Not IsNull(!CPF) Then
                   sCPF = !CPF
                ElseIf Not IsNull(!Cnpj) Then
                   sCPF = !Cnpj
                ElseIf Not IsNull(!rg) Then
                   sCPF = !rg
                Else
                   sCPF = ""
                End If
             Else
                sCPF = ""
             End If
        '     If nCodBairro = 0 Then
        '        sBairroImovel = SubNull(!NOMEBairro)
         '       sCidadeEntrega = SubNull(!NomeCidade)
          '   End If
        End With
End Select

'GRAVA TEMPORARIO
With lvDestino
    For x = 1 To .ListItems.Count - 1
        sDataVencto = .ListItems(x).SubItems(1)
        nAno = Val(Right$(sDataVencto, 4))
        nLanc = 20
        nSeq = Val(lblSeq.Caption)
        nParc = .ListItems(x).Text
        nCompl = 0
'        nAno = .ListItems(x).text
'        nSeq = lblSeq.Caption
'        nParc = .ListItems(x).SubItems(2)
        nNumDoc = .ListItems(x).SubItems(13)
'        sDataVencto = .ListItems(x).SubItems(3)
        If bAnistia And chkAnistia.Value = vbChecked Then
            sValorParc = CStr(CDbl(.ListItems(x).SubItems(12)) + 5) '5=Taxa de expediente anistia
            nValorExp = 0
        Else
            sValorParc = .ListItems(x).SubItems(12)
            If Val(Right$(sDataVencto, 4)) = 2006 Then
                nValorExp = 1.13
            Else
                Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE CODTRIBUTO=3 AND ANO=" & Year(Now)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                nValorExp = RdoAux2!VALORALIQ
                RdoAux2.Close
'                nValorExp = 1.6
            End If
        End If
        
        
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
        Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Left$(sDescImposto, 30) & "'," & nAno & ",'" & Left$(txtNumProc.Text, 25) & "','"
        Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "'," & nNumDoc & "," & RetornaDVNumDoc(nNumDoc) & ",'" & sQuadra & "','"
        Sql = Sql & sLote & "','" & Format(sDataVencto, "mm/dd/yyyy") & "'," & IIf(nParc = 0, 1, nParc) & "," & Val(txtQtdeParc.Text) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
        Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & "REPARCELAMENTO" & "'," & Virg2Ponto(CDbl(nValorExp)) & "," & "0" & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

frmConfissaoDivida.txtNumProc.Text = sNumProc
frmConfissaoDivida.txtNumProc.Locked = True
frmConfissaoDivida.lblSid.Caption = "0"
frmConfissaoDivida.CarregaProcesso
cmdPrint_Click
'cmdVoltar_Click
'Limpar
Ocupado
frmConfissaoDivida.show
frmConfissaoDivida.ZOrder 0
Liberado
cmdVoltar_Click
Limpar
'Unload frmParcelamento2
Exit Sub

Erro:

MsgBox Err.Description
Resume Next

End Sub


Private Sub cmdOK_Click()
pnlContaParcela.Visible = False
fr4.Enabled = True
End Sub

Private Sub cmdPrint_Click()

Dim nSomaLiquido As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double, nSomaPrincipal As Double, nValorPrincipal As Double
Dim nNumParcela As Integer, dVencimento As Date, nJuros As Double, nValorParcela As Double, nSaldo As Double, nJurosMesPerc As Double
Dim nJurosMesValor As Double, nValorHonorario As Double, nValorTotal As Double, x As Integer
frmConfissaoDivida.Hide
Sql = "DELETE FROM CALCULOPARCELAMENTO WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

With lvDestino
    nSomaLiquido = CDbl(.ListItems(.ListItems.Count).SubItems(2))
    nSomaJuros = CDbl(.ListItems(.ListItems.Count).SubItems(3))
    nSomaMulta = CDbl(.ListItems(.ListItems.Count).SubItems(4))
    nSomaCorrecao = CDbl(.ListItems(.ListItems.Count).SubItems(5))
    nSomaPrincipal = CDbl(.ListItems(.ListItems.Count).SubItems(6))
    nValorPrincipal = nSomaPrincipal / Val(txtQtdeParc.Text)
    For x = 1 To lvDestino.ListItems.Count - 1
        Sql = "INSERT CALCULOPARCELAMENTO(COMPUTER,CODREDUZIDO,NOME,PROCESSO,DATAPROCESSO,QTDEPARCELA,SOMALIQUIDO,SOMAJUROS,SOMAMULTA,SOMACORRECAO,SOMAPRINCIPAL,"
        Sql = Sql & "VALORPRINCIPAL,NUMPARCELA,VENCIMENTO,JUROS,VALORPARCELA,SALDO,JUROSMESPERC,JUROSMESVALOR,VALORHONORARIO,VALORTOTAL) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & Val(txtCod.Text) & ",'" & Mask(lblNome.Caption) & "','" & txtNumProc.Text & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "',"
        Sql = Sql & Val(txtQtdeParc.Text) & "," & Virg2Ponto(RemovePonto(CStr(nSomaLiquido))) & "," & Virg2Ponto(RemovePonto(CStr(nSomaJuros))) & "," & Virg2Ponto(RemovePonto(CStr(nSomaMulta))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(nSomaCorrecao))) & "," & Virg2Ponto(RemovePonto(CStr(nSomaPrincipal))) & "," & Virg2Ponto(RemovePonto(CStr(nValorPrincipal))) & "," & x & ",'"
        Sql = Sql & Format(.ListItems(x).SubItems(1), "mm/dd/yyyy") & "'," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(10)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(6)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(7)))) & "," & Virg2Ponto(CStr(RemovePonto(.ListItems(x).SubItems(8)))) & "," & Virg2Ponto(CStr(RemovePonto(.ListItems(x).SubItems(9)))) & ","
        Sql = Sql & Virg2Ponto(CStr(RemovePonto(.ListItems(x).SubItems(11)))) & "," & Virg2Ponto(CStr(RemovePonto(.ListItems(x).SubItems(12)))) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

If frmMdi.frTeste.Visible = True Then
    If frmMdi.frTeste.Caption = "ACESSANDO OS DADOS LOCAIS" Then
        frmReport.ShowReport "CALCULOPARCELAMENTO", frmMdi.hwnd, Me.hwnd
    Else
        frmReport.ShowReport "CALCULOPARCELAMENTOTMP", frmMdi.hwnd, Me.hwnd
    End If
Else
    frmReport.ShowReport "CALCULOPARCELAMENTO", frmMdi.hwnd, Me.hwnd
End If

Sql = "DELETE FROM CALCULOPARCELAMENTO WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

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
Dim nAno() As Integer, sAno As String


ReDim nAno(0)
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = True Then
        bAchou = False
        For y = 1 To UBound(nAno)
            If nAno(y) = Val(lvOrigem.ListItems(x).Text) Then
            bAchou = True
            End If
        Next
        If Not bAchou Then
            ReDim Preserve nAno(UBound(nAno) + 1)
            nAno(UBound(nAno)) = Val(lvOrigem.ListItems(x).Text)
        End If
    End If
Next
sAno = ""
For x = 1 To UBound(nAno)
    sAno = sAno & CStr(nAno(x)) & ", "
Next
If sAno = "" Then Exit Sub
sAno = Left$(sAno, Len(sAno) - 2)
lblAno.Caption = sAno

'lvOrigem.SetFocus
If Not ContaParcela Then Exit Sub

Simulado

End Sub

Private Sub cmdVoltar_Click()
lblAnistia.Caption = "0,00"
TrocaTela
End Sub


Private Sub Form_Load()
Dim dDataBase As Date
Centraliza Me
dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
lblDataParc.Caption = Format(dDataBase, "dd/mm/yyyy")
bExec = True
mskVencto.Text = Format(dDataBase, "dd/mm/yyyy")
'BUSCA O VALOR DA TAXA DE EXPEDIENTE
Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorExp = FormatNumber(!VALORPARCELA, 2)
   .Close
End With
pnlContaParcela.Visible = False

If Not bAnistia Then
    chkAnistia.Visible = False
    cmdAnistia.Visible = False
Else
    chkAnistia.Visible = True
    cmdAnistia.Visible = True
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If cmdSair.Enabled = False Then
    Cancel = 1
End If

End Sub

Private Sub frDetalhe_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
X1 = x
Y1 = y
bMove = True
End Sub

Private Sub frDetalhe_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If bMove Then
    frDetalhe.Top = frDetalhe.Top + y - Y1
    frDetalhe.Left = frDetalhe.Left + x - X1
End If
End Sub

Private Sub frDetalhe_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
bMove = False
End Sub

Private Sub grdTributo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
X1 = x
Y1 = y
bMove = True
End Sub

Private Sub grdTributo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If bMove Then
    frDetalhe.Top = frDetalhe.Top + y - Y1
    frDetalhe.Left = frDetalhe.Left + x - X1
End If
End Sub

Private Sub grdTributo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
bMove = False
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
mskVencto.SelLength = Len(mskVencto.Text)
End Sub

Private Sub txtCod_Change()
lvOrigem.ListItems.Clear
lblNome.Caption = ""
End Sub

Private Sub txtCod_GotFocus()
On Error Resume Next
txtCod.SetFocus
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_LostFocus()
Dim nCodReduz As Long, sTipoCod As String

If Val(txtCod.Text) = 0 Then Exit Sub
If Val(txtCod.Text) = 0 Then
    lblNome.Caption = ""
    Exit Sub
End If
If Val(txtCod.Text) < 100000 Then
    sTipoCod = "I"
ElseIf Val(txtCod.Text) >= 100000 And Val(txtCod.Text) < 500000 Then
    sTipoCod = "M"
ElseIf Val(txtCod.Text) >= 500000 Then
    sTipoCod = "C"
End If
txtCod.Text = Format(txtCod.Text, "000000")
nCodReduz = Val(txtCod.Text)
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
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux.RowCount > 0 Then
         If sTipoCod = "I" Or sTipoCod = "C" Then
            lblNome.Caption = !nomecidadao
         ElseIf sTipoCod = "M" Then
            lblNome.Caption = !razaosocial
         End If
    Else
       MsgBox "Código não Cadastrado.", vbExclamation, "Atenção"
       lvOrigem.ListItems.Clear
       On Error Resume Next
       txtCod.SetFocus
       Exit Sub
    End If
    .Close
End With

CarregaDebito (nCodReduz)
If lvOrigem.ListItems.Count = 0 Then
    MsgBox "O contribuinte não possue débitos a serem parcelados.", vbExclamation, "Atenção"
    txtCod.Text = 0
    On Error Resume Next
    txtCod.SetFocus
Else
    txtCod.Locked = True
    txtCod.BackColor = Kde
End If

End Sub

Private Sub CarregaDebito(nCodReduz As Long)

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double, t As Integer, Achou As Boolean
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date, nPerc As Double
Dim nSomaValorTributo As Double, sAj As String
Dim x As Integer
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, sDescLanc As String
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
qd(7) = 999 'sequencia
qd(8) = 1
qd(9) = 999 'numparcela
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
'        If !AnoExercicio = 2010 Then MsgBox "TESTE"

        'If Not IsNull(!DATAINSCRICAO) Then GoTo CONTINUA
        If !DataVencimento = CDate(Format(Now, "dd/mm/yyyy")) And !CodLancamento <> 65 And !CodLancamento <> 50 And !CodLancamento <> 62 And !CodLancamento <> 47 And !CodLancamento <> 38 And !CodLancamento <> 69 And !CodLancamento <> 16 Then GoTo Proximo
'        If !DataVencimento = CDate(Format(Now, "dd/mm/yyyy")) And IsNull(!DATAINSCRICAO) And !CodLancamento <> 65 And !CodLancamento <> 50 And !CodLancamento <> 62 And !CodLancamento <> 47 And !CodLancamento <> 38 Then GoTo PROXIMO
        If !CodLancamento = 16 Or !CodLancamento = 69 Or !CodLancamento = 50 Or !CodLancamento = 65 Or !CodLancamento = 62 Or !CodLancamento = 47 Or !CodLancamento = 38 Or !CodLancamento = 72 Then GoTo CONTINUA
        If !DataVencimento > CDate(Format(Now, "dd/mm/yyyy")) Then
            If !CodLancamento <> 69 And !CodLancamento <> 10 And !CodLancamento <> 48 Then
                GoTo Proximo
            End If
        End If
        If !CodLancamento = 20 Then GoTo Proximo
        'If !CodLancamento = 41 Then GoTo PROXIMO
        
CONTINUA:
        If !DataVencimento > CDate(Format(Now, "dd/mm/yyyy")) Then
            If !CodLancamento = 48 Then
                GoTo Proximo
            End If
        End If
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
           aDebito(nEval).nSituacao = !statuslanc
           aDebito(nEval).sSituacao = !Situacao
           aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
           aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
           If chkJuros.Value = 1 Then
                aDebito(nEval).nValorJuros = FormatNumber(!ValorJuros, 2)
           Else
                aDebito(nEval).nValorJuros = FormatNumber(0, 2)
           End If
           If chkMulta.Value = 1 Then
                aDebito(nEval).nValorMulta = FormatNumber(!ValorMulta, 2)
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
                aDebito(x).nValorJuros = aDebito(x).nValorJuros + !ValorJuros
           End If
           If chkMulta.Value = 1 Then
                aDebito(x).nValorMulta = aDebito(x).nValorMulta + !ValorMulta
           End If
           If chkCorrecao.Value = 1 Then
                aDebito(x).nValorCorrecao = aDebito(x).nValorCorrecao + !VALORCORRECAO
           End If
           aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
        End If
Proximo:
       .MoveNext
    Loop
   .Close
End With

'**********ANISTIA*********
If bAnistia And chkAnistia.Value = 1 Then
    nPerc = 100 - CDbl(lblAnistia.Caption)
    For x = 1 To UBound(aDebito)
        With aDebito(x)
            .nValorMulta = FormatNumber(CDbl(.nValorMulta) * nPerc / 100, 2)
            .nValorJuros = FormatNumber(CDbl(.nValorJuros) * nPerc / 100, 2)
            .nValorAtual = FormatNumber(nValorJuros + nValorMulta + nValorCorrecao + nValorLanc, 2)
        End With
    Next
End If
'**************************


For x = 1 To UBound(aDebito)
    With aDebito(x)
        Achou = False
        For t = 0 To lstAno.ListCount - 1
            If lstAno.List(t) = .nAno Then
                Achou = True
                Exit For
            End If
        Next
        If Not Achou Then
            lstAno.AddItem .nAno
        End If
        
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
If Trim$(txtNumProc.Text) <> "" Then
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumProc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
        nAnoProc = Right$(txtNumProc.Text, 4)
        lblNumProc.Caption = nNumProc
        lblAnoProc.Caption = nAnoProc
        sNumProc = CStr(nNumProc) & "/" & CStr(nAnoProc)
        sValidaProc = ValidaProcesso(sNumProc)
        If sValidaProc <> "OK" Then
            MsgBox sValidaProc, vbCritical, "Atenção"
            cmdVoltar_Click
            Exit Sub
        Else
            mskDataProc.Text = Format(RetornaDataProcesso(nNumProc, nAnoProc), "dd/mm/yyyy")
        End If
    Else
        MsgBox "Processo inválido.", vbExclamation, "Atenção"
        txtNumProc.SetFocus
    End If
End If
End Sub

Private Sub txtQtdeParc_GotFocus()
On Error Resume Next
txtQtdeParc.SetFocus
txtQtdeParc.SelStart = 0
txtQtdeParc.SelLength = Len(txtQtdeParc.Text)
End Sub

Private Sub txtQtdeParc_KeyPress(KeyAscii As Integer)
Tweak txtQtdeParc, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtdeParc_LostFocus()
If Trim$(txtQtdeParc.Text) = "" Then txtQtdeParc.Text = "0"

If Val(txtQtdeParc.Text) > 0 And bAnistia And chkAnistia.Value = 1 Then
    If txtQtdeParc.Text <= 6 Then
        lblAnistia.Caption = "95,00"
        nPlano = 4
    ElseIf txtQtdeParc.Text > 6 And txtQtdeParc.Text <= 12 Then
        lblAnistia.Caption = "90,00"
        nPlano = 5
    ElseIf txtQtdeParc.Text > 12 And txtQtdeParc.Text <= 24 Then
        lblAnistia.Caption = "80,00"
        nPlano = 6
    ElseIf txtQtdeParc.Text > 24 And txtQtdeParc.Text <= 36 Then
        lblAnistia.Caption = "70,00"
        nPlano = 7
    ElseIf txtQtdeParc.Text > 36 And txtQtdeParc.Text <= 48 Then
        lblAnistia.Caption = "60,00"
        nPlano = 8
'    ElseIf txtQtdeParc.Text > 48 And txtQtdeParc.Text <= 120 Then
    ElseIf txtQtdeParc.Text > 48 Then
        lblAnistia.Caption = "50,00"
        nPlano = 9
    End If
    CarregaDebito Val(txtCod.Text)
Else
    lblAnistia.Caption = "0,00"
End If

End Sub

Private Sub txtValorEntrada_GotFocus()
txtValorEntrada.SetFocus
txtValorEntrada.SelStart = 0
txtValorEntrada.SelLength = Len(txtValorEntrada.Text)
End Sub

Private Sub txtValorEntrada_KeyPress(KeyAscii As Integer)
Tweak txtValorEntrada, KeyAscii, DecimalPositive
End Sub

Private Sub AtualizaTotal()
On Error Resume Next
Dim x As Integer, y As Integer, nContaParcela As Integer, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double
Dim nValorCorrecao As Double, nValorTotal As Double, nValorAjuizado As Double, bAchou As Boolean
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, RdoAux3 As rdoResultset, nTotal As Double
Dim qd As New rdoQuery

LimpaContador
If lvOrigem.ListItems.Count = 0 Then Exit Sub

nContaParcela = 0: nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorTotal = 0: nValorAjuizado = 0

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
            nValorTotal = nValorTotal + CDbl(.ListItems(x).SubItems(11))
        End If
    Next
End With
lblNumParc.Caption = Format(nContaParcela, "000")
lblValorCorrecao.Caption = FormatNumber(nValorCorrecao, 2)
lblValorJuros.Caption = FormatNumber(nValorJuros, 2)
lblValorMulta.Caption = FormatNumber(nValorMulta, 2)
lblValorPrincipal.Caption = FormatNumber(nValorPrincipal, 2)
lblValorHon.Caption = FormatNumber(nValorAjuizado * 0.1, 2)
lblValorTotal.Caption = FormatNumber(nValorTotal, 2)




End Sub

Private Sub LimpaContador()
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
    cmdSimulado.Visible = False
    cmdPrint.Visible = True
    txtNumProc.Locked = True
    txtNumProc.BackColor = Kde
    txtQtdeParc.Locked = True
    txtQtdeParc.BackColor = Kde
    mskVencto.Locked = True
    mskVencto.BackColor = Kde
    txtValorEntrada.Locked = True
    txtValorEntrada.BackColor = Kde
    chkMulta.Enabled = False
    chkJuros.Enabled = False
    chkCorrecao.Enabled = False
    chkHon.Enabled = False
    lvOrigem.Visible = False
    lvDestino.Visible = True
    lblTitulo.Caption = " Débitos que serão gerados no parcelamento"
    cmdGrava.Visible = True
    cmdCancel.Visible = False
'    lvDestino.SetFocus
Else
    cmdGeraDebito.Visible = True
    cmdVoltar.Visible = False
    cmdAdd.Enabled = True
    cmdDel.Enabled = True
    cmdSair.Enabled = True
    cmdSimulado.Visible = True
    cmdPrint.Visible = False
    txtNumProc.Locked = False
    txtNumProc.BackColor = Branco
    txtQtdeParc.Locked = False
    txtQtdeParc.BackColor = Branco
    mskVencto.Locked = False
    mskVencto.BackColor = Branco
    txtValorEntrada.Locked = False
    txtValorEntrada.BackColor = Branco
    chkMulta.Enabled = True
    chkJuros.Enabled = True
    chkCorrecao.Enabled = True
    chkHon.Enabled = True
    lvOrigem.Visible = True
    lvDestino.Visible = False
    lblTitulo.Caption = " Débitos disponíveis para parcelamento"
    cmdGrava.Visible = False
    cmdCancel.Visible = True
'    cmdGeraDebito.SetFocus
End If

End Sub

Private Sub DefineParcelas(nQtdPrc As Integer)
On Error GoTo Erro
Dim x As Integer, nSeq As Integer, a As Integer, sVencimento As String, sVencimento2 As String, y As Integer, z As Integer, nValorEntrada As Double
Dim nValorPrimeira As Double, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double, nValorHon As Double, nValorDil As Double
Dim nValorPrincipal1 As Double, nValorJuros1 As Double, nValorMulta1 As Double, nValorCorrecao1 As Double, nValorTotal1 As Double, nValorHon1 As Double
Dim nPrincipal As Double, nJuros As Double, nMulta As Double, nCorrecao As Double, nTotal As Double, nHonorario As Double, nDiligencia As Double, nItem As Integer, nSomaJurosValor As Double
Dim nDia As Integer, nMes As Integer, nAno As Integer, itmX As ListItem, nQtdeParc As Integer, nDif As Double, nPerc As Double, nSaldo As Double, nJurosMesPerc As Double, nJurosMesValor As Double
Dim nDiaFixo As Integer, dDataBase As Date
'LIMPA TELA
lvDestino.ListItems.Clear

dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
'BUSCA ULTIMA SEQUENCIA
Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If !CONTADOR > 0 Then
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO<100"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nSeq = RdoAux!SEQMAXIMA + 1
    Else
        nSeq = 0
    End If
   .Close
End With

'ADICIONA ITENS
nQtdeParc = nQtdPrc
If IsDate(mskVencto.Text) Then
    sVencimento = mskVencto.Text
Else
    sVencimento = Format(Now, "dd/mm/yyyy")
End If
If txtValorEntrada.Text = "" Then txtValorEntrada.Text = "0"
'CALCULA VALORES

AtualizaTotal
nValorEntrada = CDbl(txtValorEntrada.Text)
nPrincipal = CDbl(lblValorPrincipal.Caption)
nJuros = CDbl(lblValorJuros.Caption)
nMulta = CDbl(lblValorMulta.Caption)
nCorrecao = CDbl(lblValorCorrecao.Caption)
nHonorario = CDbl(lblValorHon.Caption)
nTotal = CDbl(lblValorTotal.Caption)

If nValorEntrada > 0 Then
    nItem = 2
    nPerc = nValorEntrada / nTotal
    nValorPrincipal1 = FormatNumber(nPrincipal * nPerc, 2)
    nValorJuros1 = FormatNumber(nJuros * nPerc, 2)
    nValorMulta1 = FormatNumber(nMulta * nPerc, 2)
    nValorCorrecao1 = FormatNumber(nCorrecao * nPerc, 2)
    nValorHon1 = FormatNumber(nHonorario * nPerc, 2)
    nValorPrincipal = (nPrincipal - nValorPrincipal1) / (nQtdeParc - 1)
    nValorJuros = (nJuros - nValorJuros1) / (nQtdeParc - 1)
    nValorMulta = (nMulta - nValorMulta1) / (nQtdeParc - 1)
    nValorCorrecao = (nCorrecao - nValorCorrecao1) / (nQtdeParc - 1)
    nValorHon = (nHonorario - nValorHon1) / (nQtdeParc - 1)
Else
    nItem = 1
    nValorPrincipal = nPrincipal / nQtdeParc
    nValorJuros = nJuros / nQtdeParc
    nValorMulta = nMulta / nQtdeParc
    nValorCorrecao = nCorrecao / nQtdeParc
    nValorHon = nHonorario / nQtdeParc
    nValorPrincipal1 = nValorPrincipal
    nValorJuros1 = nValorJuros
    nValorMulta1 = nValorMulta
    nValorCorrecao1 = nValorCorrecao
    nValorHon1 = nValorHon
End If
'nValortotal = nValorPrincipal + nValorHon + nValorJuros + nValorMulta + nValorCorrecao
'nValorTotal1 = nValorPrincipal1 + nValorHon1 + nValorJuros1 + nValorMulta1 + nValorCorrecao1
nValorTotal = nValorPrincipal + nValorJuros + nValorMulta + nValorCorrecao
nValorTotal1 = nValorPrincipal1 + nValorJuros1 + nValorMulta1 + nValorCorrecao1
nSaldo = CDbl(lblValorTotal.Caption)
'nSaldo = lvDestino.ListItems(lvDestino.ListItems.Count).text
For x = 1 To nQtdeParc
    'CALCULA SALDO
    If nValorEntrada > 0 Then
        nSaldo = nSaldo - nValorTotal1
    Else
        nSaldo = nSaldo - nValorTotal
    End If
    'CALCULA VENCIMENTO
     If x > 1 Then
       'nDia = Val(Left$(sVencimento, 2))
       nDia = nDiaFixo
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
       nDiaFixo = Val(Left$(sVencimento, 2))
       nAno = Val(Right$(sVencimento, 4))
    End If
    'sVencimento2 = RetornaDiaUtil(CDate(sVencimento))
    sVencimento2 = sVencimento
    'PERCENTUAL DE JUROS/MES
    If x = 1 Then
        nJurosMesPerc = 0
    Else
        nJurosMesPerc = 1
    End If
    
    'PREENCHE A LISTA
    lblSeq = Format(nSeq, "00")
    Set itmX = lvDestino.ListItems.Add(, , "  " & Format(x, "00"))
    itmX.SubItems(1) = Format(sVencimento2, "dd/mm/yyyy")
    If x = 1 Then
        itmX.SubItems(2) = FormatNumber(nValorPrincipal1, 2)
        itmX.SubItems(3) = FormatNumber(nValorJuros1, 2)
        itmX.SubItems(4) = FormatNumber(nValorMulta1, 2)
        itmX.SubItems(5) = FormatNumber(nValorCorrecao1, 2)
        itmX.SubItems(6) = FormatNumber(nValorTotal1, 2)
        itmX.SubItems(7) = FormatNumber(nSaldo, 2)
        itmX.SubItems(8) = FormatNumber(nJurosMesPerc, 2)
        itmX.SubItems(11) = FormatNumber(nValorHon1, 2)
        itmX.SubItems(12) = FormatNumber(0, 2)
    Else
        itmX.SubItems(2) = FormatNumber(nValorPrincipal, 2)
        itmX.SubItems(3) = FormatNumber(nValorJuros, 2)
        itmX.SubItems(4) = FormatNumber(nValorMulta, 2)
        itmX.SubItems(5) = FormatNumber(nValorCorrecao, 2)
        itmX.SubItems(6) = FormatNumber(nValorTotal, 2)
        itmX.SubItems(7) = FormatNumber(nSaldo, 2)
        itmX.SubItems(8) = FormatNumber(nJurosMesPerc, 2)
        itmX.SubItems(11) = FormatNumber(nValorHon, 2)
        itmX.SubItems(12) = FormatNumber(0, 2)
    End If
    If nAno > Year(dDataBase) Then
        For y = 1 To 12
            itmX.ForeColor = vbRed
            itmX.ListSubItems(y).ForeColor = vbRed
        Next y
    End If

    itmX.Checked = True
Next

'PREENCHE A COLUNA JUROS DO MES/VALOR
nSomaJurosValor = 0: nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorPrincipal1 = 0: nSaldo = 0
nValorHon = 0: nTotal = 0

With lvDestino
    For x = 1 To .ListItems.Count
        nJurosMesPerc = .ListItems(x).ListSubItems(8)
        If x = 1 Then
            nSaldo = CDbl(lblValorTotal.Caption)
        Else
            nSaldo = CDbl(.ListItems(x - 1).ListSubItems(7))
        End If
        .ListItems(x).ListSubItems(9) = FormatNumber(nJurosMesPerc * nSaldo / 100, 2)
        nSomaJurosValor = nSomaJurosValor + CDbl(.ListItems(x).ListSubItems(9))
        nValorPrincipal = nValorPrincipal + CDbl(.ListItems(x).ListSubItems(2))
        nValorJuros = nValorJuros + CDbl(.ListItems(x).ListSubItems(3))
        nValorMulta = nValorMulta + CDbl(.ListItems(x).ListSubItems(4))
        nValorCorrecao = nValorCorrecao + CDbl(.ListItems(x).ListSubItems(5))
        nValorPrincipal1 = nValorPrincipal1 + CDbl(.ListItems(x).ListSubItems(6))
        nValorHon = nValorHon + CDbl(.ListItems(x).ListSubItems(11))
    Next
    Set itmX = lvDestino.ListItems.Add(, , ">>>>")
    itmX.SubItems(1) = "Total --->"
    itmX.SubItems(2) = FormatNumber(nValorPrincipal, 2)
    itmX.SubItems(3) = FormatNumber(nValorJuros, 2)
    itmX.SubItems(4) = FormatNumber(nValorMulta, 2)
    itmX.SubItems(5) = FormatNumber(nValorCorrecao, 2)
    itmX.SubItems(6) = FormatNumber(nValorPrincipal1, 2)
    itmX.SubItems(7) = "N/A"
    itmX.SubItems(8) = "N/A"
    itmX.SubItems(9) = FormatNumber(nSomaJurosValor, 2)
    itmX.SubItems(10) = "N/A"
    If bAjuizado Then
        itmX.SubItems(11) = FormatNumber(nValorHon, 2)
    Else
        itmX.SubItems(11) = FormatNumber(0, 2)
    End If
    itmX.SubItems(12) = ""
    For y = 1 To 12
        itmX.ForeColor = VerdeEscuro
        itmX.ListSubItems(y).ForeColor = VerdeEscuro
    Next y
    
End With

'PREENCHE A COLUNA JUROS APLICADO E TOTAL
'***primeiro contamos quantas linhas tem juros
y = 0: nTotal = 0
With lvDestino
    For x = 1 To .ListItems.Count - 1
        If CDbl(.ListItems(x).ListSubItems(8)) > 0 Then
            y = y + 1
        End If
    Next
   '***calculamos o valor do juros somajuros/qtde parcelas com juros
    If y > 0 Then
        nValorJuros = nSomaJurosValor / (y + 1)
    Else
        nValorJuros = 0
    End If
   '***preenchemos as linhas que tem juros
    For x = 1 To .ListItems.Count - 1
'        If CDbl(.ListItems(x).ListSubItems(8)) > 0 Then
            .ListItems(x).ListSubItems(10) = FormatNumber(nValorJuros, 2)
'        Else
'            .ListItems(x).ListSubItems(10) = "0,00"
'        End If
    Next
    
    bAjuizado = False
    For x = 1 To lvOrigem.ListItems.Count - 1
        If .ListItems(x).Checked = True Then
            If lvOrigem.ListItems(x).SubItems(6) = "S" Then
                bAjuizado = True
                
            End If
        End If
        Exit For
    Next
    
    
    
   '***preenchemos as linhas com honorarios
   If bAjuizado Then
    For x = 1 To .ListItems.Count - 1
        .ListItems(x).ListSubItems(11) = FormatNumber((CDbl(.ListItems(x).ListSubItems(2)) + CDbl(.ListItems(x).ListSubItems(3)) + CDbl(.ListItems(x).ListSubItems(4)) + CDbl(.ListItems(x).ListSubItems(5)) + CDbl(.ListItems(x).ListSubItems(10))) * 0.1, 2)
    Next
    End If
    
    
    '***preenchemos a coluna total da parcela
    For x = 1 To .ListItems.Count - 1
        .ListItems(x).SubItems(12) = FormatNumber(CDbl(.ListItems(x).SubItems(6)) + CDbl(.ListItems(x).SubItems(10)) + CDbl(.ListItems(x).SubItems(11)), 2)
        nTotal = nTotal + CDbl(.ListItems(x).SubItems(12))
    Next
    .ListItems(.ListItems.Count).ListSubItems(12) = FormatNumber(nTotal, 2)
End With

GoTo fim
'CORRIGE ARREDONDAMENTO
nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorTotal = 0: nValorHon = 0: nValorDil = 0

With lvDestino
    For x = 1 To .ListItems.Count
        nValorPrincipal = nValorPrincipal + CDbl(.ListItems(x).SubItems(2))
        nValorJuros = nValorJuros + CDbl(.ListItems(x).SubItems(3))
        nValorMulta = nValorMulta + CDbl(.ListItems(x).SubItems(4))
        nValorCorrecao = nValorCorrecao + CDbl(.ListItems(x).SubItems(5))
        nValorHon = nValorHon + CDbl(.ListItems(x).SubItems(11))
        nValorTotal = nValorTotal + CDbl(.ListItems(x).SubItems(13))
    Next
    
    nDif = CDbl(lblValorPrincipal.Caption) - nValorPrincipal
    .ListItems(nItem).SubItems(2) = FormatNumber(CDbl(.ListItems(nItem).SubItems(2)) + nDif, 2)
    nDif = CDbl(lblValorJuros.Caption) - nValorJuros
    .ListItems(nItem).SubItems(3) = FormatNumber(CDbl(.ListItems(nItem).SubItems(3)) + nDif, 2)
    nDif = CDbl(lblValorMulta.Caption) - nValorMulta
    .ListItems(nItem).SubItems(4) = FormatNumber(CDbl(.ListItems(nItem).SubItems(4)) + nDif, 2)
    nDif = CDbl(lblValorCorrecao.Caption) - nValorCorrecao
    .ListItems(nItem).SubItems(5) = FormatNumber(CDbl(.ListItems(nItem).SubItems(5)) + nDif, 2)
    nDif = CDbl(lblValorHon.Caption) - nValorHon
    .ListItems(nItem).SubItems(11) = FormatNumber(CDbl(.ListItems(nItem).SubItems(11)) + nDif, 2)
    nDif = CDbl(lblValorTotal.Caption) - nValorTotal
    .ListItems(nItem).SubItems(13) = FormatNumber(CDbl(.ListItems(nItem).SubItems(13)) + nDif, 2)
End With

fim:
Exit Sub

Erro:

MsgBox Err.Description
Resume Next
End Sub

Private Sub CarregaTributos()
Dim x As Integer, y As Integer, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim nCodTributo As Integer, sDescTributo As String, bAchou As Boolean

ReDim aTributos(0)
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nAno = .ListItems(x).Text
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            nSeq = .ListItems(x).SubItems(2)
            nParc = .ListItems(x).SubItems(3)
            nCompl = .ListItems(x).SubItems(4)
            Sql = "SELECT debitotributo.codtributo,tributo.desctributo FROM debitotributo INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
            Sql = Sql & "Where debitotributo.CODREDUZIDO =  " & Val(txtCod.Text) & " And debitotributo.AnoExercicio = " & Val(nAno) & " And debitotributo.CodLancamento = " & Val(nLanc) & " And "
            Sql = Sql & "debitotributo.SeqLancamento = " & Val(nSeq) & " AND debitotributo.numparcela = " & nParc & " AND debitotributo.codcomplemento = " & nCompl
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                    nCodTributo = !CodTributo
                    sDescTributo = !desctributo
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
If Val(txtQtdeDil.Text) > 0 Then
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
Dim dDataBase As Date
dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
txtCod.Text = ""
txtNumProc.Text = ""
lblNome.Caption = ""
LimpaMascara mskDataProc
LimpaMascara mskVencto
txtQtdeParc.Text = ""
chkMulta.Value = 1
chkJuros.Value = 1
chkHon.Value = 1
txtValorEntrada.Text = ""
lvOrigem.ListItems.Clear
txtCod.Locked = False
txtCod.BackColor = Branco
mskVencto.Text = Format(dDataBase, "dd/mm/yyyy")
LimpaContador

End Sub

Private Sub Simulado()
Dim df As Integer, bAchou As Boolean, bS As Boolean, bN As Boolean, nValorParcela As Double, nLanc As Integer, z As Variant
Dim x As Integer, nSeq As Integer, a As Integer, sVencimento As String, sVencimento2 As String, y As Integer, nValorEntrada As Double
Dim nValorPrimeira As Double, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double, nValorHon As Double, nValorDil As Double
Dim nValorPrincipal1 As Double, nValorJuros1 As Double, nValorMulta1 As Double, nValorCorrecao1 As Double, nValorTotal1 As Double, nValorHon1 As Double
Dim nPrincipal As Double, nJuros As Double, nMulta As Double, nCorrecao As Double, nTotal As Double, nHonorario As Double, nDiligencia As Double, nItem As Integer
Dim nDia As Integer, nMes As Integer, nAno As Integer, itmX As ListItem, nQtdeParc As Integer, nDif As Double, nPerc As Double, bMI As Boolean

'If Not bAnistia Then
'    If Not ValidaMI Then Exit Sub
'End If
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

If Val(txtValorEntrada.Text) > 0 Then
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

If chkAnistia.Value = 1 Then
    bS = False
    With lvOrigem
        For x = 1 To .ListItems.Count
            If .ListItems(x).Checked = True Then
                If Year(CDate(.ListItems(x).SubItems(5))) = Year(Now) Then
                    bS = True
                End If
            End If
        Next
    End With
    If bS Then
        MsgBox "Não é possivel parcelar débitos com vencimento em 2013 pelo REFIS" & vbCrLf & "no mesmo parcelamento.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

CarregaLancamento
bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False: bMI = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If Val(nLanc) = 1 Or Val(nLanc) = 29 Or Val(nLanc) = 7 Then
                bIPTU = True
            ElseIf Val(nLanc) = 2 Or Val(nLanc) = 3 Or Val(nLanc) = 5 Or Val(nLanc) = 14 Or Val(nLanc) = 49 Then
                bISS = True
            ElseIf Val(nLanc) = 13 Then
                bVS = True
            ElseIf Val(nLanc) = 8 Then
                bTCD = True
            ElseIf Val(nLanc) = 6 Then
                bTLic = True
            ElseIf Val(nLanc) = 69 Then
                bMI = True
            Else
                bDIV = True
            End If
        End If
    Next
End With

'If bTCD And (bIPTU Or bISS Or bVS Or bDIV) Then
'    MsgBox "TCD não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If

'If bitpu And (bISS Or bVS Or bDIV) Then
'    MsgBox "IPTU não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If
'If bISS And (bIPTU Or bVS Or bDIV) Then
'    MsgBox "ISS não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If
'If bVS And (bISS Or bIPTU Or bDIV) Then
'    MsgBox "Vigilância Sanitária não pode ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If
'If bDIV And (bISS Or bVS Or bIPTU) Then
'If bDIV And (bISS Or bVS) Then
'    MsgBox "IPTU,ISS e VS não podem ser parcelado junto com outros lancamentos.", vbExclamation, "Atenção"
'    Exit Sub
'End If

Ocupado

Sql = "DELETE FROM SIMULADOREPARC WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

If chkAnistia.Value = vbChecked And bAnistia Then
    z = Val(txtQtdeParc.Text)
    If z = 0 Then
        MsgBox "Digite o número de parcelas.", vbExclamation, "Atenção"
        Liberado
        Exit Sub
    End If
Else
inipar:
    z = InputBox("Digite o numero de parcelas que deseja exibir no simulado (2-120)", "Parcelas do Simulado", "120")
    If Val(z) > 0 Then
        If Val(z) < 2 Or Val(z) > 120 Then
            MsgBox "Digite um numero entre 2 e 120.", vbExclamation, "Atenção"
            GoTo inipar
        End If
    Else
        GoTo inipar
    End If
End If

'ADICIONA ITENS

For nQtdeParc = 2 To Val(z)
'If nQtde = 35 Then MsgBox "teste"
    'LIMPA TELA
    lvDestino.ListItems.Clear
    DefineParcelas nQtdeParc
        
    For x = 1 To lvDestino.ListItems.Count - 1
        If Val(Right$(lvDestino.ListItems(x).ListSubItems(1), 4)) > 2006 Then
            nValorParcela = CDbl(lvDestino.ListItems(x).ListSubItems(12))
            Exit For
        End If
    Next
    If nValorParcela = 0 Then GoTo Proximo
    If chkAnistia.Value = vbChecked Then
        If nValorParcela < 245 Then
            If nQtdeParc = 2 Then
                Liberado
                MsgBox "Valor mínimo da parcela não chega aos R$250,00.", vbExclamation, "Atenção"
                Exit Sub
            End If
            Exit For
        End If
    Else
       If nValorParcela < 25 Then Exit For
    End If
    
    If bAnistia Then
        If (Val(txtQtdeParc.Text) > 12 And Val(txtQtdeParc.Text) < 25) And nQtdeParc < 13 Then
           GoTo Proximo
        End If
        If (Val(txtQtdeParc.Text) > 24 And Val(txtQtdeParc.Text) < 37) And nQtdeParc < 25 Then
           
           GoTo Proximo
        End If
        If (Val(txtQtdeParc.Text) > 36 And Val(txtQtdeParc.Text) < 49) And nQtdeParc < 37 Then
           GoTo Proximo
        End If
        If (Val(txtQtdeParc.Text) > 48 And Val(txtQtdeParc.Text) < 121) And nQtdeParc < 49 Then
           GoTo Proximo
        End If
    End If
    
    Sql = "INSERT SIMULADOREPARC(COMPUTER,QUANTIDADE,VALOR) VALUES('" & NomeDoUsuario & "'," & nQtdeParc & "," & Virg2Ponto(CStr(nValorParcela)) & ")"
    cn.Execute Sql, rdExecDirect
    
Proximo:
Next

Liberado

If frmMdi.frTeste.Visible = True Then
    If frmMdi.frTeste.Caption = "ACESSANDO OS DADOS LOCAIS" Then
        frmReport.ShowReport "SIMULADO", frmMdi.hwnd, Me.hwnd
    Else
        frmReport.ShowReport "SIMULADOTMP", frmMdi.hwnd, Me.hwnd
    End If
Else
    frmReport.ShowReport "SIMULADO", frmMdi.hwnd, Me.hwnd
End If


Sql = "DELETE FROM SIMULADOREPARC WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub AtualizaTributo()
Dim qd As New rdoQuery, RdoAux3 As rdoResultset
'Atualiza grid tributos
Set qd.ActiveConnection = cn
grdTributo.Rows = 1
lblP.Caption = "0,00": lblJ.Caption = "0,00": lblM.Caption = "0,00": lblC.Caption = "0,00": lblT.Caption = "0,00"
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nAno = .ListItems(x).Text
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            nSeq = .ListItems(x).SubItems(2)
            nParc = .ListItems(x).SubItems(3)
            nCompl = .ListItems(x).SubItems(4)

            On Error Resume Next
            RdoAux3.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = Val(txtCod.Text)
            qd(1) = Val(txtCod.Text) 'codigo
            qd(2) = nAno
            qd(3) = nAno
            qd(4) = nLanc
            qd(5) = nLanc
            qd(6) = nSeq
            qd(7) = nSeq
            qd(8) = nParc
            qd(9) = nParc
            qd(10) = nCompl
            qd(11) = nCompl
            qd(12) = 3
            qd(13) = 3 'statuslanc
            qd(14) = Format(Now, "mm/dd/yyyy") 'data atual
            qd(15) = NomeDoUsuario
            Set RdoAux3 = qd.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            With RdoAux3
                Do Until .EOF
                    Sql = "SELECT DESCTRIBUTO FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    bAchou = False
                    For y = 1 To grdTributo.Rows - 1
                        If Val(grdTributo.TextMatrix(y, 0)) = !CodTributo Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    nTotal = !ValorTributo + !ValorJuros + !ValorMulta + !VALORCORRECAO
                    If bAchou Then
                        grdTributo.TextMatrix(y, 2) = Format(CDbl(grdTributo.TextMatrix(y, 2)) + !ValorTributo, "#0.00")
                        grdTributo.TextMatrix(y, 3) = Format(CDbl(grdTributo.TextMatrix(y, 3)) + !ValorJuros, "#0.00")
                        grdTributo.TextMatrix(y, 4) = Format(CDbl(grdTributo.TextMatrix(y, 4)) + !ValorMulta, "#0.00")
                        grdTributo.TextMatrix(y, 5) = Format(CDbl(grdTributo.TextMatrix(y, 5)) + !VALORCORRECAO, "#0.00")
                        grdTributo.TextMatrix(y, 6) = Format(CDbl(grdTributo.TextMatrix(y, 6)) + nTotal, "#0.00")
                    Else
                        grdTributo.AddItem !CodTributo & Chr(9) & RdoAux!desctributo & Chr(9) & Format(!ValorTributo, "#0.00") & Chr(9) & Format(!ValorJuros, "#0.00") & Chr(9) & Format(!ValorMulta, "#0.00") & Chr(9) & Format(!VALORCORRECAO, "#0.00") & Chr(9) & Format(nTotal, "#0.00")
                    End If
                    RdoAux.Close
                   .MoveNext
                Loop
               .Close
            End With
                
            lblP.Caption = "0,00": lblJ.Caption = "0,00": lblM.Caption = "0,00": lblC.Caption = "0,00": lblT.Caption = "0,00"
            For y = 1 To grdTributo.Rows - 1
                lblP.Caption = Format(CDbl(lblP.Caption) + CDbl(grdTributo.TextMatrix(y, 2)), "#0.00")
                lblJ.Caption = Format(CDbl(lblJ.Caption) + CDbl(grdTributo.TextMatrix(y, 3)), "#0.00")
                lblM.Caption = Format(CDbl(lblM.Caption) + CDbl(grdTributo.TextMatrix(y, 4)), "#0.00")
                lblC.Caption = Format(CDbl(lblC.Caption) + CDbl(grdTributo.TextMatrix(y, 5)), "#0.00")
                lblT.Caption = Format(CDbl(lblT.Caption) + CDbl(grdTributo.TextMatrix(y, 6)), "#0.00")
            Next
            For y = 1 To grdTributo.Rows - 1
                grdTributo.TextMatrix(y, 7) = Format(CDbl(grdTributo.TextMatrix(y, 6)) * 100 / CDbl(lblT.Caption), "#0.00")
            Next
        End If
    Next
End With

End Sub

Private Function ValidaMI() As Boolean
Dim aMI() As Multa, nTipo As Integer, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, x As Integer, y As Integer, t As Integer
Dim bChecar As Boolean
ReDim aMI(0)
bChecar = False

For t = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(t).Checked Then
        If Val(Left(lvOrigem.ListItems(t).SubItems(1), 3)) = 69 Or Right(lvOrigem.ListItems(t).SubItems(1), 4) = "(MI)" Then
            bChecar = True
            Exit For
        End If
    End If
Next

If Not bChecar Then GoTo fim

With lvOrigem
    For t = 1 To .ListItems.Count
        If .ListItems(t).Checked Then
            If Val(Left(.ListItems(t).SubItems(1), 3)) = 69 Or Right(.ListItems(t).SubItems(1), 4) = "(MI)" Then
               'CARREGA DADOS DA MULTA
                nAno = Val(.ListItems(t).Text)
                nLanc = Val(Left(.ListItems(t).SubItems(1), 3))
                nSeq = Val(.ListItems(t).SubItems(2))
                nParc = Val(.ListItems(t).SubItems(3))
                nCompl = Val(.ListItems(t).SubItems(4))
                If Val(Left(.ListItems(t).SubItems(1), 3)) = 69 Then
                    nTipo = 2 'VIEW MULTA
                Else
                    nTipo = 3 'VIEW LANC
                    Sql = "SELECT * FROM MULTAINFRACAO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        nAno = !Ano
                        nLanc = !LANCAMENTO
                        nSeq = !Sequencia
                        nParc = !Parcela
                        nCompl = !Complemento
                       .Close
                    End With
                End If
                    
                'CARREGA BLOCO DAS MULTAS
                'Sql = "SELECT * FROM MULTAINFRACAO WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND ANO=" & nAno & " AND LANCAMENTO=" & nLanc & " AND SEQUENCIA=" & nSeq & " AND PARCELA=" & nParc & " AND COMPLEMENTO=" & nCompl
                Sql = "SELECT * FROM MULTAINFRACAO WHERE CODIGO=" & Val(txtCod.Text) & " AND ANO=" & nAno & " AND LANCAMENTO=" & nLanc & " AND SEQUENCIA=" & nSeq & " AND PARCELA=" & nParc & " AND COMPLEMENTO=" & nCompl
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount = 0 Then
                        GoTo Erro
                    End If
                    aMI(0).nAno = !Ano
                    aMI(0).nLanc = !LANCAMENTO
                    aMI(0).nSeq = !Sequencia
                    aMI(0).nParc = !Parcela
                    aMI(0).nCompl = !Complemento
                    aMI(0).bAchou = False
                    Do Until .EOF
                        ReDim Preserve aMI(UBound(aMI) + 1)
                        aMI(UBound(aMI)).nAno = !AnoExercicio
                        aMI(UBound(aMI)).nLanc = !CodLancamento
                        aMI(UBound(aMI)).nSeq = !SeqLancamento
                        aMI(UBound(aMI)).nParc = !NumParcela
                        aMI(UBound(aMI)).nCompl = !CODCOMPLEMENTO
                        aMI(UBound(aMI)).bAchou = False
                       .MoveNext
                    Loop
                   .Close
                End With
                
               'CONFRONTA BLOCO COM OS REGISTROS SELECIONADOS
                With lvOrigem
                    For x = 1 To .ListItems.Count
                       'APENAS AS LINHAS SELECIONADAS QUE SEJAM MULTA OU MI
                        If (Val(Left(.ListItems(x).SubItems(1), 3)) = 69 Or Right(.ListItems(x).SubItems(1), 4) = "(MI)") And .ListItems(x).Checked Then
                            nAno = Val(.ListItems(x).Text)
                            nLanc = Val(Left(.ListItems(x).SubItems(1), 3))
                            nSeq = Val(.ListItems(x).SubItems(2))
                            nParc = Val(.ListItems(x).SubItems(3))
                            nCompl = Val(.ListItems(x).SubItems(4))
                            For y = 0 To UBound(aMI)
                                If aMI(y).nAno = nAno And aMI(y).nLanc = nLanc And aMI(y).nSeq = nSeq And aMI(y).nParc And aMI(y).nCompl = nCompl Then
                                    aMI(y).bAchou = True 'marca a parcela
                                End If
                            Next
                        End If
                    Next
                End With
                
               'VERIFICA SE ALGUMA PARCELA DA MATRIZ É NEGATIVA
                For y = 0 To UBound(aMI)
                    If aMI(y).bAchou = False Then
                        GoTo Erro
                    End If
                Next
'            Else
'                GoTo fim
            End If
          End If
        Next
End With


fim:
ValidaMI = True
Exit Function

Erro:
MsgBox "Todas as parcelas que constituem a multa de infração devem ser selecionadas.", vbCritical, "Atenção"
ValidaMI = False

End Function

Private Function ContaParcela() As Boolean
Dim Sql As String, RdoAux As rdoResultset, aParcela() As Parcela, bAchou As Boolean, x As Integer, y As Integer
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, bDirty As Boolean
Dim sNumProc As String

bAchou = False
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = True Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    'Nenhuma parcela foi selecionada.
    ContaParcela = True
    Exit Function
End If

ContaParcela = False: ReDim aParcela(0)

Sql = "SELECT processoreparc.numprocesso, processoreparc.numproc, processoreparc.anoproc, processoreparc.dataprocesso, origemreparc.anoexercicio, "
Sql = Sql & "origemreparc.CodLancamento , origemreparc.numsequencia, origemreparc.NumParcela, origemreparc.CODCOMPLEMENTO FROM origemreparc INNER JOIN "
Sql = Sql & "processoreparc ON origemreparc.numprocesso = processoreparc.numprocesso Where origemreparc.CODREDUZIDO = " & Val(txtCod.Text) & " AND "
Sql = Sql & "(YEAR(processoreparc.datareparc) > 2008) ORDER BY origemreparc.anoexercicio, origemreparc.numparcela"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        bAchou = False
        For x = 1 To UBound(aParcela)
            If aParcela(x).nAno = !AnoExercicio And aParcela(x).nLanc = !CodLancamento And aParcela(x).nSeq = !numsequencia And aParcela(x).nParc = !NumParcela And aParcela(x).nCompl = !CODCOMPLEMENTO Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aParcela(UBound(aParcela) + 1)
            aParcela(UBound(aParcela)).nAnoProc1 = !AnoProc
            aParcela(UBound(aParcela)).nNumProc1 = !NumProc
            aParcela(UBound(aParcela)).nAno = !AnoExercicio
            aParcela(UBound(aParcela)).nLanc = !CodLancamento
            aParcela(UBound(aParcela)).nSeq = !numsequencia
            aParcela(UBound(aParcela)).nParc = !NumParcela
            aParcela(UBound(aParcela)).nCompl = !CODCOMPLEMENTO
            aParcela(UBound(aParcela)).nContador = 1
        Else
            aParcela(x).nAnoProc2 = !AnoProc
            aParcela(x).nNumProc2 = !NumProc
            aParcela(x).nContador = 2
        End If
       .MoveNext
    Loop
   .Close
End With

bDirty = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            If .ListItems(x).SubItems(6) = "N" Then
                nAno = .ListItems(x).Text
                nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
                nSeq = .ListItems(x).SubItems(2)
                nParc = .ListItems(x).SubItems(3)
                nCompl = .ListItems(x).SubItems(4)
                
                For y = 1 To UBound(aParcela)
                    With aParcela(y)
                        If .nAno = nAno And .nLanc = nLanc And .nSeq = nSeq And .nParc = nParc And .nCompl = nCompl And .nContador = 2 Then
                            aParcela(y).nContador = 3: bDirty = True
                            Exit For
                        End If
                    End With
                Next y
            End If
        End If
    Next x
End With

grdTemp.Rows = 1
'If bDirty Then
'    For x = 1 To UBound(aParcela)
'        With aParcela(x)
'            If .nContador = 3 Then
'                sNumProc = .nNumProc1 & "-" & RetornaDVProcesso(.nNumProc1) & "/" & .nAnoProc1 & " E " & .nNumProc2 & "-" & RetornaDVProcesso(.nNumProc2) & "/" & .nAnoProc2
'                grdTemp.AddItem .nAno & Chr(9) & .nLanc & Chr(9) & .nSeq & Chr(9) & .nParc & Chr(9) & .nCompl & Chr(9) & sNumProc
'            End If
'        End With
'    Next
 '   fr4.Enabled = False
 '   pnlContaParcela.Visible = True
'Else
    ContaParcela = True
'End If

End Function

Private Sub EmiteBoleto()

On Error GoTo Erro
Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String
Dim dDataBase As Date

dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))

If txtNumProc.Text = "" Then
    MsgBox "Digite o número do Processo.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataProc.Text) Then
    MsgBox "Data do Processo inválido.", vbCritical, "Atenção"
    Exit Sub
End If

'******ANISTIA*********
If bAnistia And chkAnistia.Value = 1 Then
    With lvOrigem
        For x = 1 To .ListItems.Count
            If .ListItems(x).Checked Then
                If CDate(.ListItems(x).SubItems(5)) > CDate("31/12/2012") And Val(Left(.ListItems(x).SubItems(1), 2)) <> 69 Then
                    MsgBox "Existem lancamentos posteriores a 31/12/2012 e não podem ser anistiados.", vbCritical, "Atenção"
                    Exit Sub
                End If
            End If
        Next
    End With

    If CDbl(lvDestino.ListItems(1).SubItems(12)) < 245 Then
        MsgBox "Valor mínimo da parcela R$250,00 (R$245,00 + R$5,00 de serviços)", vbCritical, "Atenção"
        Exit Sub
    End If
Else
    nPlano = 0
End If

If Year(CDate(lvDestino.ListItems(1).SubItems(1))) < Year(Now) Then
    MsgBox "Por favor,verifique os vencimentos das parcelas, caso esteja errado coloque para gerar novamente.", vbCritical, "Atenção"
    Exit Sub
End If



'*********************

If MsgBox("Deseja Gravar este Parcelamento ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
'Exit Sub
nNumProc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
nAnoProc = Right$(txtNumProc.Text, 4)
lblNumProc.Caption = nNumProc
lblAnoProc.Caption = nAnoProc
sNumProc = CStr(nNumProc) & "/" & CStr(nAnoProc)

'VERIFICA SE O PROCESSO JA FOI UTILIZADO
Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & " '"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este processo já foi utilizado em um parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

'GRAVA O PROCESSO
Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,NUMPROC,ANOPROC,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CODIGORESP,FUNCIONARIO,PLANO,NOVO) VALUES('"
Sql = Sql & sNumProc & "'," & nNumProc & "," & nAnoProc & ",'" & Format(mskDataProc.Text, "mm/dd/yyyy") & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "',"
Sql = Sql & Val(txtQtdeParc.Text) & "," & Virg2Ponto(txtValorEntrada.Text) & "," & 0 & ","
Sql = Sql & IIf(chkMulta.Value = vbChecked, 1, 0) & "," & IIf(chkJuros.Value = vbChecked, 1, 0) & ","
Sql = Sql & Val(txtCod.Text) & ",'" & NomeDeLogin & "'," & nPlano & "," & 1 & ")"
cn.Execute Sql, rdExecDirect

bIPTU = False: bISS = False: bVS = False: bDIV = False: bTCD = False: bTLic = False
With lvOrigem
    For x = 1 To .ListItems.Count
        If .ListItems(x).Checked = True Then
            nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
            If nLanc = 1 Or nLanc = 29 Then
                bIPTU = True
            ElseIf nLanc = 2 Or nLanc = 3 Or nLanc = 5 Or nLanc = 6 Or nLanc = 14 Then
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
    
AtualizaTributo
    
'GRAVA AS PARCELAS DE DESTINO
With lvDestino
    For x = 1 To .ListItems.Count - 1
        sDataVencto = .ListItems(x).SubItems(1)
        nAno = Val(Right$(sDataVencto, 4))
        nLanc = 20
        nSeq = Val(lblSeq.Caption)
        nParc = .ListItems(x).Text
        nCompl = 0
       'GRAVA DESTINOREPARC
        Sql = "INSERT DESTINOREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,VALORLIQUIDO,"
        Sql = Sql & "JUROS,MULTA,CORRECAO,VALORPRINCIPAL,SALDO,JUROSPERC,JUROSVALOR,JUROSAPL,HONORARIO,TOTAL) VALUES('"
        Sql = Sql & sNumProc & "'," & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(2)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(3)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(4)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(5)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(6)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(7)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(8)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(9)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(10)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(11)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(12)))) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA DEBITOPARCELA
        If nAno = Year(dDataBase) Then
            nSituacao = 3
        Else
            nSituacao = 18
        End If
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
        Sql = Sql & nSituacao & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
        cn.Execute Sql, rdExecDirect
       
       'DEFINE TRIBUTO PRINCIPAL
        With grdTributo
            For y = 1 To .Rows - 1
                nValorTributo = CDbl(lvDestino.ListItems(x).SubItems(2))
                nCodTrib = Val(.TextMatrix(y, 0))
                nPerc = CDbl(.TextMatrix(y, 7))
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
                Sql = Sql & nParc & "," & nCompl & "," & nCodTrib & "," & Virg2Ponto(CStr(nValorTributo * nPerc / 100)) & ")"
                cn.Execute Sql, rdExecDirect
            Next
        End With
       
       '***********ANISTIA*********
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 609 - TX.PARCEL.REFIS)
        If Val(lblAnistia.Caption) > 0 And bAnistia Then
            nValorExp = 5
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 609 & "," & Virg2Ponto(CStr(nValorExp)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       '**************************
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 113 - JUROS)
        nValorTributo = CDbl(.ListItems(x).SubItems(3))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 113 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 112 - MULTA)
        nValorTributo = CDbl(.ListItems(x).SubItems(4))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 112 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 26 - CORREÇÃO)
        nValorTributo = CDbl(.ListItems(x).SubItems(5))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 26 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
       'GRAVA DEBITOTRIBUTO   // (TRIBUTO 585 - JUROS APLICADO)
        nValorTributo = CDbl(.ListItems(x).SubItems(10))
        If nValorTributo > 0 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & 585 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
        If chkHon.Value = 1 And CDbl(.ListItems(x).SubItems(11)) > 0 Then
            'GRAVA DEBITOTRIBUTO   // (TRIBUTO 90 - HONORÁRIOS)
             nValorTributo = CDbl(.ListItems(x).SubItems(11))
             Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
             Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & ","
             Sql = Sql & nParc & "," & nCompl & "," & 90 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
             cn.Execute Sql, rdExecDirect
        End If
        'RETORNA ULTIMO DOCUMENTO
        Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nLastCod = !MAXIMO + 1
           .Close
        End With
       'GRAVA NUMDOCUMENTO
        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO) VALUES("
        Sql = Sql & nLastCod & ",'" & Format(dDataBase, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorExp)) & "," & IIf(Val(lblAnistia.Caption) > 0, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ")"
        cn.Execute Sql, rdExecDirect
       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.Text) & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & ","
        Sql = Sql & nCompl & "," & nLastCod & ")"
        cn.Execute Sql, rdExecDirect
       'PREENCHE DOCUMENTO NO LVDESTINO
        .ListItems(x).SubItems(13) = nLastCod
    Next
End With

'GRAVA AS PARCELAS DE ORIGEM
With lvOrigem
    For x = 1 To .ListItems.Count
        If lvOrigem.ListItems(x).Checked = True Then
             nAno = .ListItems(x).Text
             nLanc = Val(Left$(.ListItems(x).SubItems(1), 2))
             nSeq = .ListItems(x).SubItems(2)
             nParc = .ListItems(x).SubItems(3)
             nCompl = .ListItems(x).SubItems(4)
            'GRAVA ORIGEMREPARC
             Sql = "INSERT ORIGEMREPARC (NUMPROCESSO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,"
             Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,PRINCIPAL,JUROS,MULTA,CORRECAO) VALUES('" & sNumProc & "'," & Val(txtCod.Text) & ","
             Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(7)))) & ","
             Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(8)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(9)))) & ","
             Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(10)))) & ")"
             cn.Execute Sql, rdExecDirect
            'ATUALIZA O STATUS DE ORIGEM   // (4 - REPARCELADO)
             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=4 WHERE CODREDUZIDO=" & Val(txtCod.Text)
             Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
             Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
             Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
             cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'LIMPA TEMPORARIO
'Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
'cn.Execute Sql, rdExecDirect

'DADOS CABEÇALHO
'dDataProc = CDate(mskDataProc.Text)
'sDescImposto = "REPARCELAMENTO"
'NumBarra1 = Format(ExtraiNumero(txtNumProc.Text), "0000000000")
'StrBarra1 = Gera2of5Str(NumBarra1)

nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect



'ENDEREÇO DO CONTRIBUINTE
Select Case Val(txtCod.Text)
    Case 1 To 99999
        'DADOS DO IMOVEL
        Sql = "SELECT * FROM vwCnsImovel WHERE CODREDUZIDO=" & Val(txtCod.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNumInsc = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
                sCodReduz = Format(Val(txtCod.Text), "000000") & "-" & RetornaDVCodReduzido(Val(txtCod.Text))
                Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
                Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
                Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & Val(txtCod.Text) & " AND TIPOPROP='P' AND PRINCIPAL=1"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset)
                sNomeResp = RdoAux2!nomecidadao
                RdoAux2.Close
                sTipoImposto = "REPARCEL."
                sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                nNumImovel = !Li_Num
                sComplImovel = IIf(IsNull(!Li_Compl) Or !Li_Compl = "", " ", !Li_Compl)
                If !CodBairro <> 999 Then
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
               Sql = Sql & "WHERE CODREDUZIDO=" & Val(txtCod.Text)
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
        Sql = "SELECT * "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & Val(txtCod.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNumInsc = SubNull(!INSCESTADUAL)
                sCodReduz = !codigomob
                sNomeResp = !razaosocial
                sTipoImposto = "REPARCEL."
                If IsNull(!NomeLogradouro) Then
                    sEndImovel = !NomeLogr
                Else
                    sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                End If
                nNumImovel = Val(SubNull(!Numero))
                sComplImovel = SubNull(!Complemento)
                If !CodCidade <> 413 Then
                    sBairroImovel = !DescBairro
                    sCidadeEntrega = !desccidade
                End If
                GoTo fim
                If !CodBairro <> 999 Then
                     Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO=" & !CodBairro
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
                Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.Text)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sEndEntrega = SubNull(!NomeLogradouro)
                        nNumEntrega = SubNull(!numimovel)
                        sBairroEntrega = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        sComplEntrega = SubNull(!Complemento)
                        sCepEntrega = SubNull(!Cep)
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
fim:
    Case 500000 To 800000
        Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao,cidadao.cpf, cidadao.cnpj, cidadao.rg, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, "
        Sql = Sql & "cidadao.siglauf, cidade.desccidade, bairro.descbairro, cidadao.nomelogradouro AS nomerua, "
        Sql = Sql & "cidadao.codlogradouro , vwLOGRADOURO.AbrevTipoLog, vwLOGRADOURO.AbrevTitLog, vwLOGRADOURO.NomeLogradouro "
        Sql = Sql & "FROM  vwLOGRADOURO RIGHT JOIN  cidadao ON vwLOGRADOURO.CODLOGRADOURO = cidadao.codlogradouro LEFT OUTER JOIN "
        Sql = Sql & "cidade INNER JOIN  bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade ON cidadao.siglauf = bairro.siglauf AND "
        Sql = Sql & "cidadao.codcidade = bairro.codcidade And cidadao.codbairro = bairro.codbairro "
        Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & Val(txtCod.Text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sTipoImposto = "REPARCEL."
                sCodReduz = !CodCidadao
                sNomeResp = !nomecidadao
                If Not IsNull(!NomeLogradouro) Then
                    sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                Else
                    sEndImovel = SubNull(!nomerua)
                End If
                nNumImovel = SubNull(!numimovel)
                sComplImovel = SubNull(!Complemento)
                nCodBairro = Val(SubNull(!CodBairro))
                sCidadeEntrega = SubNull(!desccidade)
                sUFEntrega = SubNull(!SiglaUF)
                If Not IsNull(!CPF) Then
                   sCPF = !CPF
                ElseIf Not IsNull(!Cnpj) Then
                   sCPF = !Cnpj
                ElseIf Not IsNull(!rg) Then
                   sCPF = !rg
                Else
                   sCPF = ""
                End If
             Else
                sCPF = ""
             End If
        '     If nCodBairro = 0 Then
        '        sBairroImovel = SubNull(!NOMEBairro)
        '        sCidadeEntrega = SubNull(!NomeCidade)
        '     End If
        End With
End Select

'GRAVA TEMPORARIO
With lvDestino
    For x = 1 To .ListItems.Count - 1
        sDataVencto = .ListItems(x).SubItems(1)
        nAno = Val(Right$(sDataVencto, 4))
        nLanc = 20
        nSeq = Val(lblSeq.Caption)
        nParc = .ListItems(x).Text
        nCompl = 0
'        nAno = .ListItems(x).text
'        nSeq = lblSeq.Caption
'        nParc = .ListItems(x).SubItems(2)
        nNumDoc = .ListItems(x).SubItems(13)


        nNumGuia = nNumDoc

        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)


'        sDataVencto = .ListItems(x).SubItems(3)
        If bAnistia And chkAnistia.Value = vbChecked Then
            sValorParc = CStr(CDbl(.ListItems(x).SubItems(12)) + 5) '5,00=Valor da Taxa
            nValorTaxa = 0
        Else
            sValorParc = .ListItems(x).SubItems(12)
            Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE CODTRIBUTO=3 AND ANO=" & Year(Now)
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nValorTaxa = RdoAux2!VALORALIQ
            RdoAux2.Close
        End If
        nValorGuia = sValorParc + nValorTaxa
        sNossoNumero = Format(sNumDoc3, "0000000000000")
        sDv = Trim(Calculo_DV10("028" & Left(sNossoNumero, 7)))
        sDigitavel = "0339912354028" & Left(sNossoNumero, 7) & sDv

        sDigitavel = sDigitavel & Right(sNossoNumero, 6) & "0102"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv

        dDataBase = "07/10/1997"
        nFatorVencto = CDate(sDataVencto) - dDataBase
        sQuintoGrupo = Format(nFatorVencto, "0000")
        sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorGuia, 2)), "0000000000")

        sBarra = "0339" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorGuia, 2)), "0000000000") & "91235028"
        sBarra = sBarra & sNossoNumero & "0102"
        sDv = Trim(Calculo_DV11(sBarra))
        sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)

        sDigitavel = sDigitavel & sDv & sQuintoGrupo

        sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
        sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
        sBarra = Gera2of5Str(sBarra)

        Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numproc) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & Val(txtCod.Text) & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 1, nParc) & "," & Val(txtQtdeParc.Text) & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & "Parcelamento: " & Left$(txtNumProc.Text, 25) & "','" & Left$(txtNumProc.Text, 25) & "')"
        cn.Execute Sql, rdExecDirect
        
        
'        NumBarra2 = Gera2of5Cod(CStr(CDbl(sValorParc) + nValorExp), CDate(sDataVencto), nNumDoc, nParc, 20, nSeq, 0)
'        NumBarra2a = Left$(NumBarra2, 13)
'        NumBarra2b = Mid$(NumBarra2, 14, 13)
'        NumBarra2c = Mid$(NumBarra2, 27, 13)
'        NumBarra2d = Right$(NumBarra2, 13)
'        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        
'        Sql = "INSERT CARNETMP(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,ENDIMOVEL,NUMIMOVEL,COMPLIMOVEL,"
'        Sql = Sql & "BAIRROIMOVEL,ENDENTREGA,NUMENTREGA,COMPLENTREGA,BAIRROENTREGA,CEPENTREGA,CIDADEENTREGA,UFENTREGA,"
 '       Sql = Sql & "DESCIMPOSTO,EXERCICIO,NUMPROCESSO,DATAPROCESSO,NUMDOCUMENTO,DV,QUADRA,LOTE,DATAVENCTO,NUMPARCELA,"
 '       Sql = Sql & "NUMTOTPARCELA,VALORPARCELA,STRBARRA1,STRBARRA2,NUMBARRA1,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
 '       Sql = Sql & "DADOSLANCAMENTO,TAXAEXP,SAIR) VALUES('" & NomeDoUsuario & "'," & x & ",'" & sNumInsc & "','" & sCodReduz & "','"
 '       Sql = Sql & sTipoImposto & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left$(sEndImovel, 40) & "'," & nNumImovel & ",'" & Left$(sComplImovel, 30) & "','"
 '       Sql = Sql & Left$(sBairroImovel, 25) & "','" & Left$(sEndEntrega, 40) & "'," & nNumEntrega & ",'" & Left$(sComplEntrega, 30) & "','" & Left$(sBairroEntrega, 25) & "','"
 '       Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Left$(sDescImposto, 30) & "'," & nAno & ",'" & Left$(txtNumProc.Text, 25) & "','"
 '       Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "'," & nNumDoc & "," & RetornaDVNumDoc(nNumDoc) & ",'" & sQuadra & "','"
 '       Sql = Sql & sLote & "','" & Format(sDataVencto, "mm/dd/yyyy") & "'," & IIf(nParc = 0, 1, nParc) & "," & Val(txtQtdeParc.Text) & ","
 '       Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
 '       Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & "REPARCELAMENTO" & "'," & Virg2Ponto(CDbl(nValorExp)) & "," & "0" & ")"
  '      cn.Execute Sql, rdExecDirect
    Next
End With

For x = 1 To grdTributo.Rows - 1
'    If lvTrib.ListItems(x).Checked Then
        Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & ","
        Sql = Sql & Val(grdTributo.TextMatrix(x, 0)) & ",'" & grdTributo.TextMatrix(x, 1) & "'," & Virg2Ponto(RemovePonto(grdTributo.TextMatrix(x, 6))) & ")"
        cn.Execute Sql, rdExecDirect
       ' nLinha = nLinha + 1
 '   End If
Next
x = x + 1
Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & ","
Sql = Sql & 3 & ",'Taxa de Expediente por parcela'," & Virg2Ponto(RemovePonto(CStr(nValorTaxa))) & ")"
cn.Execute Sql, rdExecDirect


frmConfissaoDivida.txtNumProc.Text = sNumProc
frmConfissaoDivida.txtNumProc.Locked = True
frmConfissaoDivida.lblSid.Caption = nSid
frmConfissaoDivida.CarregaProcesso
cmdPrint_Click
'cmdVoltar_Click
'Limpar
Ocupado
frmConfissaoDivida.show
frmConfissaoDivida.ZOrder 0
Liberado
cmdVoltar_Click
Limpar
'Unload frmParcelamento2
Exit Sub

Erro:

MsgBox Err.Description
Resume Next

End Sub
