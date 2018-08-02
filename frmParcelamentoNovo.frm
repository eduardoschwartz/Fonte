VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmParcelamentoNovo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parcelamento de Dívida (Alteração LC 07/1992 em 2017)"
   ClientHeight    =   8205
   ClientLeft      =   6315
   ClientTop       =   3480
   ClientWidth     =   13995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   13995
   Begin VB.Frame frDetalhe 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5475
      Left            =   3450
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   7905
      Begin MSFlexGridLib.MSFlexGrid grdTributo 
         Height          =   1935
         Left            =   30
         TabIndex        =   33
         Top             =   270
         Width           =   7845
         _ExtentX        =   13838
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
         FormatString    =   ">Código |<Nome do Tributo           |>Principal     |>Juros        |>Multa        |>Correção     |>Total       |>%         "
      End
      Begin MSFlexGridLib.MSFlexGrid grdDestino 
         Height          =   2655
         Left            =   30
         TabIndex        =   43
         Top             =   2790
         Width           =   7845
         _ExtentX        =   13838
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   12
         Cols            =   9
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
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   ">Código |<Nome do Tributo           |>Total          |>%         |>Proporção |>Parcela 1  |>Arr.     |>Parcela N |>Arr.    "
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Destino do Parcelamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   42
         Top             =   2520
         Width           =   7605
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         Caption         =   "Origem do Parcelamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Index           =   0
         Left            =   30
         TabIndex        =   41
         Top             =   30
         Width           =   7605
      End
      Begin VB.Label lblP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2400
         TabIndex        =   40
         Top             =   2220
         Width           =   885
      End
      Begin VB.Label lblJ 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   3300
         TabIndex        =   39
         Top             =   2220
         Width           =   825
      End
      Begin VB.Label lblM 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4140
         TabIndex        =   38
         Top             =   2220
         Width           =   825
      End
      Begin VB.Label lblC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4980
         TabIndex        =   37
         Top             =   2220
         Width           =   945
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   5940
         TabIndex        =   36
         Top             =   2220
         Width           =   735
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
         TabIndex        =   35
         Top             =   2220
         Width           =   2385
      End
      Begin VB.Label lblPerc 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   6720
         TabIndex        =   34
         Top             =   2220
         Width           =   495
      End
   End
   Begin VB.Frame frDDList 
      BackColor       =   &H00EEEEEE&
      Height          =   375
      Left            =   10200
      TabIndex        =   14
      Top             =   60
      Width           =   1230
      Begin VB.ListBox lstAno 
         Height          =   1635
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   405
         Width           =   1140
      End
      Begin prjChameleon.chameleonButton cmdDDList 
         Height          =   240
         Left            =   270
         TabIndex        =   13
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
         MICON           =   "frmParcelamentoNovo.frx":0000
         PICN            =   "frmParcelamentoNovo.frx":001C
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
   Begin VB.Frame frTop 
      Height          =   975
      Left            =   30
      TabIndex        =   23
      Top             =   -60
      Width           =   13905
      Begin VB.CheckBox chkRefis 
         Appearance      =   0  'Flat
         Caption         =   "Refis"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   8430
         TabIndex        =   46
         Top             =   630
         Width           =   765
      End
      Begin VB.TextBox txtCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   900
         MaxLength       =   6
         TabIndex        =   0
         Top             =   210
         Width           =   705
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1650
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   210
         Width           =   4965
      End
      Begin VB.CheckBox chkMulta 
         Appearance      =   0  'Flat
         Caption         =   "Multa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3420
         TabIndex        =   2
         Top             =   630
         Width           =   735
      End
      Begin VB.CheckBox chkJuros 
         Appearance      =   0  'Flat
         Caption         =   "Juros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4260
         TabIndex        =   3
         Top             =   630
         Width           =   735
      End
      Begin VB.CheckBox chkCorrecao 
         Appearance      =   0  'Flat
         Caption         =   "Correção"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   4
         Top             =   630
         Width           =   975
      End
      Begin VB.CheckBox chkHonorario 
         Appearance      =   0  'Flat
         Caption         =   "Honorários"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6090
         TabIndex        =   5
         Top             =   630
         Width           =   1125
      End
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   900
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   570
         Width           =   2025
      End
      Begin VB.CheckBox chkPenhorado 
         Appearance      =   0  'Flat
         Caption         =   "Penhorado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7260
         TabIndex        =   6
         Top             =   630
         Width           =   1095
      End
      Begin prjChameleon.chameleonButton cmdAnalisar 
         Height          =   345
         Left            =   12510
         TabIndex        =   7
         ToolTipText     =   "Analisar os débitos"
         Top             =   180
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Analisar"
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
         MICON           =   "frmParcelamentoNovo.frx":0176
         PICN            =   "frmParcelamentoNovo.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskVencto 
         Height          =   285
         Left            =   8130
         TabIndex        =   1
         Top             =   210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         MouseIcon       =   "frmParcelamentoNovo.frx":0232
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
      Begin VB.Label Label3 
         Caption         =   "DI.:"
         Height          =   225
         Index           =   1
         Left            =   13110
         TabIndex        =   52
         Top             =   630
         Width           =   255
      End
      Begin VB.Label lblDI 
         Caption         =   "N"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   13410
         TabIndex        =   51
         Top             =   630
         Width           =   255
      End
      Begin VB.Label lblAj 
         Caption         =   "N"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   12810
         TabIndex        =   45
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Aj.:"
         Height          =   225
         Index           =   2
         Left            =   12510
         TabIndex        =   44
         Top             =   630
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Anos..:"
         Height          =   225
         Left            =   9540
         TabIndex        =   29
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Código....:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "1º Vencimento..:"
         Height          =   195
         Index           =   1
         Left            =   6870
         TabIndex        =   27
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cpf/Cnpj.:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   26
         Top             =   630
         Width           =   735
      End
   End
   Begin Tributacao.jcFrames frWait 
      Height          =   705
      Left            =   4620
      Top             =   2040
      Visible         =   0   'False
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   1244
      FrameColor      =   255
      FillColor       =   4210688
      TextBoxColor    =   8454016
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
      ThemeColor      =   3
      ColorFrom       =   192
      ColorTo         =   8438015
      Begin VB.Label lblWait 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aguarde...Processando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   210
         Width           =   3885
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   765
      Left            =   30
      Top             =   3780
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   1349
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
      Begin VB.TextBox txtPlano 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1380
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "(SEM PLANO)"
         Top             =   60
         Width           =   4785
      End
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3810
         TabIndex        =   9
         Top             =   390
         Width           =   1275
      End
      Begin VB.ComboBox cmbQtde 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   390
         Width           =   1005
      End
      Begin prjChameleon.chameleonButton cmdOpcao 
         Height          =   345
         Left            =   12510
         TabIndex        =   10
         ToolTipText     =   "Menu de Opções"
         Top             =   300
         Width           =   1185
         _ExtentX        =   2090
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
         MCOL            =   16711935
         MPTR            =   1
         MICON           =   "frmParcelamentoNovo.frx":024E
         PICN            =   "frmParcelamentoNovo.frx":026A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblPercDesconto 
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   8550
         TabIndex        =   50
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "% desconto de multa e juros..:"
         Height          =   225
         Index           =   2
         Left            =   6330
         TabIndex        =   49
         Top             =   120
         Width           =   2115
      End
      Begin VB.Label Label1 
         Caption         =   "Plano Desconto:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   47
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblDataProc 
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   5190
         TabIndex        =   31
         Top             =   150
         Width           =   945
      End
      Begin VB.Label lbl30 
         Caption         =   "30% de 1254,52=2532,22"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   10020
         TabIndex        =   30
         Top             =   450
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblValorAdicional 
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   8550
         TabIndex        =   22
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Valor adicional a 1ª parcela..:"
         Height          =   225
         Index           =   1
         Left            =   6360
         TabIndex        =   21
         Top             =   450
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Nº do Processo..:"
         Height          =   225
         Index           =   0
         Left            =   2490
         TabIndex        =   18
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "Qtde. Parcelas..:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   450
         Width           =   1215
      End
   End
   Begin MSComctlLib.ListView lvDestino 
      Height          =   3405
      Left            =   30
      TabIndex        =   12
      Top             =   4770
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   6006
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
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nothing"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Pc"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Vencto."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Vl.Liquido"
         Object.Width           =   1763
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Juros"
         Object.Width           =   1483
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Multa"
         Object.Width           =   1483
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Correção"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Principal"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Saldo"
         Object.Width           =   1765
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Jr%"
         Object.Width           =   899
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Jur/Mes"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Jur Apl."
         Object.Width           =   1498
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Hon."
         Object.Width           =   1130
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Total"
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvOrigem 
      Height          =   2625
      Left            =   30
      TabIndex        =   11
      Top             =   1140
      Width           =   13905
      _ExtentX        =   24527
      _ExtentY        =   4630
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Ano"
         Object.Width           =   1340
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Lançamento"
         Object.Width           =   4939
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
         Text            =   "Qtde"
         Object.Width           =   1059
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "%"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "Vl.Entrada"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   14
         Text            =   "Vl.Total"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H00000080&
      Caption         =   " Débitos gerados após o parcelamento"
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
      TabIndex        =   19
      Top             =   4560
      Width           =   13905
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
      Index           =   0
      Left            =   30
      TabIndex        =   16
      Top             =   930
      Width           =   13905
   End
   Begin VB.Menu mnuOpcao 
      Caption         =   "Opções"
      Visible         =   0   'False
      Begin VB.Menu mnuCriar 
         Caption         =   "Gerar parcelamento"
      End
      Begin VB.Menu mnuSimulado 
         Caption         =   "Simulado"
      End
      Begin VB.Menu mnuCalculo 
         Caption         =   "Folha de cálculo"
      End
      Begin VB.Menu mnuDetalhe 
         Caption         =   "Exibir detalhes"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelar 
         Caption         =   "Cancelar"
      End
   End
End
Attribute VB_Name = "frmParcelamentoNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tTotal
    Valor_Principal As Double
    Valor_Multa As Double
    Valor_Juros As Double
    Valor_Correcao As Double
    Valor_Honorario As Double
    Valor_Adicional As Double
End Type

Private Type tDestino
    Qtde_Parcela As Integer
    Numero_Parcela As Integer
    Data_Vencimento As Date
    Valor_Liquido As Double
    Valor_Juros As Double
    Valor_Multa As Double
    Valor_Correcao As Double
    Valor_Principal As Double
    Saldo As Double
    juros_perc As Double
    juros_mes As Double
    juros_aplicado As Double
    honorario As Double
    valor_parcela As Double
    penalidade As Double
    Valor_Total As Double
End Type

Private Type TRIBUTO
    nCodTributo  As Integer
    sNomeTributo As String
    nValorTributo As Double
    nPercentual As Double
End Type

Dim aTotal() As tTotal, aDestino() As tDestino, bExec As Boolean, nValor_Minimo_Fisica As Double, nValor_Minimo_Juridica As Double, nValor_Minimo_FisicaDI As Double, nValor_Minimo_JuridicaDI As Double
Dim nNumproc As Long, nAnoproc As Integer, aTributo() As TRIBUTO, aTributos() As TRIBUTO, bMove As Boolean, X1 As Integer, Y1 As Integer, bRefisAtivo As Boolean, bRefisAtivoDI As Boolean

Private Sub chkCorrecao_Click()
If lvOrigem.ListItems.Count > 0 Then
    CarregaOrigem
    AtualizaTotal
End If
End Sub

Private Sub chkJuros_Click()
If lvOrigem.ListItems.Count > 0 Then
    CarregaOrigem
    AtualizaTotal
End If
End Sub

Private Sub chkMulta_Click()
If lvOrigem.ListItems.Count > 0 Then
    CarregaOrigem
    AtualizaTotal
End If
End Sub

Private Sub chkPenhorado_Click()
AtualizaTotal
End Sub

Private Sub cmbQtde_Click()
Dim x As Integer, Y As Integer, nSomaLiquido As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double, nSomaPrincipal As Double, nSomaJurosMes As Double, nSomaTotal As Double, nSomaTotal2 As Double
Dim nSomaJurosApl As Double, nSomaHonorario As Double
nSomaLiquido = 0: nSomaJuros = 0: nSomaMulta = 0: nSomaCorrecao = 0: nSomaPrincipal = 0: nSomaJurosMes = 0: nSomaTotal = 0: nSomaTotal2 = 0
If Not bExec Then Exit Sub

lvDestino.ListItems.Clear
For x = 1 To UBound(aDestino)
    If aDestino(x).Qtde_Parcela = Val(cmbQtde.Text) Then
        Set itmX = lvDestino.ListItems.Add(, , Format(aDestino(x).Numero_Parcela, "00"))
        itmX.SubItems(1) = Format(aDestino(x).Numero_Parcela, "00")
        itmX.SubItems(2) = aDestino(x).Data_Vencimento
        itmX.SubItems(3) = Format(aDestino(x).Valor_Liquido, "#0.00")
        itmX.SubItems(4) = Format(aDestino(x).Valor_Juros, "#0.00")
        itmX.SubItems(5) = Format(aDestino(x).Valor_Multa, "#0.00")
        itmX.SubItems(6) = Format(aDestino(x).Valor_Correcao, "#0.00")
        itmX.SubItems(7) = Format(aDestino(x).Valor_Principal, "#0.00")
        itmX.SubItems(8) = Format(aDestino(x).Saldo, "#0.00")
        itmX.SubItems(9) = Format(aDestino(x).juros_perc, "#0.00")
        itmX.SubItems(10) = Format(aDestino(x).juros_mes, "#0.00")
        itmX.SubItems(11) = Format(aDestino(x).juros_aplicado, "#0.00")
        itmX.SubItems(12) = Format(aDestino(x).honorario, "#0.00")
        itmX.SubItems(13) = Format(aDestino(x).valor_parcela, "#0.00")
        
        nSomaLiquido = nSomaLiquido + itmX.SubItems(3)
        nSomaJuros = nSomaJuros + itmX.SubItems(4)
        nSomaMulta = nSomaMulta + itmX.SubItems(5)
        nSomaCorrecao = nSomaCorrecao + itmX.SubItems(6)
        nSomaPrincipal = nSomaPrincipal + itmX.SubItems(7)
        nSomaJurosMes = nSomaJurosMes + itmX.SubItems(10)
        nSomaJurosApl = nSomaJurosApl + itmX.SubItems(11)
        nSomaHonorario = nSomaHonorario + itmX.SubItems(12)
        nSomaTotal = nSomaTotal + aDestino(x).valor_parcela
        
    End If
Next

Set itmX = lvDestino.ListItems.Add(, , ">>>>")
itmX.SubItems(2) = "Total --->"
itmX.SubItems(3) = Format(nSomaLiquido, "#0.00")
itmX.SubItems(4) = Format(nSomaJuros, "#0.00")
itmX.SubItems(5) = Format(nSomaMulta, "#0.00")
itmX.SubItems(6) = Format(nSomaCorrecao, "#0.00")
itmX.SubItems(7) = Format(nSomaPrincipal, "#0.00")
itmX.SubItems(8) = "N/A"
itmX.SubItems(9) = "N/A"
itmX.SubItems(10) = "N/A"
itmX.SubItems(11) = Format(nSomaJurosApl, "#0.00")
itmX.SubItems(12) = Format(nSomaHonorario, "#0.00")
itmX.SubItems(13) = Format(nSomaTotal, "#0.00")
For Y = 1 To 13
    itmX.ForeColor = VerdeEscuro
    itmX.ListSubItems(Y).ForeColor = vbBlue
Next Y


Set itmX = lvDestino.ListItems.Add(, , ">>>>")
itmX.SubItems(2) = "% Prop. --->"
itmX.SubItems(3) = Format(nSomaLiquido * 100 / nSomaTotal, "#0.00")
itmX.SubItems(4) = Format(nSomaJuros * 100 / nSomaTotal, "#0.00")
itmX.SubItems(5) = Format(nSomaMulta * 100 / nSomaTotal, "#0.00")
itmX.SubItems(6) = Format(nSomaCorrecao * 100 / nSomaTotal, "#0.00")
itmX.SubItems(7) = "N/A"
itmX.SubItems(8) = "N/A"
itmX.SubItems(9) = "N/A"
itmX.SubItems(10) = "N/A"
itmX.SubItems(11) = Format(nSomaJurosApl * 100 / nSomaTotal, "#0.00")
itmX.SubItems(12) = Format(nSomaHonorario * 100 / nSomaTotal, "#0.00")
itmX.SubItems(13) = Format(nSomaTotal * 100 / nSomaTotal, "#0.00")


For Y = 1 To 13
    itmX.ForeColor = VerdeEscuro
    itmX.ListSubItems(Y).ForeColor = vbBlue
Next Y


End Sub

Private Sub cmdAnalisar_Click()
Dim Sql As String, x As Integer, bAchou As Boolean, bAjS As Boolean, bAjN As Boolean, bDIS As Boolean, bDIN As Boolean, z As Variant, nPlano As Integer, bMultaObraS As Boolean, bMultaObraN As Boolean

bAjS = False: bAjN = False: bDIS = False: bDIN = False

If lvOrigem.ListItems.Count = 0 Then
    MsgBox "Nada a analisar!", vbCritical, "ERRO"
    Exit Sub
End If

bAchou = False
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
    MsgBox "Selecione ao menos uma parcela.", vbCritical, "ERRO"
    Exit Sub
End If


If Not IsDate(mskVencto.Text) Then
    MsgBox "Digite a data do 1º vencimento.", vbCritical, "ERRO"
    Exit Sub
End If

If CDate(mskVencto.Text) < CDate(Format(Now, "dd/mm/yyyy")) Then
    MsgBox "Data do 1º vencimento inválida.", vbCritical, "ERRO"
    Exit Sub
End If


For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked Then
        If lvOrigem.ListItems(x).SubItems(6) = "N" Then
            bAjN = True
        Else
            bAjS = True
        End If
        If Val(Left(lvOrigem.ListItems(x).SubItems(1), 3)) = 81 Then
            bDIS = True
        Else
            bDIN = True
        End If
    End If
Next

If bAjN And bAjS Then
    MsgBox "Não é permitido parcelar débitos ajuizados e não ajuizados no mesmo parcelamento.", vbCritical, "ERRO"
    Exit Sub
End If

If bAjN Then
    lblAj.Caption = "N"
Else
    lblAj.Caption = "S"
End If

If bDIS And bDIN Then
    MsgBox "Não é permitido parcelar débitos de Alienação de Área do Distrito Industrial junto com outros débitos.", vbCritical, "ERRO"
    Exit Sub
End If

If bDIN Then
    lblDI.Caption = "N"
Else
    lblDI.Caption = "S"
End If

'****** REFIS *************
If (bRefisAtivo Or bRefisAtivoDI) And chkRefis.value = vbChecked Then
    If Year(Now) = 2017 Then
        '******** 2017 ********
        If lblDI.Caption = "S" Then
            '** DISTRITO INDUSTRIAL **
            If Val(txtPlano.Tag) = 4 Then
InicioData:
                z = InputBox("O desconto do Refis para o Distrito Industrial é ajustado conforme a quantidade de parcelas." & vbCrLf & "Por favor digite:" & vbCrLf & vbCrLf & "1 - Para parcelar em até 12 vezes." & vbCrLf & "2 - Para parcelar em até 24 vezes.", "Quantidade de Parcelas")
                If Val(z) < 1 Or Val(z) > 2 Then
                    MsgBox "Opção inválida!", vbCritical, "Erro"
                    GoTo InicioData
                Else
                    If Val(z) = 1 Then
                        txtPlano.Text = "REFIS 2017 PARC.DI - 60%"
                        txtPlano.Tag = "24"
                        lblPercDesconto.Caption = "60%"
                    ElseIf Val(z) = 2 Then
                        txtPlano.Text = "REFIS 2017 PARC.DI - 40%"
                        txtPlano.Tag = "25"
                        lblPercDesconto.Caption = "40%"
                    End If
                End If
            End If
        Else
            '*** OUTROS ***
            
            For x = 1 To lvOrigem.ListItems.Count
                If lvOrigem.ListItems(x).Checked And CDate(lvOrigem.ListItems(x).SubItems(5)) > CDate("12/31/2016") Then
                    If Val(Left(lvOrigem.ListItems(x).SubItems(1), 2)) <> 62 Then
                        MsgBox "No refis só podem entrar débitos vencidos até 31/12/2016", vbCritical, "Atenção"
                        Exit Sub
                    End If
                End If
            Next
            
            
            
            If CDate(mskVencto.Text) <= CDate("29/09/2017") Then
                nPlano = 20
            ElseIf CDate(mskVencto.Text) > CDate("29/09/2017") And CDate(mskVencto.Text) <= CDate("31/10/2017") Then
                nPlano = 21
            ElseIf CDate(mskVencto.Text) > CDate("31/10/2017") And CDate(mskVencto.Text) <= CDate("22/12/2017") Then
                nPlano = 22
            End If
                                    
            Sql = "select nome,desconto from plano where codigo=" & nPlano
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            txtPlano.Text = RdoAux2!Nome
            txtPlano.Tag = nPlano
            lblPercDesconto.Caption = RdoAux2!desconto & "%"
            
            nPerc = RdoAux2!desconto
            RdoAux2.Close
          
        End If
        '**********************
    End If
    
End If
'**************************

If chkRefis.value = vbChecked Then
    If bRefisAtivo And lblDI.Caption = "N" Then
        AtualizaRefis
    ElseIf bRefisAtivoDI And lblDI.Caption = "S" Then
        AtualizaRefis
    End If
End If

Inicio:
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked = False Then
        lvOrigem.ListItems.Remove (x)
        GoTo Inicio:
    End If
Next

If chkPenhorado.value = vbChecked Then
    lbl30.Visible = True
Else
    lbl30.Visible = False
End If

LockPanel True
bExec = False
CarregaDestino
AtualizaTributo

End Sub

Private Sub cmdDDList_Click()
Dim nAno As Integer, x As Integer, Y As Integer

If cmdDDList.value = True Then
    frDDList.Height = 2130
Else
    For x = 0 To lstAno.ListCount - 1
        nAno = lstAno.List(x)
        If lstAno.Selected(x) = True Then
            For Y = 1 To lvOrigem.ListItems.Count
                If lvOrigem.ListItems(Y).Text = nAno Then
                    lvOrigem.ListItems(Y).Checked = True
                End If
            Next
        Else
            For Y = 1 To lvOrigem.ListItems.Count
                If lvOrigem.ListItems(Y).Text = nAno Then
                    lvOrigem.ListItems(Y).Checked = False
                End If
            Next
        End If
    Next
    AtualizaTotal
    frDDList.Height = 375
End If

End Sub

Private Sub cmdOpcao_Click()
PopupMenu mnuOpcao
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset, dDataIni As Date, dDataFim As Date
ReDim aTotal(0)
Centraliza Me
CarregaValorMinimo

Sql = "select valparam from parametros where nomeparam='REFIS_INICIO'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataIni = CDate(RdoAux!valparam)

Sql = "select valparam from parametros where nomeparam='REFIS_FIM'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataFim = CDate(RdoAux!valparam)

RdoAux.Close

If Now >= dDataIni And Now <= dDataFim Then
    bRefisAtivo = True
Else
    bRefisAtivo = False
End If

Sql = "select valparam from parametros where nomeparam='REFISDI_INICIO'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataIni = CDate(RdoAux!valparam)

Sql = "select valparam from parametros where nomeparam='REFISDI_FIM'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataFim = CDate(RdoAux!valparam)

RdoAux.Close

If Now >= dDataIni And Now <= dDataFim Then
    bRefisAtivoDI = True
Else
    bRefisAtivoDI = False
End If

Limpa

If bRefisAtivo Then
    chkRefis.value = vbChecked
Else
    chkRefis.value = vbUnchecked
End If

End Sub

Private Sub frDetalhe_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
X1 = x
Y1 = Y
bMove = True

End Sub

Private Sub frDetalhe_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bMove Then
    frDetalhe.Top = frDetalhe.Top + Y - Y1
    frDetalhe.Left = frDetalhe.Left + x - X1
End If
End Sub

Private Sub frDetalhe_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
bMove = False
End Sub

Private Sub grdDestino_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
X1 = x
Y1 = Y
bMove = True
End Sub

Private Sub grdDestino_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bMove Then
    frDetalhe.Top = frDetalhe.Top + Y - Y1
    frDetalhe.Left = frDetalhe.Left + x - X1
End If
End Sub

Private Sub grdDestino_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
bMove = False
End Sub

Private Sub lvOrigem_Click()

If frDDList.Height > 2000 Then
    cmdDDList.value = False
    cmdDDList_Click
End If

End Sub

Private Sub lvOrigem_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If Val(Item.Text) = 0 Then
    Item.Checked = False
    Exit Sub
End If
AtualizaTotal

End Sub

Private Sub mnuCalculo_Click()

Dim nSomaPrincipal As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double, nSomaTotal As Double, sDataProc As String, sExercicio As String, nValorEntrada As Double
Dim nPercPrincipal As Double, nPercJuros As Double, nPercMulta As Double, nPercCorrecao As Double, nDivPrincipal As Double, nDivJuros As Double, nDivMulta As Double, nDivCorrecao As Double
Dim nNumproc As Long, nAnoproc As Integer, nPercDesconto As Double, x As Integer, aAno() As Integer, nNumParcela As Integer, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double
Dim nValorCorrecao As Double, nValorTotal1 As Double, nSaldo As Double, nJurosPerc As Double, nJurosMes As Double, nJurosApl As Double, nHonorarios As Double, nValorParcela As Double, sDataVencto As String
Dim nCalc_Juros As Integer, nCalc_Multa As Integer, nCalc_Correcao As Integer, nCalcHon As Integer, nPenhorado As Integer, nRefis As Integer

If lvDestino.ListItems.Count = 0 Then
    MsgBox "Não existem parcelas a serem calculadas.", vbInformation, "Atenção"
    Exit Sub
End If

Sql = "delete from calculo_parcelamento_origem where usuario='" & NomeDeLogin & "' and codreduzido=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect
Sql = "delete from calculo_parcelamento_destino where usuario='" & NomeDeLogin & "' and codreduzido=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

'***Grava Origem***
If txtNumProc.Text = "" Then
    nNumproc = 0
    nAnoproc = 0
Else
    nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
    nAnoproc = Right$(txtNumProc.Text, 4)
End If

If lblDataProc.Caption = "" Then
    sDataProc = "01/01/1900"
Else
    sDataProc = lblDataProc.Caption
End If


ReDim aAno(0)
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked Then
        bFind = False
        For Y = 0 To UBound(aAno)
            If aAno(Y) = Val(lvOrigem.ListItems(x).Text) Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = Val(lvOrigem.ListItems(x).Text)
        End If
    End If
Next

sExercicio = ""
For x = 1 To UBound(aAno)
    sExercicio = sExercicio & CStr(aAno(x)) & ","
Next
sExercicio = Left(sExercicio, Len(sExercicio) - 1)

nPercDesconto = CDbl(Left(lblPercDesconto.Caption, Len(lblPercDesconto.Caption) - 1))

If lblValorAdicional.Caption = "" Then
    nValorEntrada = 0
Else
    nValorEntrada = CDbl(lblValorAdicional.Caption)
End If

nCalc_Juros = IIf(chkJuros.value = vbChecked, 1, 0)
nCalc_Multa = IIf(chkMulta.value = vbChecked, 1, 0)
nCalc_Correcao = IIf(chkCorrecao.value = vbChecked, 1, 0)
nCalcHon = IIf(chkHonorario.value = vbChecked, 1, 0)
nPenhorado = IIf(chkPenhorado.value = vbChecked, 1, 0)
nRefis = IIf(chkRefis.value = vbChecked, 1, 0)

With lvOrigem
    nSomaPrincipal = CDbl(.ListItems(.ListItems.Count - 2).SubItems(7))
    nSomaJuros = CDbl(.ListItems(.ListItems.Count - 2).SubItems(8))
    nSomaMulta = CDbl(.ListItems(.ListItems.Count - 2).SubItems(9))
    nSomaCorrecao = CDbl(.ListItems(.ListItems.Count - 2).SubItems(10))
    nSomaTotal = CDbl(.ListItems(.ListItems.Count - 2).SubItems(14))
    nPercPrincipal = CDbl(.ListItems(.ListItems.Count - 1).SubItems(7))
    nPercJuros = CDbl(.ListItems(.ListItems.Count - 1).SubItems(8))
    nPercMulta = CDbl(.ListItems(.ListItems.Count - 1).SubItems(9))
    nPercCorrecao = CDbl(.ListItems(.ListItems.Count - 1).SubItems(10))
    nDivPrincipal = CDbl(.ListItems(.ListItems.Count).SubItems(7))
    nDivJuros = CDbl(.ListItems(.ListItems.Count).SubItems(8))
    nDivMulta = CDbl(.ListItems(.ListItems.Count).SubItems(9))
    nDivCorrecao = CDbl(.ListItems(.ListItems.Count).SubItems(10))
    
    
    Sql = "insert into calculo_parcelamento_origem (usuario,codreduzido,nome,ano_processo,numero_processo,exercicios,data_processo,valor_principal,valor_juros,valor_multa,valor_correcao,perc_principal,perc_juros,perc_multa,perc_correcao,"
    Sql = Sql & "valor_total,qtde_parcela,div_principal,div_juros,div_multa,div_correcao,valor_entrada,plano_codigo,plano_descricao,perc_desconto,calculo_juros,calculo_multa,calculo_correcao,calculo_honorario,penhorado,refis,data_calculo)"
    Sql = Sql & "values('" & NomeDeLogin & "'," & Val(txtCod.Text) & ",'" & Mask(txtNome.Text) & "'," & nAnoproc & "," & nNumproc & ",'" & sExercicio & "','"
    Sql = Sql & Format(sDataProc, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nSomaPrincipal)) & "," & Virg2Ponto(CStr(nSomaJuros)) & "," & Virg2Ponto(CStr(nSomaMulta)) & "," & Virg2Ponto(CStr(nSomaCorrecao)) & ","
    Sql = Sql & Virg2Ponto(CStr(nPercPrincipal)) & "," & Virg2Ponto(CStr(nPercJuros)) & "," & Virg2Ponto(CStr(nPercMulta)) & "," & Virg2Ponto(CStr(nPercCorrecao)) & "," & Virg2Ponto(CStr(nSomaTotal)) & ","
    Sql = Sql & Val(cmbQtde.Text) & "," & Virg2Ponto(CStr(nDivPrincipal)) & "," & Virg2Ponto(CStr(nDivJuros)) & "," & Virg2Ponto(CStr(nDivMulta)) & "," & Virg2Ponto(CStr(nDivCorrecao)) & "," & Virg2Ponto(CStr(nValorEntrada)) & ","
    Sql = Sql & Val(txtPlano.Tag) & ",'" & txtPlano.Text & "'," & nPercDesconto & "," & nCalc_Juros & "," & nCalc_Multa & "," & nCalc_Correcao & "," & nCalcHon & "," & nPenhorado & "," & nRefis & ",'" & Format(Now, "mm/dd/yyyy hh:mm") & "')"
    cn.Execute Sql, rdExecDirect
End With

'***Grava Destino***
With lvDestino
    For x = 1 To lvDestino.ListItems.Count - 2
        nNumParcela = .ListItems(x).SubItems(1)
        sDataVencto = .ListItems(x).SubItems(2)
        nValorPrincipal = CDbl(.ListItems(x).SubItems(3))
        nValorJuros = CDbl(.ListItems(x).SubItems(4))
        nValorMulta = CDbl(.ListItems(x).SubItems(5))
        nValorCorrecao = CDbl(.ListItems(x).SubItems(6))
        nValorTotal1 = CDbl(.ListItems(x).SubItems(7))
        nSaldo = CDbl(.ListItems(x).SubItems(8))
        nJurosPerc = CDbl(.ListItems(x).SubItems(9))
        nJurosMes = CDbl(.ListItems(x).SubItems(10))
        nJurosApl = CDbl(.ListItems(x).SubItems(11))
        nHonorarios = CDbl(.ListItems(x).SubItems(12))
        nValorParcela = CDbl(.ListItems(x).SubItems(13))
        
        Sql = "insert into calculo_parcelamento_destino (usuario,codreduzido,numero_parcela,data_vencimento,valor_principal,valor_juros,valor_multa,valor_correcao,valor_total1,"
        Sql = Sql & "saldo,juros_perc,juros_mes,juros_apl,honorarios,valor_parcela) values('" & NomeDeLogin & "'," & Val(txtCod.Text) & "," & nNumParcela & ",'"
        Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPrincipal)) & "," & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorCorrecao)) & "," & Virg2Ponto(CStr(nValorTotal1)) & "," & Virg2Ponto(CStr(nSaldo)) & "," & Virg2Ponto(CStr(nJurosPerc)) & ","
        Sql = Sql & Virg2Ponto(CStr(nJurosMes)) & "," & Virg2Ponto(CStr(nJurosApl)) & "," & Virg2Ponto(CStr(nHonorarios)) & "," & Virg2Ponto(CStr(nValorParcela)) & ")"
        cn.Execute Sql, rdExecDirect
        
    Next
End With

If frmMdi.frTeste.Visible = True Then
    frmReport.ShowReport3 "CALCULO_PARCELAMENTO2_TMP", frmMdi.hwnd, Me.hwnd
Else
    frmReport.ShowReport3 "CALCULO_PARCELAMENTO2", frmMdi.hwnd, Me.hwnd
End If

Sql = "delete from calculo_parcelamento_origem where usuario='" & NomeDeLogin & "' and codreduzido=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect
Sql = "delete from calculo_parcelamento_destino where usuario='" & NomeDeLogin & "' and codreduzido=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mnuCancelar_Click()
'If lvDestino.ListItems.Count = 0 Then
'    LockPanel False
'    MsgBox "Nada a cancelar.", vbCritical, "Atenção"
'    Exit Sub
'End If
If MsgBox("Cancelar operação?", vbCritical + vbYesNo + vbQuestion, "Confirmação") = vbNo Then Exit Sub

LockPanel False
Limpa
txtCod_KeyPress vbKeyReturn
cmbQtde.Clear
txtNumProc.Text = ""
lvDestino.ListItems.Clear

End Sub

Private Sub mnuCriar_Click()
Dim sNumProc As String, Sql As String, RdoAux As rdoResultset

If lvDestino.ListItems.Count = 0 Then
    MsgBox "Não existe nada a ser parcelado.", vbCritical, "ERRO"
    Exit Sub
End If

If txtNumProc.Text = "" Then
    MsgBox "Digite o número do Processo.", vbCritical, "Atenção"
    Exit Sub
End If

If nNumproc = 0 Then
    MsgBox "Processo inválido.", vbCritical, "Atenção"
    Exit Sub
End If

If CDate(mskVencto.Text) < CDate(Format(Now, "dd/mm/yyyy")) Then
    MsgBox "Data do 1º vencimento tem que ser maior ou igua a data atual.", vbCritical, "Atenção"
    Exit Sub
End If

sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
'VERIFICA SE O PROCESSO JA FOI UTILIZADO
Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & " '"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este processo já foi utilizado em um parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Deseja criar este parcelamento com " & Val(cmbQtde.Text) & " parcelas?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    FillDetalhe
    mnuCalculo_Click
    EmiteBoleto
End If
End Sub

Private Sub mnuDetalhe_Click()

If lvDestino.ListItems.Count = 0 Then
    MsgBox "Não existe nada a ser exibido.", vbCritical, "ERRO"
    Exit Sub
End If

If mnuDetalhe.Checked = False Then
    FillDetalhe
    frDetalhe.Visible = True
    frDetalhe.ZOrder 0
    mnuDetalhe.Checked = True
Else
    frDetalhe.Visible = False
    mnuDetalhe.Checked = False
End If
End Sub

Private Sub mnuSimulado_Click()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, Y As Integer, nValorEntrada As Double, nValorParcela As Double, sExercicio As String
Dim aAno() As Integer, bFind As Boolean, xImovel As clsImovel, nCodReduz As Long, sEndereco As String, nNumero As Integer, sComplemento As String
Dim sBairro As String, sCidade As String, sCep As String, sUF As String, sFullEndereco As String

nCodReduz = Val(txtCod.Text)
If lvDestino.ListItems.Count = 0 Then
    MsgBox "Não existem parcelas a serem simuladas.", vbInformation, "Atenção"
    Exit Sub
End If

ReDim aAno(0)
For x = 1 To lvOrigem.ListItems.Count
    If lvOrigem.ListItems(x).Checked Then
        bFind = False
        For Y = 0 To UBound(aAno)
            If aAno(Y) = Val(lvOrigem.ListItems(x).Text) Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = Val(lvOrigem.ListItems(x).Text)
        End If
    End If
Next

If nCodReduz < 100000 Then
    Set xImovel = New clsImovel
    xImovel.CarregaImovel nCodReduz
    sEndereco = xImovel.EnderecoCompleto2
    nNumero = Val(xImovel.Li_Num)
    sComplemento = xImovel.Li_Compl
    sBairro = xImovel.DescBairro
    sCidade = "JABOTICABAL"
    sCep = RetornaCEP(xImovel.CodLogr, nNumero)
    sUF = "SP"
ElseIf nCodReduz > 100000 And nCodReduz < 300000 Then
    Sql = "SELECT * FROM vwFULLEMPRESA3 WHERE CODIGOMOB=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        sEndereco = !Logradouro
        nNumero = !Numero
        sComplemento = SubNull(!Complemento)
        sBairro = SubNull(!DescBairro)
        sCidade = SubNull(!descCidade)
        sUF = SubNull(!SiglaUF)
        sCep = SubNull(!Cep)
     End With
Else
    Sql = "SELECT * from vwFULLCIDADAO WHERE CODCIDADAO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        sEndereco = SubNull(!Endereco)
        nNumero = Val(SubNull(!NUMIMOVEL))
        sComplemento = SubNull(!Complemento)
        sBairro = SubNull(!DescBairro)
        sCidade = SubNull(!descCidade)
        sUF = SubNull(!SiglaUF)
        sCep = SubNull(!Cep)
    End With
End If

sFullEndereco = sEndereco & " - " & sBairro & " - " & sCidade & "/" & sUF & " " & sCep

sExercicio = ""
For x = 1 To UBound(aAno)
    sExercicio = sExercicio & CStr(aAno(x)) & ","
Next
sExercicio = Left(sExercicio, Len(sExercicio) - 1)

Sql = "delete from parcelamento_simulado where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


For x = Val(cmbQtde.List(0)) To Val(cmbQtde.List(cmbQtde.ListCount - 1))
    For Y = 1 To UBound(aDestino)
        If aDestino(Y).Qtde_Parcela = x And aDestino(Y).Numero_Parcela = 1 Then
            nValorEntrada = aDestino(Y).valor_parcela
        ElseIf aDestino(Y).Qtde_Parcela = x And aDestino(Y).Numero_Parcela = 2 Then
            nValorParcela = aDestino(Y).valor_parcela
        End If
    Next
    Sql = "insert parcelamento_simulado(usuario,codigo,qtde_parcela,nome,documento,exercicios,valor_entrada,valor_parcela,endereco) values('"
    Sql = Sql & NomeDeLogin & "'," & Val(txtCod.Text) & "," & x & ",'" & Left(Mask(txtNome.Text), 50) & "','" & txtDoc.Text & "','" & sExercicio & "',"
    Sql = Sql & Virg2Ponto(CStr(nValorEntrada)) & "," & Virg2Ponto(CStr(nValorParcela)) & ",'" & Mask(sFullEndereco) & "')"
    cn.Execute Sql, rdExecDirect
Next

If frmMdi.frTeste.Visible = True Then
    frmReport.ShowReport3 "PARCELAMENTO_SIMULADO_TMP", frmMdi.hwnd, Me.hwnd
Else
    frmReport.ShowReport3 "PARCELAMENTO_SIMULADO", frmMdi.hwnd, Me.hwnd
End If

Sql = "delete from parcelamento_simulado where usuario='" & NomeDeLogin & "' and codigo=" & Val(txtCod.Text)
cn.Execute Sql, rdExecDirect

End Sub

Private Sub txtCod_Change()
Limpa
End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    CarregaContribuinte Val(txtCod.Text)
    If txtNome.Text <> "" Then
        If txtDoc.Text <> "" Then
            CarregaOrigem
            If lvOrigem.ListItems.Count = 0 Then
                MsgBox "Não existem débitos a serem parcelados.", vbInformation, "Atenção"
            End If
        Else
            MsgBox "O contribuinte deve possuir CPF ou CNPJ cadastrado.", vbCritical, "ERRO"
            Limpa
        End If
    End If
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If
End Sub

Private Sub CarregaContribuinte(Codigo As Long)
Dim Sql As String, RdoAux As rdoResultset, sDoc As String

txtDoc.Text = ""
txtNome.Text = ""
sDoc = ""

If Codigo = 0 Or Codigo > 700000 Then
    MsgBox "Digite um código válido de contribuinte.", vbCritical, "Erro!"
    Exit Sub
End If

If Codigo < 100000 Then
    Sql = "select nomecidadao,cpf,cnpj from vwfullimovel where codreduzido=" & Codigo
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
           .Close
            GoTo SemCadastro
        Else
            txtNome.Text = RdoAux!nomecidadao
            If SubNull(!CPF) <> "" Then
                sDoc = Format(!CPF, "000\.000\.000-00")
            Else
                If SubNull(!Cnpj) <> "" Then
                    sDoc = Format(!Cnpj, "00\.000\.000/0000-00")
                End If
            End If
        End If
       .Close
    End With
ElseIf Codigo >= 100000 And Codigo < 500000 Then
    Sql = "select razaosocial,cpf,cnpj from mobiliario where codigomob=" & Codigo
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
           .Close
            GoTo SemCadastro
        Else
            txtNome.Text = RdoAux!RazaoSocial
            If SubNull(!CPF) <> "" Then
                sDoc = Format(!CPF, "000\.000\.000-00")
            Else
                If SubNull(!Cnpj) <> "" Then
                    sDoc = Format(!Cnpj, "00\.000\.000/0000-00")
                End If
            End If
        End If
       .Close
    End With
ElseIf Codigo >= 500000 Then
    Sql = "select nomecidadao,cpf,cnpj from cidadao where codcidadao=" & Codigo
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
           .Close
            GoTo SemCadastro
        Else
            txtNome.Text = RdoAux!nomecidadao
            If SubNull(!CPF) <> "" Then
                sDoc = Format(!CPF, "000\.000\.000-00")
            Else
                If SubNull(!Cnpj) <> "" Then
                    sDoc = Format(!Cnpj, "00\.000\.000/0000-00")
                End If
            End If
        End If
       .Close
    End With
End If

txtDoc.Text = sDoc
Exit Sub

SemCadastro:
MsgBox "Contribuinte não cadastrado", vbCritical, "Erro!"

End Sub

Private Sub txtCod_LostFocus()
'txtCod_KeyPress vbKeyReturn

'If Val(txtCod.Text) > 0 Then
'    CarregaContribuinte Val(txtCod.Text)
'End If

End Sub

Private Sub Limpa()

'txtNome.Text = ""
txtDoc.Text = ""
lvOrigem.ListItems.Clear
lvDestino.ListItems.Clear
mskVencto.Text = Format(Now, "dd/mm/yyyy")
txtPlano.Text = ""
txtPlano.Text = "(SEM PLANO)"
txtPlano.Tag = "4"
lblPerc.Caption = "0,00%"
lblPercDesconto.Caption = "0,00%"
lblAj.Caption = "N"
lblDI.Caption = "N"
lstAno.Clear
cmbQtde.Clear
txtNumProc.Text = ""
chkMulta.value = vbChecked
chkJuros.value = vbChecked
chkCorrecao.value = vbChecked
chkHonorario.value = vbChecked
chkPenhorado.value = vbUnchecked
lblValorAdicional.Caption = "0,00"
lbl30.Visible = False
grdTributo.Rows = 1
grdDestino.Rows = 1
frDetalhe.Visible = False
mnuDetalhe.Checked = False
lblP.Caption = "0,00": lblJ.Caption = "0,00": lblM.Caption = "0,00": lblC.Caption = "0,00": lblT.Caption = "0,00"

End Sub

Private Sub CarregaOrigem()
Dim RdoAux As rdoResultset, Sql As String, qd As New rdoQuery, t As Integer, Achou As Boolean, nPlano As Integer

lvOrigem.ListItems.Clear
lvDestino.ListItems.Clear
ShowWait True

cn.QueryTimeout = 180
Set qd.ActiveConnection = cn

On Error Resume Next
RdoAux3.Close
On Error GoTo 0
If NomeDeLogin = "SCHWARTZ2" Then
    qd.Sql = "{ Call spParcelamentoOrigem2(?,?,?,?,?)}"
Else
    qd.Sql = "{ Call spParcelamentoOrigem(?,?,?,?,?)}"
End If
qd(0) = Val(txtCod.Text)
qd(1) = IIf(Len(RetornaNumero(txtDoc.Text)) = 11, "F", "J")
qd(2) = IIf(chkJuros.value = vbChecked, 1, 0)
qd(3) = IIf(chkMulta.value = vbChecked, 1, 0)
qd(4) = IIf(chkCorrecao.value = vbChecked, 1, 0)
Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
With RdoAux
    Do Until .EOF
    
        'If !Data_Vencimento >= CDate(Format(Now, "dd/mm/yyyy")) And !lancamento <> 69 And !lancamento <> 10 And !lancamento <> 48 And !lancamento <> 76 And !lancamento <> 65 And !lancamento <> 62 Then
        If !Data_Vencimento >= CDate(Format(Now, "dd/mm/yyyy")) And !lancamento <> 65 And !lancamento <> 62 And !lancamento <> 16 Then
            GoTo proximo
        End If
    
        Achou = False
        For t = 0 To lstAno.ListCount - 1
            If lstAno.List(t) = !exercicio Then
                Achou = True
                Exit For
            End If
        Next
        If Not Achou Then
            lstAno.AddItem !exercicio
        End If
    
        Set itmX = lvOrigem.ListItems.Add(, , !exercicio)
        itmX.SubItems(1) = Format(!lancamento, "00") & "-" & !nome_lancamento
        itmX.SubItems(2) = Format(!Sequencia, "00")
        itmX.SubItems(3) = Format(!Parcela, "00")
        itmX.SubItems(4) = Format(!Complemento, "00")
        itmX.SubItems(5) = Format(!Data_Vencimento, "dd/mm/yyyy")
        itmX.SubItems(6) = !ajuizado
        itmX.SubItems(7) = FormatNumber(!Valor_Principal, 2)
        itmX.SubItems(8) = FormatNumber(!Valor_Juros, 2)
        itmX.SubItems(9) = FormatNumber(!Valor_Multa, 2)
        itmX.SubItems(10) = FormatNumber(!Valor_Correcao, 2)
        itmX.SubItems(11) = !qtde_parcelamento
        itmX.SubItems(12) = !perc_penalidade & "%"
        itmX.SubItems(13) = FormatNumber(!valor_penalidade, 2)
        itmX.SubItems(14) = FormatNumber(!Valor_Total, 2)
        
proximo:
       .MoveNext
    Loop
   .Close
End With

ShowWait False

End Sub

Private Sub ShowWait(bShow As Boolean)

If bShow Then
    frWait.Visible = True
    frWait.ZOrder 0
    Ocupado
Else
    frWait.Visible = False
    Liberado
End If
Me.Refresh
DoEvents

End Sub

Private Sub AtualizaTotal()
Dim nValorAdicional As Double, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double
Dim nValorCorrecao As Double, nValorTotal As Double, nValor30 As Double, nLinhas As Integer

nValorAdicional = 0
With lvOrigem
    For x = 1 To .ListItems.Count
        'If x < lvOrigem.ListItems.Count And .ListItems(x).Checked = True Then
        If .ListItems(x).Checked = True And .ListItems(x).Text <> "" Then
            nValorPrincipal = nValorPrincipal + CDbl(.ListItems(x).ListSubItems(8))
            nValorJuros = nValorJuros + CDbl(.ListItems(x).ListSubItems(9))
            nValorMulta = nValorMulta + CDbl(.ListItems(x).ListSubItems(10))
            nValorCorrecao = nValorCorrecao + CDbl(.ListItems(x).ListSubItems(11))
            nValorAdicional = nValorAdicional + CDbl(.ListItems(x).ListSubItems(13))
            nValorTotal = nValorTotal + CDbl(.ListItems(x).ListSubItems(14))
        End If
    Next
End With

With aTotal(0)
    .Valor_Principal = nValorPrincipal
    .Valor_Juros = nValorJuros
    .Valor_Multa = nValorMulta
    .Valor_Correcao = nValorCorrecao
    .Valor_Adicional = nValorAdicional
End With

If chkPenhorado.value = vbChecked Then
    nValor30 = nValorTotal * 0.3
    lbl30.Caption = "30% de " & FormatNumber(nValorTotal, 2) & " = " & FormatNumber(nValor30, 2)
    lbl30.Visible = True
Else
    nValor30 = 0
    lbl30.Visible = False
End If

lblValorAdicional.Caption = FormatNumber(nValorAdicional + nValor30, 2)
End Sub

Private Sub LockPanel(bLock As Boolean)

cmdDDList.Enabled = Not bLock
frTop.Enabled = Not bLock

End Sub

Private Sub txtNumProc_LostFocus()
Dim sValidaProc As String, sNumProc As String
On Error Resume Next
If Trim$(txtNumProc.Text) <> "" Then
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
        nAnoproc = Right$(txtNumProc.Text, 4)
        sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
        sValidaProc = ValidaProcesso(sNumProc)
        If sValidaProc <> "OK" Then
            MsgBox sValidaProc, vbCritical, "Atenção"
            lblDataProc.Caption = ""
            nNumproc = 0
            nAnoproc = 0
            Exit Sub
        Else
            lblDataProc.Caption = Format(RetornaDataProcesso(nNumproc, nAnoproc), "dd/mm/yyyy")
        End If
    Else
        MsgBox "Processo inválido.", vbExclamation, "Atenção"
        nNumproc = 0
        nAnoproc = 0
        lblDataProc.Caption = ""
    End If
End If

End Sub

Private Sub CarregaDestino()

Dim RdoAux As rdoResultset, Sql As String, qd As New rdoQuery, t As Integer, bFind As Boolean, bAjuizado As Boolean, c As Integer, nQtdeParc As Integer
Dim nSomaPrincipal As Double, nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double, nTotal As Double, sTipoPessoa As String
Dim lvSI As ListSubItem, nPlano As Integer, nValorMinF As Double, nValorMinJ As Double

nPlano = 0

If Len(txtDoc.Text) = 14 Then
    sTipoPessoa = "F"
Else
    sTipoPessoa = "J"
End If

If Val(txtCod.Text) > 100000 And Val(txtCod.Text) < 3000000 Then
    'empresas encerradas e suspensas são tratadas como empresas físicas
    Sql = "select codigomob,dataencerramento from mobiliario where codigomob=" & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If Not IsNull(RdoAux!dataencerramento) Then
        sTipoPessoa = "F"
    End If
    RdoAux.Close
    
    Sql = "SELECT CODTIPOEVENTO,DATAEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & Val(txtCod.Text) & " ORDER BY DATAEVENTO DESC"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If !CODTIPOEVENTO = 2 Then
                sTipoPessoa = "F"
            End If
        End If
       .Close
    End With
    
    
End If



ShowWait True
ReDim aDestino(0)

If chkRefis.value = vbChecked Then
    CarregaValorMinimo
Else
    nValor_Minimo_Fisica = 50
    nValor_Minimo_Juridica = 200
End If

nSomaPrincipal = 0: nSomaJuros = 0: nSomaMulta = 0: nSomaCorrecao = 0: nTotal = 0

'Agrupa totais
With lvOrigem
    For t = 1 To .ListItems.Count
        nSomaPrincipal = nSomaPrincipal + CDbl(.ListItems(t).SubItems(7))
        nSomaJuros = nSomaJuros + CDbl(.ListItems(t).SubItems(8))
        nSomaMulta = nSomaMulta + CDbl(.ListItems(t).SubItems(9))
        nSomaCorrecao = nSomaCorrecao + CDbl(.ListItems(t).SubItems(10))
        nTotal = nTotal + CDbl(.ListItems(t).SubItems(14))
    Next
End With

Set itmX = lvOrigem.ListItems.Add(, , "")
itmX.SubItems(1) = "Valor Total  --->"
itmX.SubItems(2) = ""
itmX.SubItems(3) = ""
itmX.SubItems(4) = ""
itmX.SubItems(5) = ""
itmX.SubItems(6) = ""
itmX.SubItems(7) = Format(nSomaPrincipal, "#0.00")
itmX.SubItems(8) = Format(nSomaJuros, "#0.00")
itmX.SubItems(9) = Format(nSomaMulta, "#0.00")
itmX.SubItems(10) = Format(nSomaCorrecao, "#0.00")
itmX.SubItems(11) = "---"
itmX.SubItems(12) = "---"
itmX.SubItems(13) = lblValorAdicional.Caption
itmX.SubItems(14) = Format(nSomaPrincipal + nSomaJuros + nSomaMulta + nSomaCorrecao, "#0.00")
For Y = 1 To 14
    itmX.ForeColor = VerdeEscuro
    itmX.ListSubItems(Y).ForeColor = vbBlue
Next Y

Set itmX = lvOrigem.ListItems.Add(, , "")
itmX.SubItems(1) = "% do Total  --->"
itmX.SubItems(2) = ""
itmX.SubItems(3) = ""
itmX.SubItems(4) = ""
itmX.SubItems(5) = ""
itmX.SubItems(6) = ""
itmX.SubItems(7) = Format(nSomaPrincipal * 100 / nTotal, "#0.00")
itmX.SubItems(8) = Format(nSomaJuros * 100 / nTotal, "#0.00")
itmX.SubItems(9) = Format(nSomaMulta * 100 / nTotal, "#0.00")
itmX.SubItems(10) = Format(nSomaCorrecao * 100 / nTotal, "#0.00")
itmX.SubItems(11) = "---"
itmX.SubItems(12) = "---"
itmX.SubItems(13) = "---"
itmX.SubItems(14) = Format(nTotal * 100 / nTotal, "#0.00")

For Y = 1 To 14
    itmX.ForeColor = VerdeEscuro
    itmX.ListSubItems(Y).ForeColor = vbBlue
Next Y

Set itmX = lvOrigem.ListItems.Add(, , "")
itmX.SubItems(1) = "Div.Valor Adic --->"
itmX.SubItems(2) = ""
itmX.SubItems(3) = ""
itmX.SubItems(4) = ""
itmX.SubItems(5) = ""
itmX.SubItems(6) = ""
itmX.SubItems(7) = Format(CDbl(lvOrigem.ListItems(lvOrigem.ListItems.Count - 1).SubItems(7)) * CDbl(lblValorAdicional.Caption) / 100, "#0.00")
itmX.SubItems(8) = Format(CDbl(lvOrigem.ListItems(lvOrigem.ListItems.Count - 1).SubItems(8)) * CDbl(lblValorAdicional.Caption) / 100, "#0.00")
itmX.SubItems(9) = Format(CDbl(lvOrigem.ListItems(lvOrigem.ListItems.Count - 1).SubItems(9)) * CDbl(lblValorAdicional.Caption) / 100, "#0.00")
itmX.SubItems(10) = Format(CDbl(lvOrigem.ListItems(lvOrigem.ListItems.Count - 1).SubItems(10)) * CDbl(lblValorAdicional.Caption) / 100, "#0.00")
itmX.SubItems(11) = "---"
itmX.SubItems(12) = "---"
itmX.SubItems(13) = "---"
itmX.SubItems(14) = Format(CDbl(lvOrigem.ListItems(lvOrigem.ListItems.Count - 1).SubItems(14)) * CDbl(lblValorAdicional.Caption) / 100, "#0.00")

For Y = 1 To 14
    itmX.ForeColor = VerdeEscuro
    itmX.ListSubItems(Y).ForeColor = vbBlue
Next Y


nPlano = Val(txtPlano.Tag)

If chkRefis.value = vbChecked Then
    nValorMinF = 100
    nValorMinJ = 250
Else
    Sql = "SELECT valor From parcelamento_valor_minimo WHERE distritoindustrial = 0 AND tipo = 'F' AND ano = " & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nValorMinF = RdoAux!Valor
    RdoAux.Close
    
    Sql = "SELECT valor From parcelamento_valor_minimo WHERE distritoindustrial = 0 AND tipo = 'J' AND ano = " & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nValorMinJ = RdoAux!Valor
    RdoAux.Close
End If

Set qd.ActiveConnection = cn
On Error Resume Next
RdoAux3.Close
On Error GoTo 0


qd.Sql = "{ Call spParcelamentoDestino(?,?,?,?,?,?,?,?,?,?,?)}"
qd(0) = nPlano
qd(1) = Format(mskVencto.Text, "mm/dd/yyyy")
qd(2) = IIf(lblAj.Caption = "S", 1, 0)
qd(3) = IIf(chkHonorario.value = vbChecked, 1, 0)
qd(4) = nSomaPrincipal
qd(5) = nSomaJuros
qd(6) = nSomaMulta
qd(7) = nSomaCorrecao
qd(8) = nTotal
qd(9) = CDbl(lblValorAdicional.Caption)
If lblDI.Caption = "S" Then
    qd(10) = IIf(sTipoPessoa = "F", nValor_Minimo_FisicaDI, nValor_Minimo_JuridicaDI)
Else
    qd(10) = IIf(sTipoPessoa = "F", nValorMinF, nValorMinJ)
End If

Set RdoAux = qd.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aDestino(UBound(aDestino) + 1)
        aDestino(UBound(aDestino)).Qtde_Parcela = !Qtde_Parcela
        aDestino(UBound(aDestino)).Data_Vencimento = Format(!Data_Vencimento, "dd/mm/yyyy")
        aDestino(UBound(aDestino)).Numero_Parcela = !Numero_Parcela
        aDestino(UBound(aDestino)).Valor_Liquido = !Valor_Liquido
        aDestino(UBound(aDestino)).Valor_Juros = !Valor_Juros
        aDestino(UBound(aDestino)).Valor_Multa = !Valor_Multa
        aDestino(UBound(aDestino)).Valor_Correcao = !Valor_Correcao
        aDestino(UBound(aDestino)).Valor_Principal = !Valor_Principal
        aDestino(UBound(aDestino)).Saldo = !Saldo
        aDestino(UBound(aDestino)).juros_perc = !juros_perc
        aDestino(UBound(aDestino)).juros_mes = !juros_mes
        aDestino(UBound(aDestino)).juros_aplicado = !juros_aplicado
        aDestino(UBound(aDestino)).honorario = !honorario
        aDestino(UBound(aDestino)).valor_parcela = !valor_parcela
       .MoveNext
    Loop
   .Close
End With

'preenche lista das qtdes de parcela
For t = 1 To UBound(aDestino)
    nQtdeParc = aDestino(t).Qtde_Parcela
    bFind = False
    For c = 0 To cmbQtde.ListCount - 1
        If nQtdeParc = Val(cmbQtde.List(c)) Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        cmbQtde.AddItem Format(nQtdeParc, "00")
    End If
Next

bExec = True
If cmbQtde.ListCount > 0 Then
    cmbQtde.ListIndex = 0
End If

ShowWait False

If lvDestino.ListItems.Count = 0 Then
    If sTipoPessoa = "F" Then
        If lblDI.Caption = "S" Then
            MsgBox "As parcelas selecionadas não alcançaram o valor mínimo de R$" & FormatNumber(nValor_Minimo_FisicaDI, 2) & ".", vbCritical, "ERRO"
        Else
            MsgBox "As parcelas selecionadas não alcançaram o valor mínimo de R$" & FormatNumber(nValor_Minimo_Fisica, 2) & ".", vbCritical, "ERRO"
        End If
    Else
        If lblDI.Caption = "S" Then
            MsgBox "As parcelas selecionadas não alcançaram o valor mínimo de R$" & FormatNumber(nValor_Minimo_JuridicaDI, 2) & ".", vbCritical, "ERRO"
        Else
            MsgBox "As parcelas selecionadas não alcançaram o valor mínimo de R$" & FormatNumber(nValor_Minimo_Juridica, 2) & ".", vbCritical, "ERRO"
        End If
    End If
End If

End Sub

Private Sub CarregaValorMinimo()
Dim Sql As String, RdoAux As rdoResultset

Sql = "select valor from parcelamento_valor_minimo where ano=" & Year(Now) & " and distritoindustrial=0 and tipo='F'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValor_Minimo_Fisica = FormatNumber(RdoAux!Valor, 2)
RdoAux.Close

Sql = "select valor from parcelamento_valor_minimo where ano=" & Year(Now) & " and distritoindustrial=0 and tipo='J'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValor_Minimo_Juridica = FormatNumber(RdoAux!Valor, 2)
RdoAux.Close



Sql = "select valor from parcelamento_valor_minimo where ano=" & Year(Now) & " and distritoindustrial=1 and tipo='F'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValor_Minimo_FisicaDI = FormatNumber(RdoAux!Valor, 2)
RdoAux.Close

Sql = "select valor from parcelamento_valor_minimo where ano=" & Year(Now) & " and distritoindustrial=1 and tipo='J'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValor_Minimo_JuridicaDI = FormatNumber(RdoAux!Valor, 2)
RdoAux.Close

End Sub

Private Sub grdTributo_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
X1 = x
Y1 = Y
bMove = True
End Sub

Private Sub grdTributo_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If bMove Then
    frDetalhe.Top = frDetalhe.Top + Y - Y1
    frDetalhe.Left = frDetalhe.Left + x - X1
End If
End Sub

Private Sub grdTributo_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
bMove = False
End Sub

Private Sub AtualizaTributo()
Dim qd As New rdoQuery, RdoAux3 As rdoResultset, nValorTributo As Double, nValorMulta As Double, nValorJuros As Double, nValorCorrecao As Double, nTotal As Double
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
                    For Y = 1 To grdTributo.Rows - 1
                        If Val(grdTributo.TextMatrix(Y, 0)) = !CodTributo Then
                            bAchou = True
                            Exit For
                        End If
                    Next
                    
                    nValorTributo = !ValorTributo
                    nValorJuros = !ValorJuros
                    nValorMulta = !ValorMulta
                    nValorCorrecao = !ValorCorrecao
                    
                    If chkJuros.value = vbUnchecked Then nValorJuros = 0
                    If chkMulta.value = vbUnchecked Then nValorMulta = 0
                    If chkCorrecao.value = vbUnchecked Then nValorCorrecao = 0
                    
                    
                    nTotal = nValorTributo + nValorJuros + nValorMulta + nValorCorrecao
                    If bAchou Then
                        grdTributo.TextMatrix(Y, 2) = Format(CDbl(grdTributo.TextMatrix(Y, 2)) + nValorTributo, "#0.00")
                        grdTributo.TextMatrix(Y, 3) = Format(CDbl(grdTributo.TextMatrix(Y, 3)) + nValorJuros, "#0.00")
                        grdTributo.TextMatrix(Y, 4) = Format(CDbl(grdTributo.TextMatrix(Y, 4)) + nValorMulta, "#0.00")
                        grdTributo.TextMatrix(Y, 5) = Format(CDbl(grdTributo.TextMatrix(Y, 5)) + nValorCorrecao, "#0.00")
                        grdTributo.TextMatrix(Y, 6) = Format(CDbl(grdTributo.TextMatrix(Y, 6)) + nTotal, "#0.00")
                    Else
                        grdTributo.AddItem !CodTributo & Chr(9) & RdoAux!desctributo & Chr(9) & Format(nValorTributo, "#0.00") & Chr(9) & Format(nValorJuros, "#0.00") & Chr(9) & Format(nValorMulta, "#0.00") & Chr(9) & Format(nValorCorrecao, "#0.00") & Chr(9) & Format(nTotal, "#0.00")
                    End If
                    RdoAux.Close
                   .MoveNext
                Loop
               .Close
            End With
                
            lblP.Caption = "0,00": lblJ.Caption = "0,00": lblM.Caption = "0,00": lblC.Caption = "0,00": lblT.Caption = "0,00"
            For Y = 1 To grdTributo.Rows - 1
                lblP.Caption = Format(CDbl(lblP.Caption) + CDbl(grdTributo.TextMatrix(Y, 2)), "#0.00")
                lblJ.Caption = Format(CDbl(lblJ.Caption) + CDbl(grdTributo.TextMatrix(Y, 3)), "#0.00")
                lblM.Caption = Format(CDbl(lblM.Caption) + CDbl(grdTributo.TextMatrix(Y, 4)), "#0.00")
                lblC.Caption = Format(CDbl(lblC.Caption) + CDbl(grdTributo.TextMatrix(Y, 5)), "#0.00")
                lblT.Caption = Format(CDbl(lblT.Caption) + CDbl(grdTributo.TextMatrix(Y, 6)), "#0.00")
            Next
            For Y = 1 To grdTributo.Rows - 1
                grdTributo.TextMatrix(Y, 7) = Format(CDbl(grdTributo.TextMatrix(Y, 6)) * 100 / CDbl(lblT.Caption), "#0.00")
            Next
        End If
    Next
End With

End Sub

Private Sub EmiteBoleto()
Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, Y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String
Dim dDataBase As Date, sTipoEnd As String, bBoleto As Boolean, nValorDif As Double
Dim sValor As String, dDataVencto As Date, nLastCod As Long

dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
nPlano = 0

'BUSCA ULTIMA SEQUENCIA
Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If !contador > 0 Then
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO<100"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nSeq = RdoAux!SEQMAXIMA + 1
    Else
        nSeq = 0
    End If
   .Close
End With

'GRAVA O PROCESSO
'Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,NUMPROC,ANOPROC,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
'Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CALCULACORRECAO,PENHORA,HONORARIO,CODIGORESP,FUNCIONARIO,PLANO,NOVO) VALUES('"
'Sql = Sql & sNumProc & "'," & nNumproc & "," & nAnoproc & ",'" & Format(lblDataProc.Caption, "mm/dd/yyyy") & "','" & Format(dDataBase, "mm/dd/yyyy") & "',"
'Sql = Sql & Val(cmbQtde.Text) & "," & 0 & "," & 0 & "," & IIf(chkMulta.value = vbChecked, 1, 0) & "," & IIf(chkJuros.value = vbChecked, 1, 0) & ","
'Sql = Sql & IIf(chkCorrecao.value = vbChecked, 1, 0) & "," & IIf(chkPenhorado.value = vbChecked, 1, 0) & "," & IIf(chkHonorario.value = vbChecked, 1, 0) & ","
'Sql = Sql & Val(txtCod.Text) & ",'" & NomeDeLogin & "'," & nPlano & "," & 1 & ")"
Sql = "INSERT PROCESSOREPARC (NUMPROCESSO,NUMPROC,ANOPROC,DATAPROCESSO,DATAREPARC,QTDEPARCELA,VALORENTRADA,"
Sql = Sql & "PERCENTRADA,CALCULAMULTA,CALCULAJUROS,CALCULACORRECAO,PENHORA,HONORARIO,CODIGORESP,userid,PLANO,NOVO) VALUES('"
Sql = Sql & sNumProc & "'," & nNumproc & "," & nAnoproc & ",'" & Format(lblDataProc.Caption, "mm/dd/yyyy") & "','" & Format(dDataBase, "mm/dd/yyyy") & "',"
Sql = Sql & Val(cmbQtde.Text) & "," & 0 & "," & 0 & "," & IIf(chkMulta.value = vbChecked, 1, 0) & "," & IIf(chkJuros.value = vbChecked, 1, 0) & ","
Sql = Sql & IIf(chkCorrecao.value = vbChecked, 1, 0) & "," & IIf(chkPenhorado.value = vbChecked, 1, 0) & "," & IIf(chkHonorario.value = vbChecked, 1, 0) & ","
Sql = Sql & Val(txtCod.Text) & "," & RetornaUsuarioID(NomeDeLogin) & "," & nPlano & "," & 1 & ")"
cn.Execute Sql, rdExecDirect
    
'GRAVA AS PARCELAS DE DESTINO
With lvDestino
    For x = 1 To .ListItems.Count - 2
        sDataVencto = .ListItems(x).SubItems(2)
        nAno = Val(Right$(sDataVencto, 4))
        nLanc = 20
        nParc = .ListItems(x).Text
        nCompl = 0
       'GRAVA DESTINOREPARC
        Sql = "INSERT DESTINOREPARC (NUMPROCESSO,ANOPROC,NUMPROC,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,NUMSEQUENCIA,NUMPARCELA,CODCOMPLEMENTO,VALORLIQUIDO,"
        Sql = Sql & "JUROS,MULTA,CORRECAO,VALORPRINCIPAL,SALDO,JUROSPERC,JUROSVALOR,JUROSAPL,HONORARIO,TOTAL) VALUES('"
        Sql = Sql & sNumProc & "'," & nAnoproc & "," & nNumproc & "," & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(3)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(4)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(5)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(6)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(7)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(8)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(9)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(10)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(11)))) & "," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(12)))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(13)))) & ")"
        cn.Execute Sql, rdExecDirect
        
       'GRAVA DEBITOPARCELA
        If nAno = Year(dDataBase) Then
            nSituacao = 3
        Else
            nSituacao = 18
        End If
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
'        Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
'        Sql = Sql & nSituacao & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(dDataBase, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
        Sql = Sql & nSituacao & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(dDataBase, "mm/dd/yyyy") & "'," & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect

       'DEFINE TRIBUTO PRINCIPAL
        With grdDestino
            For Y = 1 To .Rows - 2
                nCodTrib = Val(.TextMatrix(Y, 0))
                If x = 1 Then
                    If .TextMatrix(Y, 6) = "" Then
                        nValorDif = 0
                    Else
                        nValorDif = CDbl(.TextMatrix(Y, 6))
                    End If
                    nValorTributo = CDbl(.TextMatrix(Y, 5)) + nValorDif
                Else
                    If .TextMatrix(Y, 8) = "" Then
                        nValorDif = 0
                    Else
                        nValorDif = CDbl(.TextMatrix(Y, 8))
                    End If
                    nValorTributo = CDbl(.TextMatrix(Y, 7)) + nValorDif
                End If
                
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nCodTrib & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
                cn.Execute Sql, rdExecDirect
            Next
        End With
        
        'RETORNA ULTIMO DOCUMENTO
        Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nLastCod = !maximo + 1
           .Close
        End With
        
        ' Grava NumDocumento
        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO,TIPODOC,emissor,valorguia) VALUES("
        Sql = Sql & nLastCod & ",'" & Format(dDataBase, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",2,'" & NomeDeLogin & " (PARCELAMENTO)" & "'," & Virg2Ponto(RemovePonto(CStr(.ListItems(x).SubItems(13)))) & ")"
        cn.Execute Sql, rdExecDirect
        
        nPlano = Val(txtPlano.Tag)
        'Grava PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,PLANO) VALUES(" & Val(txtCod.Text) & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nLastCod & "," & nPlano & ")"
        cn.Execute Sql, rdExecDirect
        
        lvDestino.ListItems(x).SubItems(14) = nLastCod
        
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
                          sCidadeEntrega = !descCidade
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
                sNomeResp = !RazaoSocial
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
                    sCidadeEntrega = !descCidade
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
                        nNumEntrega = SubNull(!NUMIMOVEL)
                        sBairroEntrega = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        sComplEntrega = SubNull(!Complemento)
                        sCepEntrega = SubNull(!Cep)
                        sCidadeEntrega = IIf(IsNull(!descCidade), SubNull(!DESCCIDADE1), SubNull(!descCidade))
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
        sTipoImposto = "REPARCEL."
        sTipoEnd = "R"
        Sql = "select * from cidadao where codcidadao=" & Val(txtCod.Text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If SubNull(RdoAux2!etiqueta) = "N" And SubNull(RdoAux2!etiqueta2) = "S" Then
                sTipoEnd = "C"
            End If
            RdoAux2.Close
        End If
        
        
        If sTipoEnd = "R" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
            Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
            Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
            Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
        Else
            Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
            Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
            Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
            Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        On Error Resume Next
        With RdoAux2
            If .RowCount > 0 Then
                 sCodReduz = !CodCidadao
                 sNomeResp = !nomecidadao
                 If Val(SubNull(!FCodLogradouro)) > 0 Then
                     Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                     Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                     Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                     Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                     With RdoS
                         If .RowCount > 0 Then
                            sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                         Else
                            sEndImovel = ""
                         End If
                        .Close
                     End With
                 Else
                    sEndImovel = SubNull(!FNomeLogradouro)
                 End If
                 nNumImovel = Val(SubNull(RdoAux2!fNUMIMOVEL))
                  
                 Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                 Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                 If RdoS.RowCount > 0 Then
                     sCidadeEntrega = RdoS!descCidade
                 Else
                      sCidadeEntrega = ""
                 End If
                 If Not IsNull(!CodBairro) Then
                     Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                     If .RowCount > 0 Then
                         sBairroEntrega = RdoS!DescBairro
                     Else
                         sBairroEntrega = ""
                     End If
                 Else
                     sBairroEntrega = ""
                 End If
                 sUFEntrega = SubNull(!fsiglauf)
                 sCepEntrega = SubNull(!FCEP)
            Else
                sEndImovel = ""
                sBairroEntrega = ""
                sCidadeEntrega = ""
                sUFEntrega = ""
                sCepEntrega = ""
            End If
                If SubNull(!CPF) <> "" Then
                   sCPF = !CPF
                ElseIf SubNull(!Cnpj) <> "" Then
                   sCPF = !Cnpj
                ElseIf SubNull(!rg) <> "" Then
                   sCPF = !rg
                Else
                    sCPF = ""
                End If
           .Close
        End With
    
End Select

'GRAVA TEMPORARIO
With lvDestino
    For x = 1 To .ListItems.Count - 2
        sDataVencto = .ListItems(x).SubItems(2)
        nAno = Val(Right$(sDataVencto, 4))
        nLanc = 20
'        nSeq = nSeq
        nParc = .ListItems(x).Text
        nCompl = 0
        nNumDoc = .ListItems(x).SubItems(14)
        nNumGuia = nNumDoc
        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)

'        If bAnistia And chkAnistia.value = vbChecked Then
'            sValorParc = CStr(CDbl(.ListItems(X).SubItems(12)) + 7) '7,00=Valor da Taxa
'            nValorTaxa = 0
'        Else
            sValorParc = .ListItems(x).SubItems(13)
'            nValorTaxa = 0
'        End If
        nValorGuia = sValorParc
        
        sValor = nValorGuia
        dDataVencto = CDate(sDataVencto)
        sDescImposto = "PARCELAMENTO"
        NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumDoc, CLng(txtCod.Text))
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
    
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        sBarra = StrBarra2
    
    '*******************************************
        Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numproc,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & Val(txtCod.Text) & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 1, nParc) & "," & Val(cmbQtde.Text) & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & "Parcelamento: " & Left$(txtNumProc.Text, 25) & "','" & Left$(txtNumProc.Text, 25) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
        cn.Execute Sql, rdExecDirect
        
    Next
End With

For x = 1 To grdTributo.Rows - 1
    Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & ","
    Sql = Sql & Val(grdTributo.TextMatrix(x, 0)) & ",'" & grdTributo.TextMatrix(x, 1) & "'," & Virg2Ponto(RemovePonto(grdTributo.TextMatrix(x, 6))) & ")"
    cn.Execute Sql, rdExecDirect
Next

nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
nAnoproc = Right$(txtNumProc.Text, 4)

frmConfissaoDivida.txtNumProc.Text = CStr(nNumproc) & "/" & CStr(nAnoproc)
frmConfissaoDivida.txtNumProc.Locked = True
frmConfissaoDivida.lblSid.Caption = nSid
frmConfissaoDivida.lblDI.Caption = lblDI.Caption
frmConfissaoDivida.CarregaProcesso

LockPanel False
Limpa
txtCod_KeyPress vbKeyReturn
cmbQtde.Clear
txtNumProc.Text = ""
lvDestino.ListItems.Clear



End Sub

Private Sub FillDetalhe()
Dim nValorTributo1 As Double, nValorTributoN As Double, nValorJuros1 As Double, nValorJurosN As Double, nDif1 As Double, nDifN As Double
Dim nValorMulta1 As Double, nValorMultaN As Double, nValorCorrecao1 As Double, nValorCorrecaoN As Double, nValorJurosApl1 As Double, nValorTotal As Double
Dim nValorJurosAplN As Double, nValorHonorario1 As Double, nValorHonorarioN As Double, nPercTributo As Double, nSomaTotal1 As Double, nSomaTotalN As Double, Y, nPerc1 As Double, nPercN As Double, nPerc As Double
Dim nTotalPrincipal As Double, nTotalMulta As Double, nTotalJuros As Double, nTotalCorrecao As Double, nTotalJrApl As Double, nTotalHon As Double, nValorP1 As Double, nValorPN As Double, nQtdeParcela As Integer
    
nTotalPrincipal = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count - 1).SubItems(3))
nTotalJuros = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count - 1).SubItems(4))
nTotalMulta = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count - 1).SubItems(5))
nTotalCorrecao = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count - 1).SubItems(6))
nTotalJrApl = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count - 1).SubItems(11))
nTotalHon = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count - 1).SubItems(12))
nValorTotal = nTotalPrincipal + nTotalJuros + nTotalMulta + nTotalCorrecao + nTotalJrApl + nTotalHon
'verificar a proporção  entre as parcelas
nValorP1 = CDbl(lvDestino.ListItems(1).SubItems(3))
nValorPN = CDbl(lvDestino.ListItems(2).SubItems(3))
nQtdeParcela = lvDestino.ListItems.Count - 3

grdDestino.Rows = 1
With grdTributo
    For x = 1 To .Rows - 1
        If Val(.TextMatrix(x, 2)) > 0 Then
            nValorP1 = CDbl(lvDestino.ListItems(1).SubItems(3))
            nValorPN = CDbl(lvDestino.ListItems(2).SubItems(3))
            nPerc = CDbl(.TextMatrix(x, 7))
            grdDestino.AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & FormatNumber(nTotalPrincipal) & Chr(9) & .TextMatrix(x, 7) & Chr(9) & FormatNumber(nTotalPrincipal * CDbl(.TextMatrix(x, 7)) / 100, 2) & Chr(9) & _
            FormatNumber(nValorP1 * nPerc / 100, 2) & Chr(9) & "" & Chr(9) & FormatNumber((nValorPN * nPerc / 100), 2)
        End If
    Next
    
    nValorP1 = CDbl(lvDestino.ListItems(1).SubItems(13))
    nValorPN = CDbl(lvDestino.ListItems(2).SubItems(13))
    
    nPerc = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count).SubItems(4))
    grdDestino.AddItem 113 & Chr(9) & "JUROS" & Chr(9) & FormatNumber(nTotalJuros, 2) & Chr(9) & "100,00" & Chr(9) & FormatNumber(nTotalJuros, 2) & Chr(9) & _
                        FormatNumber(nValorP1 * nPerc / 100, 2) & Chr(9) & "" & Chr(9) & FormatNumber((nValorPN * nPerc / 100), 2)
    
    nPerc = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count).SubItems(5))
    grdDestino.AddItem 112 & Chr(9) & "MULTA" & Chr(9) & FormatNumber(nTotalMulta, 2) & Chr(9) & "100,00" & Chr(9) & FormatNumber(nTotalMulta, 2) & Chr(9) & _
                        FormatNumber(nValorP1 * nPerc / 100, 2) & Chr(9) & "" & Chr(9) & FormatNumber((nValorPN * nPerc / 100), 2)
    
    nPerc = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count).SubItems(6))
    grdDestino.AddItem 26 & Chr(9) & "CORREÇÃO" & Chr(9) & FormatNumber(nTotalCorrecao, 2) & Chr(9) & "100,00" & Chr(9) & FormatNumber(nTotalCorrecao, 2) & Chr(9) & _
                        FormatNumber(nValorP1 * nPerc / 100, 2) & Chr(9) & "" & Chr(9) & FormatNumber((nValorPN * nPerc / 100), 2)
                        
    nPerc = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count).SubItems(11))
    grdDestino.AddItem 585 & Chr(9) & "JUROS APL." & Chr(9) & FormatNumber(nTotalJrApl, 2) & Chr(9) & "100,00" & Chr(9) & FormatNumber(nTotalJrApl, 2) & Chr(9) & _
                        FormatNumber(nValorP1 * nPerc / 100, 2) & Chr(9) & "" & Chr(9) & FormatNumber((nValorPN * nPerc / 100), 2)
                        
    nPerc = CDbl(lvDestino.ListItems(lvDestino.ListItems.Count).SubItems(12))
    grdDestino.AddItem 90 & Chr(9) & "HONORÁRIOS" & Chr(9) & FormatNumber(nTotalHon, 2) & Chr(9) & "100,00" & Chr(9) & FormatNumber(nTotalHon, 2) & Chr(9) & _
                        FormatNumber(nValorP1 * nPerc / 100, 2) & Chr(9) & "" & Chr(9) & FormatNumber((nValorPN * nPerc / 100), 2)
                        
    nSomaTotal1 = 0: nSomaTotalN = 0
    For Y = 1 To grdDestino.Rows - 1
        nSomaTotal1 = nSomaTotal1 + CDbl(grdDestino.TextMatrix(Y, 5))
        nSomaTotalN = nSomaTotalN + CDbl(grdDestino.TextMatrix(Y, 7))
    Next
    nDif1 = (nSomaTotal1 - CDbl(lvDestino.ListItems(1).SubItems(13))) * (-1)
    nDifN = (nSomaTotalN - CDbl(lvDestino.ListItems(2).SubItems(13))) * (-1)
    grdDestino.TextMatrix(1, 6) = FormatNumber(nDif1, 2)
    grdDestino.TextMatrix(1, 8) = FormatNumber(nDifN, 2)
    grdDestino.AddItem "" & Chr(9) & "TOTAL==>" & Chr(9) & "" & Chr(9) & "" & Chr(9) & "" & Chr(9) & FormatNumber(nSomaTotal1, 2) & Chr(9) & FormatNumber(nDif1, 2) & Chr(9) & FormatNumber(nSomaTotalN, 2) & Chr(9) & FormatNumber(nDifN, 2)
End With
   
End Sub

Private Sub AtualizaRefis()
Dim nPerc As Double, Sql As String, RdoAux As rdoResultset, x As Integer, nValor_Juros As Double, nValor_Multa As Double
Dim nValor_Desconto As Double, nValor_Correcao As Double, nValor_Principal As Double, nValor_Total As Double


Sql = "select desconto from plano where codigo=" & Val(txtPlano.Tag)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nPerc = RdoAux!desconto
RdoAux.Close

With lvOrigem
    For x = 1 To .ListItems.Count
        nValor_Principal = CDbl(.ListItems(x).ListSubItems(7))
        nValor_Juros = CDbl(.ListItems(x).ListSubItems(8))
        nValor_Multa = CDbl(.ListItems(x).ListSubItems(9))
        nValor_Correcao = CDbl(.ListItems(x).ListSubItems(10))
        nValor_Desconto = nValor_Juros - (nValor_Juros * nPerc / 100)
       .ListItems(x).ListSubItems(8) = FormatNumber(nValor_Desconto, 2)
        nValor_Desconto = nValor_Multa - (nValor_Multa * nPerc / 100)
       .ListItems(x).ListSubItems(9) = FormatNumber(nValor_Desconto, 2)
        nValor_Juros = CDbl(.ListItems(x).ListSubItems(8))
        nValor_Multa = CDbl(.ListItems(x).ListSubItems(9))
        nValor_Total = nValor_Principal + nValor_Multa + nValor_Juros + nValor_Correcao
       .ListItems(x).ListSubItems(14) = FormatNumber(nValor_Total, 2)
    Next
End With

End Sub
