VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frm2ViaLaser 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de guias"
   ClientHeight    =   6255
   ClientLeft      =   1755
   ClientTop       =   3495
   ClientWidth     =   11850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6255
   ScaleWidth      =   11850
   Begin VB.ComboBox cmbEnd 
      Height          =   315
      ItemData        =   "frm2ViaLaser.frx":0000
      Left            =   2970
      List            =   "frm2ViaLaser.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   84
      Top             =   5805
      Visible         =   0   'False
      Width           =   1410
   End
   Begin prjChameleon.chameleonButton cmdZoomP 
      Height          =   240
      Left            =   10620
      TabIndex        =   35
      ToolTipText     =   "Expandir tela"
      Top             =   45
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
      MICON           =   "frm2ViaLaser.frx":0026
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdZoomM 
      Height          =   240
      Left            =   10950
      TabIndex        =   36
      ToolTipText     =   "Reduzir a tela"
      Top             =   45
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
      MICON           =   "frm2ViaLaser.frx":0042
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frZoom 
      BackColor       =   &H00EEEEEE&
      Height          =   2940
      Left            =   7710
      TabIndex        =   33
      Top             =   225
      Width           =   4020
      Begin MSComctlLib.ListView lvTrib 
         Height          =   2760
         Left            =   45
         TabIndex        =   34
         Top             =   135
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   4868
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Desc.Tributo"
            Object.Width           =   4047
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Qtde"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descrição Completa"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin VB.Frame fr7 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Composição do lancamento"
      Height          =   4065
      Left            =   7650
      TabIndex        =   21
      Top             =   45
      Width           =   4125
      Begin VB.TextBox txtAbate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3135
         TabIndex        =   31
         Text            =   "0,00"
         Top             =   3420
         Width           =   915
      End
      Begin VB.CheckBox chkTxExp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEEE&
         Caption         =   "Emitir com Taxa de Expediente..:"
         Enabled         =   0   'False
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
         Left            =   705
         TabIndex        =   29
         Top             =   3780
         Width           =   3255
      End
      Begin prjChameleon.chameleonButton cmdQtde 
         Height          =   330
         Left            =   90
         TabIndex        =   22
         ToolTipText     =   "Altera a Qtde do Tributo"
         Top             =   3195
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   582
         BTYPE           =   3
         TX              =   "Qtde"
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
         MICON           =   "frm2ViaLaser.frx":005E
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
         Caption         =   "Abatimento em NF:"
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
         Height          =   210
         Left            =   1395
         TabIndex        =   30
         Top             =   3450
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Total:"
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
         Height          =   210
         Left            =   1950
         TabIndex        =   25
         Top             =   3150
         Width           =   1005
      End
      Begin VB.Label lblTotal 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   3015
         TabIndex        =   24
         Top             =   3150
         Width           =   975
      End
      Begin VB.Label lblTotalUnica 
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   2730
         TabIndex        =   23
         Top             =   3930
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   45
      TabIndex        =   77
      Top             =   2700
      Width           =   7575
      Begin VB.TextBox txtAno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6315
         MaxLength       =   4
         TabIndex        =   7
         Top             =   930
         Width           =   915
      End
      Begin VB.CheckBox chkUnica 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Parcela Única"
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
         Left            =   5415
         TabIndex        =   5
         Top             =   630
         Width           =   2070
      End
      Begin VB.TextBox txtNumParc 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   1485
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   2
         Top             =   585
         Width           =   735
      End
      Begin VB.ComboBox cmbLanc 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   225
         Width           =   6015
      End
      Begin esMaskEdit.esMaskedEdit mskDataInicio 
         Height          =   285
         Left            =   3645
         TabIndex        =   6
         Top             =   930
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   15658734
         MouseIcon       =   "frm2ViaLaser.frx":007A
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
      Begin prjChameleon.chameleonButton cmdAddData 
         Height          =   270
         Left            =   4785
         TabIndex        =   4
         ToolTipText     =   "Editar Datas de Vencimento"
         Top             =   600
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
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
         MICON           =   "frm2ViaLaser.frx":0096
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskDataVencimento 
         Height          =   285
         Left            =   3645
         TabIndex        =   3
         Top             =   585
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   16777215
         MouseIcon       =   "frm2ViaLaser.frx":00B2
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
         BackStyle       =   0  'Transparent
         Caption         =   "Ano do Exercício:"
         Height          =   225
         Index           =   17
         Left            =   4935
         TabIndex        =   83
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label lblDataVencto 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5475
         TabIndex        =   82
         Top             =   1065
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Efetuar o cálculo proporcional a partir da data.....:"
         Height          =   225
         Index           =   16
         Left            =   60
         TabIndex        =   81
         Top             =   975
         Width           =   3555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data 1º Vencto.:"
         Height          =   225
         Index           =   15
         Left            =   2385
         TabIndex        =   80
         Top             =   630
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Parcelas.....:"
         Height          =   225
         Index           =   14
         Left            =   60
         TabIndex        =   79
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Lançamento.:"
         Height          =   225
         Index           =   13
         Left            =   45
         TabIndex        =   78
         Top             =   300
         Width           =   1470
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   45
      TabIndex        =   62
      Top             =   1620
      Width           =   7575
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...:"
         Height          =   225
         Index           =   12
         Left            =   4275
         TabIndex        =   76
         Top             =   750
         Width           =   390
      End
      Begin VB.Label lblCepEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6015
         TabIndex        =   75
         Top             =   735
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep......:"
         Height          =   225
         Index           =   8
         Left            =   5325
         TabIndex        =   74
         Top             =   750
         Width           =   585
      End
      Begin VB.Label lblCidadeEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1275
         TabIndex        =   73
         Top             =   720
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade...........:"
         Height          =   225
         Index           =   7
         Left            =   45
         TabIndex        =   72
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label lblBairroEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4995
         TabIndex        =   71
         Top             =   435
         Width           =   2460
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...:"
         Height          =   225
         Index           =   5
         Left            =   4290
         TabIndex        =   70
         Top             =   450
         Width           =   690
      End
      Begin VB.Label lblComplentrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1275
         TabIndex        =   69
         Top             =   435
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.:"
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   68
         Top             =   435
         Width           =   1155
      End
      Begin VB.Label lblNumEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6750
         TabIndex        =   67
         Top             =   150
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Index           =   3
         Left            =   6330
         TabIndex        =   66
         Top             =   150
         Width           =   405
      End
      Begin VB.Label lblRuaEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1275
         TabIndex        =   65
         Top             =   150
         Width           =   4860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço.......:"
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   64
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label lbluf 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   4725
         TabIndex        =   63
         Top             =   735
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   45
      TabIndex        =   49
      Top             =   540
      Width           =   7575
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4050
         TabIndex        =   61
         Top             =   765
         Width           =   1845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro..:"
         Height          =   225
         Index           =   11
         Left            =   3450
         TabIndex        =   60
         Top             =   765
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.............:"
         Height          =   225
         Index           =   10
         Left            =   45
         TabIndex        =   59
         Top             =   750
         Width           =   1740
      End
      Begin VB.Label lblCompl 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1710
         TabIndex        =   58
         Top             =   765
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep.:"
         Height          =   225
         Index           =   9
         Left            =   6060
         TabIndex        =   57
         Top             =   780
         Width           =   420
      End
      Begin VB.Label lblCep 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6480
         TabIndex        =   56
         Top             =   765
         Width           =   990
      End
      Begin VB.Label lblNumImovel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6480
         TabIndex        =   55
         Top             =   480
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Index           =   1
         Left            =   6060
         TabIndex        =   54
         Top             =   495
         Width           =   405
      End
      Begin VB.Label lblRua 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1710
         TabIndex        =   53
         Top             =   480
         Width           =   3690
      End
      Begin VB.Label lblProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1710
         TabIndex        =   52
         Top             =   195
         Width           =   5700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...................:"
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   51
         Top             =   450
         Width           =   1695
      End
      Begin VB.Label lblRS 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário.................:"
         Height          =   225
         Left            =   60
         TabIndex        =   50
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   510
      Left            =   45
      TabIndex        =   43
      Top             =   45
      Width           =   7575
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   165
         Width           =   945
      End
      Begin prjChameleon.chameleonButton cmdCnsImovel 
         Height          =   315
         Left            =   2730
         TabIndex        =   45
         ToolTipText     =   "Consulta Imóvel"
         Top             =   135
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
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
         MICON           =   "frm2ViaLaser.frx":00CE
         PICN            =   "frm2ViaLaser.frx":00EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblNumInsc 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4680
         TabIndex        =   48
         Top             =   180
         Width           =   2790
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido/I.M.:"
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   47
         Top             =   180
         Width           =   1695
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         Caption         =   "Insc.Cadastral:"
         Height          =   225
         Left            =   3540
         TabIndex        =   46
         Top             =   180
         Width           =   1200
      End
   End
   Begin VB.OptionButton optGuia 
      Caption         =   "Boleto"
      Height          =   195
      Index           =   1
      Left            =   6300
      TabIndex        =   42
      Top             =   5760
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.OptionButton optGuia 
      Caption         =   "Normal"
      Height          =   195
      Index           =   0
      Left            =   5355
      TabIndex        =   41
      Top             =   5760
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox cmbAnoTabela 
      Height          =   315
      ItemData        =   "frm2ViaLaser.frx":0244
      Left            =   855
      List            =   "frm2ViaLaser.frx":0263
      Style           =   2  'Dropdown List
      TabIndex        =   40
      Top             =   5805
      Width           =   1005
   End
   Begin VB.CheckBox chk2010 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Tabela de 2011"
      Enabled         =   0   'False
      Height          =   240
      Left            =   6975
      TabIndex        =   38
      Top             =   5850
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CheckBox chkCalculo2010 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Calculo para 2010"
      Enabled         =   0   'False
      Height          =   240
      Left            =   6885
      TabIndex        =   32
      Top             =   5895
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Frame fr6 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Atividade Vig.Sanitária"
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   7650
      TabIndex        =   12
      Top             =   4185
      Width           =   4125
      Begin MSComctlLib.ListView lvVS 
         Height          =   1230
         Left            =   60
         TabIndex        =   15
         Top             =   225
         Width           =   3990
         _ExtentX        =   7038
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CNAE"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "CR"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   9878
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Qtde"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   1235
         EndProperty
      End
   End
   Begin VB.Frame fr5 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Atividade Tx.Licença"
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   3870
      TabIndex        =   11
      Top             =   4185
      Width           =   3810
      Begin MSComctlLib.ListView lvTL 
         Height          =   1230
         Left            =   60
         TabIndex        =   14
         Top             =   225
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   9878
         EndProperty
      End
   End
   Begin VB.Frame fr4 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Atividade ISS"
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   30
      TabIndex        =   10
      Top             =   4185
      Width           =   3810
      Begin MSComctlLib.ListView lvISS 
         Height          =   1230
         Left            =   60
         TabIndex        =   13
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1236
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   10233
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1260
      Left            =   270
      TabIndex        =   8
      Top             =   6510
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   2223
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código     |^Ano     |^Lanc. |^Seq  |^Parc. |^Compl. |^Vencimento      |>Vl.Lançado  |<Num.Documento      "
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   10065
      TabIndex        =   9
      ToolTipText     =   "Sair da Tela"
      Top             =   5805
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frm2ViaLaser.frx":029D
      PICN            =   "frm2ViaLaser.frx":02B9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   345
      Left            =   8685
      TabIndex        =   0
      ToolTipText     =   "Gera as guias informadas"
      Top             =   5805
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Emitir Guia"
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
      MICON           =   "frm2ViaLaser.frx":0327
      PICN            =   "frm2ViaLaser.frx":0343
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp2 
      Height          =   1260
      Left            =   270
      TabIndex        =   28
      Top             =   7920
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   2223
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código     |^Ano     |^Lanc. |^Seq  |^Parc. |^Compl. |^Vencimento      |>Vl.Lançado  |<Num.Documento      "
   End
   Begin VB.Frame pnlObs 
      BackColor       =   &H00000080&
      Caption         =   "Observação da Parcela"
      ForeColor       =   &H00FFFFFF&
      Height          =   2235
      Left            =   2430
      TabIndex        =   18
      Top             =   2160
      Visible         =   0   'False
      Width           =   6645
      Begin VB.TextBox txtObs 
         Height          =   1545
         Left            =   90
         MaxLength       =   450
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   6435
      End
      Begin prjChameleon.chameleonButton cmdSairObs 
         Height          =   300
         Left            =   5580
         TabIndex        =   20
         ToolTipText     =   "Sair da Tela"
         Top             =   1845
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
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
         MICON           =   "frm2ViaLaser.frx":049D
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
   Begin Tributacao.jcFrames pnlData 
      Height          =   4785
      Left            =   4815
      Top             =   630
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   8440
      FillColor       =   14745599
      Style           =   4
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Vencimentos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin MSFlexGridLib.MSFlexGrid grdData 
         Height          =   3795
         Left            =   90
         TabIndex        =   16
         Top             =   480
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   6694
         _Version        =   393216
         RowHeightMin    =   280
         BackColor       =   16777215
         BackColorBkg    =   12648447
         Appearance      =   0
         FormatString    =   "Parcela       |^Data               "
      End
      Begin prjChameleon.chameleonButton cmdSairData 
         Height          =   345
         Left            =   660
         TabIndex        =   17
         ToolTipText     =   "Sair da Tela"
         Top             =   4320
         Width           =   1350
         _ExtentX        =   2381
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
         MICON           =   "frm2ViaLaser.frx":04B9
         PICN            =   "frm2ViaLaser.frx":04D5
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
   Begin VB.Label lblEnd 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço..:"
      Height          =   240
      Left            =   2070
      TabIndex        =   85
      Top             =   5850
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tabela..:"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   39
      Top             =   5850
      Width           =   645
   End
   Begin VB.Label lblArea 
      Caption         =   "Area"
      Height          =   240
      Left            =   7650
      TabIndex        =   37
      Top             =   5895
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label lbltipoend 
      Height          =   225
      Left            =   10410
      TabIndex        =   27
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lbllanc 
      Caption         =   "Label3"
      Height          =   225
      Left            =   9540
      TabIndex        =   26
      Top             =   6690
      Width           =   735
   End
End
Attribute VB_Name = "frm2ViaLaser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuContrib As cPopupMenu
Attribute m_cMenuContrib.VB_VarHelpID = -1
Private Type TRIBUTO
    nCodTributo As Integer
    nValorTributo As Double
End Type
Private Type TAXALICENCA
    nValorAliq As Double
    nArea As Double
End Type
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim sRet As String, bExec As Boolean
Dim xImovel As clsImovel, sTipoIss As String, nArea As Double
Dim nValorTxExpParc As Double, nValorTxExpUnica As Double, aValorAliquotaTxL() As TAXALICENCA
Dim aISS() As Integer, aTL() As Long, aVS() As String, nCodigoImovel
Dim nQtdeProf As Integer, bGerado As Boolean, bRocada As Boolean, sObsIss As String, nAreaIss As Double, nCPF As Single

Private Sub chkTxExp_Click()
FillTotal
End Sub

Private Sub chkUnica_Click()
If chkUnica.value = 1 And Val(txtNumParc.Text) < 2 Then
   MsgBox "Parcelas únicas somente acima de 2 parcelas.", vbExclamation, "Atenção"
   chkUnica.value = 0
   Exit Sub
End If
End Sub

Private Sub cmbEnd_Click()

If Not bExec Then Exit Sub
If Val(txtCod.Text) = 0 Then Exit Sub

If cmbEnd.ListIndex = 0 Then
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
bExec = True
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblProp.Caption = !nomecidadao
         If Not IsNull(!Cnpj) Then
             If Val(!Cnpj) > 0 Then
                 lblNum.Caption = "CNPJ:"
                 lblNumInsc.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
             End If
         Else
             If Not IsNull(!CPF) Then
                 If Val(!CPF) > 0 Then
                     lblNum.Caption = "CPF:"
                     lblNumInsc.Caption = Format(RetornaNumero(!CPF), "00#\.###\.###-##")
                 End If
             End If
         End If
         If Val(SubNull(!FCodLogradouro)) > 0 Then
             Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
             Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
             Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
             Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
             Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoS
                 If .RowCount > 0 Then
                    sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                 Else
                    sEnd = ""
                 End If
                .Close
             End With
         Else
            sEnd = SubNull(!FNomeLogradouro)
         End If
               
         Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & Val(SubNull(!fCodCidade))
         Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
         If RdoS.RowCount > 0 Then
             sCidade = RdoS!descCidade
         Else
              sCidade = ""
         End If
         If Not IsNull(RdoAux!fCodBairro) Then
             Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
             Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
             If .RowCount > 0 Then
                 sBairro = SubNull(RdoS!DescBairro)
             Else
                 sBairro = ""
             End If
         Else
             sBairro = ""
         End If
         sUF = SubNull(!fsiglauf)
         sFone = SubNull(!fTELEFONE)
         sCep = SubNull(!FCEP)
    
         lblRua.Caption = sEnd
         lblNumImovel.Caption = !fNUMIMOVEL
         lblCompl.Caption = SubNull(!fcomplemento)
         lblBairro.Caption = sBairro
         lblCidadeEntrega.Caption = sCidade
         lblCEP.Caption = sCep
         lblUF.Caption = sUF
         
    
        lblRuaEntrega.Caption = lblRua.Caption
        lblNumEntrega.Caption = lblNumImovel.Caption
        lblComplentrega.Caption = lblCompl.Caption
        lblBairroEntrega.Caption = lblBairro.Caption
        lblCepEntrega.Caption = lblCEP.Caption
         cmbLanc.SetFocus
    Else
        MsgBox "Código não cadastrado.", vbCritical, "Atenção"
    End If
   .Close
End With


End Sub

Private Sub cmbLanc_Click()
Dim itmX As ListItem, nCodLanc As Integer
Dim bIPTU As Boolean, nCodTrib As Integer
Dim z As Long

If cmbLanc.ListIndex = -1 Then Exit Sub
If Not bExec Then
   bExec = True
   Exit Sub
End If

If cmbLanc.ItemData(cmbLanc.ListIndex) = 50 Or cmbLanc.ItemData(cmbLanc.ListIndex) = 65 Then
    txtAbate.Locked = False
    txtAbate.BackColor = Branco
Else
    txtAbate.Locked = True
    txtAbate.BackColor = Kde
End If
txtAbate.Text = "0,00"

z = SendMessage(lvTrib.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvISS.HWND, LVM_DELETEALLITEMS, 0, 0)

bExec = False
If Val(txtCod.Text) = 0 Then
    MsgBox "Selecione o contribuinte.", vbExclamation, "Atenção"
    cmbLanc.ListIndex = -1
    bExec = True
    Exit Sub
End If

If lblProp.Caption = "" Then
    MsgBox "Contribuinte não Cadastrado.", vbExclamation, "Atenção"
    cmbLanc.ListIndex = -1
    bExec = True
    Exit Sub
End If

If Val(txtCod.Text) < 40000 Then bIPTU = True Else bIPTU = False
nCodLanc = cmbLanc.ItemData(cmbLanc.ListIndex)

If bIPTU And (nCodLanc = 2 Or nCodLanc = 3 Or nCodLanc = 5) Then
    MsgBox "Um imóvel não pode ter lançamentos de ISS.", vbExclamation, "Atenção"
    cmbLanc.ListIndex = -1
    bExec = True
    Exit Sub
End If

If Not bIPTU And nCodLanc = 1 Then
    MsgBox "Uma empresa/Contribuinte não pode ter lançamentos de IPTU.", vbExclamation, "Atenção"
    cmbLanc.ListIndex = -1
    bExec = True
    Exit Sub
End If

If bIPTU And nCodLanc = 6 Then
    MsgBox "Um imóvel não pode ter Taxa de Licença de Funcionamento.", vbExclamation, "Atenção"
    cmbLanc.ListIndex = -1
    bExec = True
    Exit Sub
End If

If bIPTU And nCodLanc = 13 Then
    MsgBox "Um imóvel não pode ter Taxa de Vigilância Sanitária.", vbExclamation, "Atenção"
    cmbLanc.ListIndex = -1
    bExec = True
    Exit Sub
End If

bExec = True
If cmbLanc.ListIndex > -1 Then
    Sql = "SELECT DESCREDUZ FROM LANCAMENTO WHERE CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        lblLanc.Caption = !descreduz
       .Close
    End With
    
    Sql = "SELECT CODTRIBUTO,ABREVTRIBUTO,DESCTRIBUTO FROM vwTRIBUTOLANCAMENTO WHERE CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex) & " AND CODTRIBUTO<>3 and codtributo<>13 ORDER BY ABREVTRIBUTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            If !CodTributo <> 124 Then
                Set itmX = lvTrib.ListItems.Add(, "C" & Format(!CodTributo, "000"), !ABREVTRIBUTO)
                itmX.SubItems(1) = 0
                itmX.SubItems(2) = "0,0000"
                itmX.SubItems(3) = !desctributo
            End If
           .MoveNext
        Loop
       .Close
    End With
End If

'ISS
nCodTrib = 0
If nCodLanc = 2 Or nCodLanc = 14 Then
    nCodTrib = 11
ElseIf nCodLanc = 5 Then
    nCodTrib = 13
ElseIf nCodLanc = 3 Then
    nCodTrib = 12
End If
If nCodTrib > 0 Then
    Sql = "SELECT MOBILIARIOATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE FROM MOBILIARIOATIVIDADEISS INNER JOIN "
    Sql = Sql & "ATIVIDADEISS ON MOBILIARIOATIVIDADEISS.CODATIVIDADE = ATIVIDADEISS.CODATIVIDADE "
    Sql = Sql & "Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & Val(txtCod.Text) & " AND CODTRIBUTO=" & nCodTrib
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    With RdoAux
        Do Until .EOF
           Set itmX = lvISS.ListItems.Add(, "IS" & Format(!codatividade, "000") & .AbsolutePosition, Format(!codatividade, "000"))
           itmX.SubItems(1) = !descatividade
          .MoveNext
        Loop
    End With
End If
Select Case nCodLanc
    Case 2, 3, 5, 6, 11, 13, 14, 33
        mskDataInicio.Locked = False
        mskDataInicio.BackColor = Branco
        chkUnica.Enabled = True
        txtNumParc.Locked = False
        txtNumParc.BackColor = Branco
    Case Else
        mskDataInicio.Locked = True
        mskDataInicio.BackColor = Kde
        chkUnica.Enabled = False
        txtNumParc.Text = 1
        txtNumParc.BackColor = Kde
        txtNumParc.Locked = False
End Select
FillTotal
txtNumParc.SetFocus
End Sub

Private Sub cmdAddData_Click()
Dim Sql As String, RdoAux As rdoResultset

If Val(txtNumParc.Text) = 0 Then
    MsgBox "Digite o nº de parcelas.", vbExclamation, "Atenção"
    txtNumParc.SetFocus
    Exit Sub
End If

If Val(txtNumParc.Text) > 48 Then
    MsgBox "Máximo 48 parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataVencimento.Text) Then
    MsgBox "Digite o 1º vencimento.", vbExclamation, "Atenção"
    mskDataVencimento.SetFocus
    Exit Sub
End If


'df = ValidaFeriado(CDate(mskDataVencimento.Text))
'If df = 1 Then
'    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
'    Else
'        Exit Sub
'    End If
'ElseIf df = 2 Then
'    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
'    Else
'        Exit Sub
'    End If
'ElseIf df = 3 Then
'    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
'    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
'    Sql = Sql & " Where DIA = " & Day(CDate(mskDataVencimento.Text))
'    Sql = Sql & " AND MES=" & Month(CDate(mskDataVencimento.Text)) & " AND ANO=" & Year(CDate(mskDataVencimento.Text))
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If .RowCount > 0 Then
'            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
'                mskDataVencimento.Text = RetornaDiaUtil(CDate(mskDataVencimento.Text))
'            Else
'                Exit Sub
'            End If
'          .Close
'        End If
'    End With
'End If
'
''If grdData.Rows = 1 Then
''    grdData.Rows = Val(txtNumParc.Text) + 1
''    AutoFillDate2
''    For x = 1 To Val(txtNumParc.Text)
''        grdData.TextMatrix(x, 0) = "Parcela " & Format(x, "00")
''    Next
''End If

pnlData.Visible = True
pnlData.ZOrder 0
cmdSair.Enabled = False
cmdBaixa.Enabled = False


End Sub

Private Sub cmdBaixa_Click()
Dim x As Integer, Achou As Boolean, z As Variant, Sql As String, RdoAux As rdoResultset
Dim NumProc As Long, AnoProc As Integer, nDV As Integer

If bLocal Then
    Exit Sub
End If

bRocada = False
If lblProp.Caption = "" Then
    MsgBox "Selecione o Contribuinte.", vbExclamation, "Atenção"
    Exit Sub
End If

If nCPF = 0 Then
    MsgBox "Não é possível emitir guia para contribuinte que não possui um CPF/CNPJ válido.", vbExclamation, "Atenção"
    Exit Sub
End If


If cmbLanc.ListIndex = -1 Then
    MsgBox "Selecione o Lançamento.", vbExclamation, "Atenção"
    Exit Sub
End If

If lvTrib.ListItems.Count = 0 Then
    MsgBox "Lançamento não possue tributos relacionados.", vbExclamation, "Atenção"
    Exit Sub
End If

If lblRuaEntrega.Caption = "" Then
    MsgBox "Endereço de Entrega incompleto.", vbExclamation, "Atenção"
    Exit Sub
End If

Achou = False
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
       Achou = True
       Exit For
    End If
Next

If Not Achou Then
    MsgBox "Selecione os tributos.", vbExclamation, "Atenção"
    Exit Sub
End If

For x = 1 To lvTrib.ListItems.Count
    If Val(Right(lvTrib.ListItems(x).Key, 3)) = 170 Then
       bRocada = True
       Exit For
    End If
Next

'If chkTxExp.Value = vbUnchecked Then
'    If MsgBox("Deseja emitir esta guia SEM TAXA DE EXPEDIENTE ?", vbQuestion + vbYesNo + vbDefaultButton2, "TAXA DE EXPEDIENTE") = vbNo Then Exit Sub
'End If

Achou = False
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True And CDbl(lvTrib.ListItems(x).SubItems(2)) = 0 And lvTrib.ListItems(x) <> "ISS VARIÁVEL" Then
       Achou = True
       Exit For
    End If
Next

If Achou Then
    If MsgBox("Existem tributos sem valor definido, deseja continuar e imprimir estes tributos com o Valor 0 (Zero) ?.", vbQuestion + vbYesNo, "Atenção") = vbNo Then
       Exit Sub
    End If
End If

If Val(txtNumParc.Text) = 0 Then
    MsgBox "Digite a qtde de parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataVencimento.Text) Then
    MsgBox "Data do 1º Vencimento inválido.", vbCritical, "Atenção"
    cmdAddData.SetFocus
    Exit Sub
End If

For x = 1 To grdData.Rows - 1
    If Not IsDate(grdData.TextMatrix(x, 1)) Then
        MsgBox "Data de vencimento da parcela nº " & x & " inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
Next

For x = 1 To grdData.Rows - 2
    If CDate(grdData.TextMatrix(x, 1)) >= CDate(grdData.TextMatrix(x + 1, 1)) Then
        MsgBox "Data de vencimento da parcela nº " & x & " não pode ser maior ou igual ao vencimento da parcela nº " & x + 1 & ".", vbExclamation, "Atenção"
        Exit Sub
    End If
Next


If Val(txtAno.Text) < 1995 And Val(txtAno.Text) > Year(Now) Then
    MsgBox "Ano do exercício fora do intervalo.", vbExclamation, "Atenção"
    Exit Sub
End If

If (lblNum.Caption = "") Then
    MsgBox "CPF/CNPJ obrigatório para emissão de guias.", vbCritical, "Erro"
    Exit Sub
End If

If (lblRua.Caption = "") Then
    MsgBox "Endereço obrigatório para emissão de guias.", vbCritical, "Erro"
    Exit Sub
End If

If (lblBairro.Caption = "") Then
    MsgBox "Bairro obrigatório para emissão de guias.", vbCritical, "Erro"
    Exit Sub
End If

If (lblCEP.Caption = "") Then
    MsgBox "Cep obrigatório para emissão de guias.", vbCritical, "Erro"
    Exit Sub
End If


NumeroProcesso = ""
If bRocada Then
    z = InputBox("Digite o número do processo.", "Entre com os dados")
    If z = "" Then
        MsgBox "Digite o número do processo.", vbCritical, "Atenção"
        Exit Sub
    Else
        NumeroProcesso = z
        If InStr(1, NumeroProcesso, "/", vbBinaryCompare) = 0 Or Len(NumeroProcesso) < 2 Then
            GoTo Erro
        Else
            If Right(NumeroProcesso, 1) = "/" Then
                GoTo Erro
            End If
        End If
        nDV = Val(Right(Left$(NumeroProcesso, InStr(1, NumeroProcesso, "/", vbBinaryCompare) - 1), 1))
        If InStr(1, NumeroProcesso, "/", vbBinaryCompare) < 3 Then
            GoTo Erro
        End If
        NumProc = Left$(NumeroProcesso, InStr(1, NumeroProcesso, "/", vbBinaryCompare) - 2)
        AnoProc = Right$(NumeroProcesso, 4)
        If nDV <> RetornaDVProcesso(NumProc) Then
            GoTo Erro
        End If
        Sql = "SELECT * FROM PROCESSOGTI WHERE ANO=" & AnoProc & " AND NUMERO=" & NumProc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount = 0 Then
            GoTo Erro
        End If
        txtObs.Text = "Número do Processo: " & NumeroProcesso
        txtObs.SelStart = Len(txtObs.Text)
        
        Sql = "SELECT * FROM ETIQUETAROCADA WHERE CODREDUZIDO=" & Val(txtCod.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                Sql = "INSERT ETIQUETAROCADA(CODREDUZIDO,DATA) VALUES(" & Val(txtCod.Text) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
            Else
                Sql = "UPDATE ETIQUETAROCADA SET DATA='" & Format(Now, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & Val(txtCod.Text)
            End If
            cn.Execute Sql, rdExecDirect
           .Close
        End With
    End If
End If

If cmbLanc.ItemData(cmbLanc.ListIndex) = 1 Then
    If MsgBox("Emitir guia para o exercício de " & txtAno.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
        Exit Sub
    End If
End If

pnlObs.Visible = True
pnlObs.ZOrder 0
txtObs.SetFocus
cmdBaixa.Enabled = False
cmdSair.Enabled = False

Exit Sub
Erro:
MsgBox "Número de Processo inválido ou não cadastrado.", vbCritical, "Atenção"

End Sub


Private Sub cmdCnsImovel_Click()
lIndex = m_cMenuContrib.ShowPopupMenu(cmdCnsImovel.Left, cmdCnsImovel.Top, cmdCnsImovel.Left, cmdCnsImovel.Top, Me.ScaleWidth - cmdCnsImovel.Left - cmdCnsImovel.Width, cmdCnsImovel.Top + cmdCnsImovel.Height, False)
End Sub



Private Sub cmdQtde_Click()
Dim z As Variant, nCodTributo As Integer

If lvTrib.ListItems.Count = 0 Then
    MsgBox "Não existem tributos.", vbExclamation, "Atenção"
    Exit Sub
End If

If lvTrib.SelectedItem.Checked = False Then
    MsgBox "Selecione o item antes de alterar sua qtde.", vbExclamation, "Atenção"
    Exit Sub
End If

nCodTributo = Val(Right$(lvTrib.ListItems(lvTrib.SelectedItem.Index).Key, 3))
Select Case nCodTributo
    Case 11, 12, 13, 14, 25, 170
        MsgBox "A Quantidade deste tributo não pode ser alterada.", vbExclamation, "Atenção"
    Case Else
        z = InputBox("Digite a Qtde.", "Quantidade", lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1))
        If z = "" Then z = lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1)
        z = Ponto2Virg(CStr(z))
        If Not IsNumeric(z) Then
            MsgBox "Qtde inválida.", vbExclamation, "Atenção"
        Else
            If z = 0 Then z = 1
            lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1) = z
            FillTotal
        End If
End Select


End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSairData_Click()
pnlData.Visible = False
cmdSair.Enabled = True
cmdBaixa.Enabled = True
If grdData.Rows > 1 Then
    mskDataVencimento.Text = grdData.TextMatrix(1, 1)
End If
'chkUnica.SetFocus
End Sub

Private Sub cmdSairObs_Click()
Dim x As Integer, bFind As Boolean, sTributo As String
Dim Sql As String, RdoAux As rdoResultset, nSeq As Integer
sObsIss = ""
pnlObs.Visible = False
cmdBaixa.Enabled = True
cmdSair.Enabled = True

sTributo = ""
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked Then
        If (cmbLanc.ItemData(cmbLanc.ListIndex) = 50 Or cmbLanc.ItemData(cmbLanc.ListIndex) = 65) Then
            sTributo = lvTrib.ListItems(x).Text
            nAreaIss = lvTrib.ListItems(x).SubItems(1)
        End If
    End If
Next

If sTributo <> "" Then
Inicio:
    nCodigoImovel = InputBox("Digite o código do imóvel a ser lançado o ISS Constução Civil", "Campo obrigatório")
    If Trim(nCodigoImovel) = "" Then GoTo Inicio
    Sql = "select codreduzido from cadimob where codreduzido=" & Val(nCodigoImovel)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        bFind = True
    End If
    RdoAux.Close
    If Not bFind Then GoTo Inicio
    sObsIss = "Emitido guia de ISS construção civil do tipo -> " & sTributo & " com área de: " & Format(nAreaIss, "#0.00") & " m² para o imóvel:" & nCodigoImovel & " Lançado por " & NomeDeLogin & " no código cidadão:" & txtCod.Text
    
'    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodigoImovel
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If IsNull(!maximo) Then
'            nSeq = 1
'        Else
'            nSeq = !maximo + 1
'        End If
'        .Close
'    End With
'
'    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
'    Sql = Sql & nCodigoImovel & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
'    cn.Execute Sql, rdExecDirect

    
End If

If Val(txtNumParc.Text) = 1 Then
'    If NomeDeLogin = "SCHWARTZ" Then
        EmiteBoletoRegistrado
'    Else
'        EmiteBoleto
'    End If
Else
    EmiteBoleto
End If

If bRocada Then
    GravaMulta
    frmReport.ShowReport "MULTAINF", frmMdi.HWND, Me.HWND
End If

If bGerado Then Limpa

End Sub

Private Sub cmdZoomM_Click()
lvTrib.ColumnHeaders(4).Width = 0
frZoom.Left = 7695
frZoom.Width = 3705
lvTrib.Width = 3615
End Sub

Private Sub cmdZoomP_Click()
lvTrib.ColumnHeaders(4).Width = 6800
frZoom.Left = 30
frZoom.Width = Me.Width - 200
lvTrib.Width = frZoom.Width - 100
End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
     txtCod.Text = Val(Left$(CodImovel, 7))
     CodImovel = 0
     txtCod_LostFocus
Else
    If Val(CodEmpresa) > 0 Then
         txtCod.Text = Val(Left$(CodEmpresa, 7))
         CodEmpresa = 0
         txtCod_LostFocus
    Else
        If Val(CodCidadao) > 0 Then
             Unload frmCnsCidadao
             If cGetInputState() <> 0 Then DoEvents
             txtCod.Text = Val(CodCidadao)
             CodCidadao = 0
             txtCod_LostFocus
        End If
    End If
End If
End Sub

Private Sub Form_Load()
cmbEnd.ListIndex = 0
If NomeDeLogin <> "SCHWARTZ" Then
    optGuia(1).Enabled = False
End If

Ocupado
MontaMenu
Set xImovel = New clsImovel
Centraliza Me
Liberado
sRet = RetEventUserForm(Me.Name)
'If RetornaUsuarioFiscal Then
'    Sql = "SELECT CODLANCAMENTO,DESCFULL FROM LANCAMENTO WHERE CODLANCAMENTO not in (25,29) ORDER BY DESCFULL"
'Else
    Sql = "SELECT CODLANCAMENTO,DESCFULL FROM LANCAMENTO WHERE codlancamento not in (25,29,5,3) ORDER BY DESCFULL"
'End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbLanc.AddItem !DESCFULL
        cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With
bExec = True
txtAbate.Locked = True
txtAbate.BackColor = Kde
cmbAnoTabela.Text = Year(Now)
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "RITA" Or _
    NomeDeLogin = "LEANDRO" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "SOLANGE" Or NomeDeLogin = "RODRIGOG" Or NomeDeLogin = "AAFMARTINS" Or IsAtendente Then
    cmbAnoTabela.Enabled = True
Else
    cmbAnoTabela.Enabled = False
End If
lvTrib.ColumnHeaders(4).Width = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set xImovel = Nothing
Set m_cMenuContrib = Nothing
End Sub

Private Sub grdData_DblClick()
Dim z As Variant, nRow As Integer
nRow = grdData.Row

Inicio:
z = InputBox("Digite o novo vencimento para a parcela " & nRow, "Alteração de Vencimento", grdData.TextMatrix(grdData.Row, 1))
If IsDate(z) Then
    If nRow > 1 Then
        If CDate(z) < CDate(grdData.TextMatrix(nRow - 1, 1)) Then
            MsgBox "Data da parcela " & nRow & " não pode ser inferior ou igual a parcela anterior.", vbExclamation, "Atenção"
            GoTo Inicio
        End If
    End If
    grdData.TextMatrix(nRow, 1) = z
End If


End Sub


Private Sub lvISS_ItemCheck(ByVal Item As MSComctlLib.ListItem)

Dim x As Integer
For x = 1 To lvTrib.ListItems.Count
    If Val(Right$(lvTrib.ListItems(x).Key, 3)) = 11 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 12 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 13 Then
       lvTrib.ListItems(x).Selected = True
       lvTrib.SelectedItem.Checked = True
       Exit For
    End If
Next

lvTrib_ItemCheck lvTrib.SelectedItem

End Sub

Private Sub lvTL_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim x As Integer
For x = 1 To lvTrib.ListItems.Count
    If Val(Right$(lvTrib.ListItems(x).Key, 3)) = 14 Then
       lvTrib.ListItems(x).Selected = True
       lvTrib.SelectedItem.Checked = True
       Exit For
    End If
Next

lvTrib_ItemCheck lvTrib.SelectedItem

End Sub

Private Sub lvTrib_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim nCodTributo As Integer, RdoVig As rdoResultset, nValorAliquotaVS As Double
Dim nValorTxLic As Double, x As Integer, nAno As Integer, RdoAux2 As rdoResultset
Dim nQtdeVS As Integer

If lvTrib.ListItems.Count = 0 Then
    MsgBox "Não há impostos selecionados.", vbExclamation, "Atenção"
    CarregaLista
    Exit Sub
End If

If Val(txtNumParc.Text) = 0 Then
    MsgBox "Selecione o nº de parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

'If chk2009.Value = vbChecked Then
'    nAno = 2009
'If chk2010.Value = vbChecked Then
'    nAno = 2011
'Else
'    nAno = Year(Now)
'End If
nAno = Val(cmbAnoTabela.Text)

nCodTributo = Val(Right$(lvTrib.ListItems(Item.Index).Key, 3))

'If nCodTributo = 502 And Item.Checked = True Then
'    MsgBox "O ISS Variável Taxa Diversa deve ser utilizado apenas quando se deseja lançar o Boleto com Valor já definido pois este valor sera dividido entre o número de parcelas. Caso voce deseja emitir os boletos com o valor do ISS Variável zerado, desmarque esta opção, e marque a opção de ISS Variável normal.", vbInformation, "Atenção"
'End If

If Item.Checked = False And (nCodTributo = 11 Or nCodTributo = 12 Or nCodTributo = 13) Then
    For x = 1 To lvISS.ListItems.Count
        lvISS.ListItems(x).Checked = False
    Next
End If
If Item.Checked = False And nCodTributo = 14 Then
    For x = 1 To lvTL.ListItems.Count
        lvTL.ListItems(x).Checked = False
    Next
End If
If Item.Checked = False And nCodTributo = 25 Then
    For x = 1 To lvVS.ListItems.Count
        lvVS.ListItems(x).Checked = False
    Next
End If


ReDim aISS(0): ReDim aTL(0): ReDim aVS(0)
If Item.Checked = True And (nCodTributo = 11 Or nCodTributo = 12 Or nCodTributo = 13 Or nCodTributo = 14 Or nCodTributo = 25) Then
    'Carrega Matriz
    For x = 1 To lvISS.ListItems.Count
        If lvISS.ListItems(x).Checked = True Then
            ReDim Preserve aISS(UBound(aISS) + 1)
            aISS(UBound(aISS)) = Val(lvISS.ListItems(x).Text)
        End If
    Next
    For x = 1 To lvTL.ListItems.Count
        If lvTL.ListItems(x).Checked = True Then
            ReDim Preserve aTL(UBound(aTL) + 1)
            aTL(UBound(aTL)) = Val(lvTL.ListItems(x).Text)
        End If
    Next
    For x = 1 To lvVS.ListItems.Count
        If lvVS.ListItems(x).Checked = True Then
            ReDim Preserve aVS(UBound(aVS) + 1)
            'aVS(UBound(aVS)) = lvVS.ListItems(x).text & lvVS.ListItems(x).SubItems(1)
            aVS(UBound(aVS)) = lvVS.ListItems(x).Text & "-" & lvVS.ListItems(x).SubItems(1) 'joga na matriz o cnae e o criterio
        End If
    Next
End If


'TAXA DE LICENÇA
If Item.Checked = True And nCodTributo = 14 Then
    ReDim aValorAliquotaTxL(0)
    Sql = "SELECT MOBILIARIO.CODATIVIDADE,QTDEPROF ,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3,AREATL,CODIGOALIQ FROM MOBILIARIO INNER JOIN "
    Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE Where CODIGOMOB =" & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        For x = 1 To UBound(aTL)
            If !codatividade = aTL(x) Then
                Select Case !CODIGOALIQ
                    Case 1
                        aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                    Case 2
                        aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                    Case 3
                        aValorAliquotaTxL(0).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                End Select
                If Not IsNull(!areatl) Then
                   aValorAliquotaTxL(0).nArea = IIf(!areatl = 0, 1, !areatl)
                Else
                   aValorAliquotaTxL(0).nArea = 0
                End If
                nQtdeProf = Val(SubNull(!QTDEPROF))
                If nQtdeProf = 0 Then nQtdeProf = 1
            End If
        Next
       .Close
    End With
    
    Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,MOBILIARIOATIVIDADETL.AREA,ATIVIDADE.DESCATIVIDADE,MOBILIARIOATIVIDADETL.CODIGOALIQ,"
    Sql = Sql & "ATIVIDADE.VALORALIQ1, ATIVIDADE.VALORALIQ2,ATIVIDADE.VALORALIQ3, MOBILIARIO.AREATL,MOBILIARIO.QTDEPROF "
    Sql = Sql & "FROM ATIVIDADE INNER JOIN MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
    Sql = Sql & "Inner Join MOBILIARIO ON MOBILIARIOATIVIDADETL.CODIGOMOB = MOBILIARIO.CODIGOMOB "
    Sql = Sql & "where MOBILIARIOATIVIDADETL.CODIGOMOB=" & Val(txtCod.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            For x = 1 To UBound(aTL)
                If !codatividade = aTL(x) Then
                    ReDim Preserve aValorAliquotaTxL(UBound(aValorAliquotaTxL) + 1)
                    Select Case !CODIGOALIQ
                        Case 1
                            aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ1, 2)
                        Case 2
                            aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ2, 2)
                        Case 3
                            aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nValorAliq = FormatNumber(!VALORALIQ3, 2)
                    End Select
                    aValorAliquotaTxL(UBound(aValorAliquotaTxL)).nArea = IIf(IsNull(!Area), 0, !Area)
                End If
            Next
           .MoveNext
        Loop
       .Close
    End With
End If


With lvTrib
    lvTrib.ListItems(Item.Index).Selected = True
    Select Case nCodTributo
        Case 3 'TAXA EXPEDIÇÃO DE DOCUMENTO
            If .ListItems(Item.Key).Checked = True Then
                Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE CODTRIBUTO=3 AND ANO=" & nAno
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    nValorTxExpParc = !valoraliq
                    nValorTxExpUnica = !valoraliq
                    lvTrib.ListItems(Item.Index).SubItems(1) = Val(txtNumParc.Text)
                    lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(nValorTxExpParc, 2)
                    .Close
                End With
                lvTrib.ListItems(Item.Index).ForeColor = vbRed
            Else
                lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
            End If
        Case 11 'ISS FIXO
            If .ListItems(Item.Key).Checked = True Then
                If UBound(aISS) = 0 Then
                   lvTrib.ListItems(Item.Index).Checked = False
                   lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                   lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
                   Exit Sub
                End If
                If TemISSnoAno(2) Then
                    lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                    lvTrib.ListItems(Item.Index).SubItems(2) = CalculaISS("F", nCodTributo)
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                Else
                    lvTrib.ListItems(Item.Index).SubItems(2) = CalculaISS("F", nCodTributo)
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                End If
            Else
                lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
            End If
        Case 12 'ISS ESTIMADO
            If .ListItems(Item.Key).Checked = True Then
                If UBound(aISS) = 0 Then
                   MsgBox "Selecione as atividades de ISS.", vbExclamation, "Atenção"
                   lvTrib.ListItems(Item.Index).Checked = False
                   lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                   lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
                   Exit Sub
                End If
                If TemISSnoAno(3) Then
                    lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                    lvTrib.ListItems(Item.Index).SubItems(2) = CalculaISS("E", nCodTributo)
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                Else
                    lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                    lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(CalculaISS("E", nCodTributo), 2)
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                End If
            Else
                lvTrib.ListItems(Item.Key).Checked = False
                lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
            End If
        Case 13 'ISS VARIAVEL
            If .ListItems(Item.Key).Checked = True Then
                If TemISSnoAno(5) Then
                    lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                    lvTrib.ListItems(Item.Index).SubItems(2) = CalculaISS("V", nCodTributo)
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                Else
                    lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                    lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
                    lvTrib.ListItems(Item.Index).ForeColor = vbRed
                End If
            Else
                lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
            End If
        Case 14 'TAXA DE LICENÇA
            If .ListItems(Item.Key).Checked = True Then
                If UBound(aTL) = 0 Then
                   lvTrib.ListItems(Item.Index).Checked = False
                   lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                   lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
                   FillTotal
                   Exit Sub
                End If
                'CALCULA TAXA DE LICENÇA
                nValorTxLic = 0
                If sTipoIss = "F" Then
                   For x = 0 To UBound(aValorAliquotaTxL)
                       If chkCalculo2010.value = vbUnchecked Then
                          nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAno) * nQtdeProf)
                       Else
                          nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(2010) * nQtdeProf)
                       End If
                   Next
                Else
                   For x = 0 To UBound(aValorAliquotaTxL)
                       If aValorAliquotaTxL(x).nValorAliq <= 14 Then
                          If Not IsDate(mskDataInicio.Text) Then
                             If chkCalculo2010.value = vbUnchecked Then
                                nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(x).nArea * RetornaUFIR(Year(Now)) * nQtdeProf)
                             Else
                                nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(x).nArea * RetornaUFIR(Year(Now)) * nQtdeProf)
                             End If
                          Else
                             nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * aValorAliquotaTxL(x).nArea * nQtdeProf) * RetornaUFIR(Year(Now))
                          End If
                       Else
                          If Not IsDate(mskDataInicio.Text) Then
'                             If chkCalculo2010.Value = vbUnchecked Then
'                                 nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(Year(Now)) * nQtdeProf)
'                             Else
                                 nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(nAno) * nQtdeProf)
'                             End If
                          Else
                             nValorTxLic = nValorTxLic + (aValorAliquotaTxL(x).nValorAliq * RetornaUFIR(Year(mskDataInicio.Text)) * nQtdeProf)
                          End If
                       End If
                   Next
                End If
                nValorTxLic = nValorTxLic * MesesProporcional / 12
                lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(nValorTxLic, 2)
                lvTrib.ListItems(Item.Index).ForeColor = vbRed
            Else
                lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
                lvTrib.ListItems(Item.Index).ForeColor = vbBlack
            End If
        Case 25 'VIGILÂNCIA SANITÁRIA
            If .ListItems(Item.Key).Checked = True Then
                If UBound(aVS) = 0 Then
                   lvTrib.ListItems(Item.Index).Checked = False
                   lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                   lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
                   FillTotal
                   Exit Sub
                End If
                
                'Sql = "SELECT * FROM MOBILIARIOATIVIDADEVS2 WHERE CODMOBILIARIO=" & Val(txtCod.Text)
                Sql = "SELECT mobiliariovs.codigo, mobiliariovs.cnae, mobiliariovs.criterio, mobiliariovs.qtde, cnaecriterio.valor "
                Sql = Sql & "FROM  mobiliariovs INNER JOIN cnaecriterio ON mobiliariovs.cnae = cnaecriterio.cnae "
                Sql = Sql & "Where mobiliariovs.Codigo = " & Val(txtCod.Text)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount = 0 Then
                        MsgBox "Esta empresa não possue Taxa de Vigilância Sanitária.", vbExclamation, "Atenção"
                        lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                        lvTrib.ListItems(Item.Key).Checked = False
                        FillTotal
                        Exit Sub
                    Else
                        nValorAliquotaVS = 0
                        For x = 1 To lvVS.ListItems.Count
                            If lvVS.ListItems(x).Checked Then
                                nQtdeVS = Val(lvVS.ListItems(x).SubItems(3))
                                If nQtdeVS = 0 Then nQtdeVS = 1
                                nValorAliquotaVS = nValorAliquotaVS + (CDbl(lvVS.ListItems(x).SubItems(4)) * nQtdeVS)
                            End If
                        Next
                    End If
                End With
                nValorAliquotaVS = nValorAliquotaVS * MesesProporcional / 12 'proporcional
                lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(nValorAliquotaVS, 2)
                lvTrib.ListItems(Item.Index).ForeColor = vbRed
            Else
                lvTrib.ListItems(Item.Index).SubItems(1) = "0"
                lvTrib.ListItems(Item.Index).SubItems(2) = "0,0000"
            End If
        Case 170 'ROÇADA
            If lblArea.Caption = "" Then lblArea.Caption = "0"
            lvTrib.ListItems(Item.Index).SubItems(1) = lblArea.Caption
            Sql = "select valoraliq from tributoaliquota where ano=" & Year(Now) & " and codtributo=170"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            lvTrib.ListItems(Item.Index).SubItems(2) = RdoAux2!valoraliq
            RdoAux2.Close
        Case Else
            If .ListItems(Item.Key).Checked = True Then
                Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & nAno & " AND CODTRIBUTO = " & nCodTributo
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        If lvTrib.ListItems(Item.Index).SubItems(1) = "0" Then
                            lvTrib.ListItems(Item.Index).SubItems(1) = "1"
                        End If
                        lvTrib.ListItems(Item.Index).SubItems(2) = FormatNumber(!valoraliq, 4)
                        
                        lvTrib.ListItems(Item.Index).ForeColor = vbRed
                    Else
                        MsgBox "Não existe tarifa para este tributo." & vbCrLf & "Consulte a Tabela de Preços Públicos.", vbExclamation, "Atenção"
                        lvTrib.ListItems(Item.Index).Checked = False
                        lvTrib.ListItems(Item.Index).ForeColor = vbBlack
                    End If
                End With
            Else
                .ListItems(Item.Index).SubItems(1) = "0"
                .ListItems(Item.Index).SubItems(2) = "0,0000"
                lvTrib.ListItems(Item.Index).ForeColor = vbBlack
            End If
   End Select
End With

FillTotal

End Sub

Private Sub lvVS_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim x As Integer
For x = 1 To lvTrib.ListItems.Count
    If Val(Right$(lvTrib.ListItems(x).Key, 3)) = 25 Then
       lvTrib.ListItems(x).Selected = True
       lvTrib.SelectedItem.Checked = True
       Exit For
    End If
Next

lvTrib_ItemCheck lvTrib.SelectedItem

End Sub

Private Sub m_cMenuContrib_Click(ItemNumber As Long)

Select Case m_cMenuContrib.ItemKey(ItemNumber)
    Case "mnuMob"
        sFormMob = "EI2"
        frmCnsMob.show
        frmCnsMob.ZOrder 0
    Case "mnuImob"
        sForm = "EI"
        frmCnsImovel.show
        frmCnsImovel.ZOrder 0
    Case "mnuOutros"
        Set frm = frmCnsCidadao
        frm.sForm = "frm2ViaLaser"
        frm.show
        frm.ZOrder 0
End Select

End Sub

Private Sub mskDataInicio_Click()
mskDataInicio.SelStart = 0
mskDataInicio.SelLength = 10
End Sub

Private Sub mskDataInicio_GotFocus()
mskDataInicio.SelStart = 0
mskDataInicio.SelLength = 10
'mskDataInicio.SetFocus
mskDataInicio.Refresh
'If mskDataInicio.Locked = True Then
'   lvISS.SetFocus
'End If
End Sub


Private Sub mskDataVencimento_GotFocus()
mskDataVencimento.SetFocus
mskDataVencimento.SelStart = 1
mskDataVencimento.SelLength = Len(mskDataVencimento.Text)
End Sub


Private Sub mskDataVencimento_LostFocus()
grdData.Rows = 1


If Not IsDate(mskDataVencimento.Text) Then Exit Sub

df = ValidaFeriado(CDate(mskDataVencimento.Text))
If df = 1 Then
    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
'    Else
 '       Exit Sub
    End If
ElseIf df = 2 Then
    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
    Else
        If grdData.Rows = 1 Then
            grdData.Rows = Val(txtNumParc.Text) + 1
            AutoFillDate2
            For x = 1 To Val(txtNumParc.Text)
                grdData.TextMatrix(x, 0) = "Parcela " & Format(x, "00")
            Next
        End If
        Exit Sub
    End If
ElseIf df = 3 Then
    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
    Sql = Sql & " Where DIA = " & Day(CDate(mskDataVencimento.Text))
    Sql = Sql & " AND MES=" & Month(CDate(mskDataVencimento.Text)) & " AND ANO=" & Year(CDate(mskDataVencimento.Text))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                mskDataVencimento.Text = RetornaDiaUtil(CDate(mskDataVencimento.Text))
            Else
                'Exit Sub
            End If
          .Close
        End If
    End With
End If

If grdData.Rows = 1 Then
    grdData.Rows = Val(txtNumParc.Text) + 1
    AutoFillDate2
    For x = 1 To Val(txtNumParc.Text)
        grdData.TextMatrix(x, 0) = "Parcela " & Format(x, "00")
    Next
End If

End Sub

Private Sub txtAbate_Change()
    FillTotal
End Sub

Private Sub txtAbate_KeyPress(KeyAscii As Integer)
Tweak txtAbate, KeyAscii, DecimalPositive
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_Change()
Dim z As Long
z = SendMessage(lvTrib.HWND, LVM_DELETEALLITEMS, 0, 0)
End Sub

Private Sub txtCod_GotFocus()

txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    If cmbLanc.Enabled = True Then
        cmbLanc.SetFocus
    End If
    Exit Sub
End If

Tweak txtCod, KeyAscii, IntegerPositive

End Sub

Public Sub txtCod_LostFocus()
CarregaDados
If NomeDeLogin = "REGINALDO" Then
    cmbLanc.Text = "TAXAS DE CEMITÉRIO"
    cmbLanc.Enabled = False
End If
End Sub

Private Sub CarregaDados()
Dim nCodImovel As Long, sTipoEnd As String

If Val(txtCod.Text) = 0 Then Exit Sub
nCodImovel = Val(txtCod.Text)

If nCodImovel >= 500000 Then
    lblEnd.Visible = True
    cmbEnd.Visible = True
Else
    lblEnd.Visible = False
    cmbEnd.Visible = False
End If

txtAno.Text = Year(Now)
Limpa
Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & txtCod.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If !Inativo = 1 Then
           MsgBox "Este imóvel encontra-se inativo.", vbExclamation, "Atenção"
           Exit Sub
        End If
'        lblNum.Caption = "Insc.Cadastral"
        lblRS.Caption = "Proprietário"
        CarregaImovel nCodImovel
        Sql = "select codreduzido,cpf,cnpj from vwfullimovel where codreduzido=" & Val(txtCod.Text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If SubNull(RdoAux2!CPF) <> "" Then
'            nCPF = RdoAux2!CPF
'        Else
'            If SubNull(RdoAux2!Cnpj) <> "" Then
'                nCPF = RdoAux2!Cnpj
'            Else
'                nCPF = 0
'            End If
  '      End If
        
        nCPF = 0
        If Val(SubNull(RdoAux2!Cnpj)) > 0 Then
            lblNum.Caption = "CNPJ:"
            lblNumInsc.Caption = Format(RdoAux2!Cnpj, "0#\.###\.###/####-##")
            nCPF = RdoAux2!Cnpj
        Else
            If Not IsNull(RdoAux2!CPF) Then
                If Val(RdoAux2!CPF) > 0 Then
                    lblNum.Caption = "CPF:"
                    lblNumInsc.Caption = Format(RetornaNumero(RdoAux2!CPF), "00#\.###\.###-##")
                    nCPF = RdoAux2!CPF
                End If
            End If
        End If

        
        
    Else
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,CNPJ,CPF,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR,CODCIDADE,DESCCIDADE "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & txtCod.Text
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
              'suspenção
               Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.Text
               Sql = Sql & " ORDER BY DATAEVENTO DESC"
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                       If !CODTIPOEVENTO = 2 Then
                           MsgBox "Esta empresa esta SUSPENSA", vbExclamation, "Atenção"
'                           Exit Sub
                       End If
                   End If
                  .Close
               End With
                nCPF = 0
                If Val(SubNull(!Cnpj)) > 0 Then
                    lblNum.Caption = "CNPJ:"
                    lblNumInsc.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
                    nCPF = !Cnpj
                Else
                    If Not IsNull(!CPF) Then
                        If Val(!CPF) > 0 Then
                            lblNum.Caption = "CPF:"
                            lblNumInsc.Caption = Format(RetornaNumero(!CPF), "00#\.###\.###-##")
                            nCPF = !CPF
                        End If
                    End If
                End If
               
               lblRS.Caption = "Raz.Social"
               lblProp.Caption = !razaosocial
               lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
               If Trim(lblRua.Caption) = "" Then
                    lblRua.Caption = SubNull(!NomeLogr)
               End If
               lblNumImovel.Caption = Val(SubNull(!Numero))
               lblCEP.Caption = IIf(IsNull(!Cep), "", Left$(!Cep, 5) & "-" & Right$(!Cep, 3))
               lblCompl.Caption = SubNull(!Complemento)
               Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=" & !CodCidade & " AND CODBAIRRO=" & !CodBairro
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                        lblBairro.Caption = !DescBairro
                   Else
                        lblBairro.Caption = ""
                   End If
                  .Close
               End With
               Sql = "SELECT NOMELOGRADOURO,NUMIMOVEL,COMPLEMENTO,UF,CIDADE.DESCCIDADE AS DESCCIDADE1,"
               Sql = Sql & "BAIRRO.DESCBAIRRO AS DESCBAIRRO1,CEP,MOBILIARIOENDENTREGA.DESCBAIRRO,"
               Sql = Sql & "MOBILIARIOENDENTREGA.DESCCIDADE FROM CIDADE INNER JOIN BAIRRO ON "
               Sql = Sql & "CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE RIGHT OUTER Join "
               Sql = Sql & "MOBILIARIOENDENTREGA ON BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND "
               Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.Text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                    If .RowCount > 0 Then
                        lblTipoEnd.Caption = "(Endereço de Entrega Específico)"
                        lblRuaEntrega.Caption = SubNull(!NomeLogradouro)
                        lblNumEntrega.Caption = SubNull(!NUMIMOVEL)
                        lblComplentrega.Caption = SubNull(!Complemento)
                        lblBairroEntrega.Caption = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        lblCidadeEntrega.Caption = IIf(IsNull(!descCidade), SubNull(!DESCCIDADE1), SubNull(!descCidade))
                        lblCepEntrega.Caption = SubNull(!Cep)
                        lblUF.Caption = SubNull(!UF)
                    Else
                        lblTipoEnd.Caption = "(Endereço da Empresa)"
                        lblRuaEntrega.Caption = lblRua.Caption
                        lblNumEntrega.Caption = lblNumImovel.Caption
                        lblComplentrega.Caption = lblCompl.Caption
                        lblBairroEntrega.Caption = lblBairro.Caption
                        lblCidadeEntrega.Caption = SubNull(RdoAux!descCidade)
                        lblCepEntrega.Caption = lblCEP.Caption
                        lblUF.Caption = "SP"
                    End If
                   .Close
               End With
               CarregaLista
               cmbLanc.SetFocus
            Else
               Sql = "SELECT CODCIDADAO,CODBAIRRO,CODBAIRRO2 FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               If Val(SubNull(RdoAux!CodBairro)) > 0 Then
                  sTipoEnd = "R"
               Else
                  If Val(SubNull(RdoAux!CodBairro2)) > 0 Then
                     sTipoEnd = "C"
                  Else
                     sTipoEnd = "R"
                  End If
               End If
               RdoAux.Close
               bExec = False
               If sTipoEnd = "R" Then
                    cmbEnd.ListIndex = 0
                    Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                    Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                    Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                    Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
               Else
                    cmbEnd.ListIndex = 1
                    Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
                    Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
                    Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                    Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
               End If
               bExec = True
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                   If .RowCount > 0 Then
                       lblProp.Caption = !nomecidadao
                        nCPF = 0
                        If SubNull(!Cnpj) <> "" Then
                            If Val(!Cnpj) > 0 Then
                                lblNum.Caption = "CNPJ:"
                                lblNumInsc.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
                                nCPF = RetornaNumero(!Cnpj)
                            End If
                        Else
                            If SubNull(!CPF) <> "" Then
                                If Val(!CPF) > 0 Then
                                    lblNum.Caption = "CPF:"
                                    lblNumInsc.Caption = Format(RetornaNumero(!CPF), "00#\.###\.###-##")
                                    nCPF = RetornaNumero(!CPF)
                                End If
                            End If
                        End If
                        If Val(SubNull(!FCodLogradouro)) > 0 Then
                            Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                            Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                            Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                            Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoS
                                If .RowCount > 0 Then
                                   sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                Else
                                   sEnd = ""
                                End If
                               .Close
                            End With
                        Else
                           sEnd = SubNull(!FNomeLogradouro)
                        End If
                              
                        Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & Val(SubNull(!fCodCidade))
                        Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                        If RdoS.RowCount > 0 Then
                            sCidade = RdoS!descCidade
                        Else
                             sCidade = ""
                        End If
                        If Not IsNull(RdoAux!fCodBairro) Then
                            Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                            Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                            If .RowCount > 0 Then
                                sBairro = SubNull(RdoS!DescBairro)
                            Else
                                sBairro = ""
                            End If
                        Else
                            sBairro = ""
                        End If
                        sUF = SubNull(!fsiglauf)
                        sFone = SubNull(!fTELEFONE)
                        sCep = SubNull(!FCEP)
                   
                        lblRua.Caption = sEnd
                        lblNumImovel.Caption = Val(SubNull(!fNUMIMOVEL))
                        lblCompl.Caption = SubNull(!fcomplemento)
                        lblBairro.Caption = sBairro
                        lblCidadeEntrega.Caption = sCidade
                        lblCEP.Caption = sCep
                        lblUF.Caption = sUF
                        
                   
                       lblRuaEntrega.Caption = lblRua.Caption
                       lblNumEntrega.Caption = lblNumImovel.Caption
                       lblComplentrega.Caption = lblCompl.Caption
                       lblBairroEntrega.Caption = lblBairro.Caption
'                       lblCidadeEntrega.Caption = SubNull(!desccidade)
                       lblCepEntrega.Caption = lblCEP.Caption
                       'lbluf.Caption = SubNull(!SiglaUF)
                        If cmbLanc.Enabled = True Then
                            cmbLanc.SetFocus
                        End If
                   Else
                       MsgBox "Código não cadastrado.", vbCritical, "Atenção"
                   End If
                  .Close
               End With
            End If
           .Close
        End With
    End If
End With

End Sub

Private Sub CarregaImovel(nCodigoImovel As Long)
Dim Sql As String, RdoAux As rdoResultset

Ocupado
With xImovel
    .CarregaImovel nCodigoImovel
    If .CodigoImovel > 0 Then
'          lblNumInsc.Caption = .Inscricao
          lblProp.Caption = .NomePropPrincipal
          lblRua.Caption = Trim$(.AbrevTipoLog) & " " & Trim$(.AbrevTitLog) & " " & .NomeLogradouro
          lblNumImovel.Caption = .Li_Num
          lblCEP.Caption = RetornaCEP(.CodLogr, .Li_Num)
          lblCompl.Caption = .Li_Compl
          lblBairro.Caption = .DescBairro
          lblArea.Caption = .Dt_AreaTerreno
          Select Case .Ee_TipoEnd
                Case 0
                    lblTipoEnd.Caption = "(Endereço do Imóvel)"
                    lblRuaEntrega.Caption = lblRua.Caption
                    lblNumEntrega.Caption = lblNumImovel.Caption
                    lblComplentrega.Caption = lblCompl.Caption
                    lblBairroEntrega.Caption = lblBairro.Caption
                    lblCidadeEntrega.Caption = "JABOTICABAL"
                    lblCepEntrega.Caption = lblCEP.Caption
                    lblUF.Caption = lblUF.Caption
                Case 1
                    lblTipoEnd.Caption = "(Endereço do Proprietário)"
                    CarregaEndCidadao .CodPropPrincipal
                Case 2
                    lblTipoEnd.Caption = "(Endereço de Entrega Específico)"
                    lblRuaEntrega.Caption = .Ee_NomeLog
                    lblNumEntrega.Caption = .Ee_NumImovel
                    lblComplentrega.Caption = .Ee_Complemento
                    Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & .Ee_Uf & "' AND CODCIDADE=" & .Ee_Cidade & " AND CODBAIRRO=" & .Ee_Bairro
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            lblBairroEntrega.Caption = !DescBairro
                        End If
                       .Close
                    End With
                    lblCidadeEntrega.Caption = .Ee_Cidade
                    lblCepEntrega.Caption = .Ee_Cep
                    lblUF.Caption = .Ee_Uf
          End Select
    End If
End With

fim:
Liberado

End Sub

Private Sub Limpa()

z = SendMessage(lvTrib.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvISS.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvTL.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvVS.HWND, LVM_DELETEALLITEMS, 0, 0)
lblArea.Caption = ""
lblNum.Caption = ""
lblProp.Caption = ""
lblRua.Caption = ""
lblNumImovel.Caption = ""
lblCompl.Caption = ""
lblBairro.Caption = ""
lblCEP.Caption = ""
lblRuaEntrega.Caption = ""
lblNumEntrega.Caption = ""
lblComplentrega.Caption = ""
lblBairroEntrega.Caption = ""
lblCidadeEntrega.Caption = ""
lblCepEntrega.Caption = ""
lblUF.Caption = ""
lblNumInsc.Caption = ""
lblLanc.Caption = ""
lblTipoEnd.Caption = ""
chkUnica.value = 0
pnlObs.Visible = False
txtObs.Text = ""
txtNumParc.Text = ""
grdData.Rows = 1
LimpaMascara mskDataInicio
LimpaMascara mskDataVencimento
bExec = False: cmbLanc.ListIndex = -1: bExec = True
chkTxExp.value = vbUnchecked
End Sub

Private Sub CarregaEndCidadao(nCodigo As Long)

Sql = "SELECT * FROM vwFULLIMOVEL2 WHERE CODCIDADAO=" & nCodigo
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    lblRuaEntrega.Caption = SubNull(!Logradouro)
    lblNumEntrega.Caption = SubNull(!Li_Num)
    lblComplentrega.Caption = SubNull(!Li_Compl)
    lblBairroEntrega.Caption = SubNull(!DescBairro)
    lblCidadeEntrega.Caption = SubNull(!descCidade)
    lblCepEntrega.Caption = RetornaCEP(!CodLogr, !Li_Num)
    lblUF.Caption = SubNull(!SiglaUF)
End With

End Sub

Private Sub GravaCarneTmp()
On Error GoTo Erro

Dim x As Integer, nAnoEmissao As Integer, sDataEmissao As String, nNumDocEspecial As Long, bAbateu As Boolean
Dim RdoAux2 As rdoResultset, nValorExp As Double
Dim sNumInsc As String
Dim nCodReduz As Long
Dim sNomeResp As String
Dim sValorParc As String
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
Dim dDataVencto As Date
Dim nNumdoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim sVencimento As String
Dim nCodLanc As Integer
Dim nSeq As Integer
Dim nComplemento As Integer
Dim nValorTotal As Double, nValorUnica As Double
Dim nValorParc As Double, sValorBoleto As String
Dim nValorParcUnica As Double
Dim aTributos() As TRIBUTO
Dim aTributosU() As TRIBUTO
Dim NumBarra1 As String, sCPF As String
Dim StrBarra1 As String
Dim NumBarra2 As String
Dim NumBarra2a As String
Dim NumBarra2b As String
Dim NumBarra2c As String
Dim NumBarra2d As String
Dim StrBarra2 As String
Dim nLastCod As Long
Dim sDadosLanc As String
Dim sFullTrib As String, bDesconto As Boolean, bISSFixo As Boolean, bTLL As Boolean, nValorISSFixo As Double

If MsgBox("Confirma criação da Guia ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nAnoEmissao = Val(txtAno.Text)
nValorParc = 0
nValorParcUnica = 0

nAno = Val(cmbAnoTabela.Text)

Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED=" & nAno & " AND CODLANCAMENTO=1"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
     If .RowCount > 0 Then
        nValorExp = FormatNumber(!VALORPARCELA, 2)
     Else
        MsgBox "Taxa de Expediente não cadastrada.", vbCritical, "Atenção"
        Exit Sub
     End If
    .Close
End With


nCodReduz = Val(txtCod.Text)
Sql = "SELECT CODREDUZIDO,CPF,CNPJ,RG,ORGAO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        If Not IsNull(!CPF) Then
           sCPF = !CPF
        ElseIf Not IsNull(!Cnpj) Then
           sCPF = !Cnpj
        ElseIf Not IsNull(!rg) Then
           sCPF = !rg
        Else
           sCPF = ""
        End If
    End If
End With

'CARREGA GRID TEMPORARIO
grdTemp.Rows = 1: grdTemp2.Rows = 1
nCodLanc = cmbLanc.ItemData(cmbLanc.ListIndex)

ReDim aTributos(0): ReDim aTributosU(0)

If chkTxExp.value = vbChecked Then
    Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED = " & nAno
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         nValorTxExpParc = FormatNumber(!VALORPARCELA, 2)
        .Close
    End With
Else
    nValorTxExpParc = 0
End If
bAbateu = False
'CALCULA O VALOR PARCELADO
nValorTotal = CDbl(lblTotal.Caption) - nValorTxExpParc
nValorUnica = CDbl(lblTotalUnica.Caption) - nValorTxExpParc
nValorParc = FormatNumber(nValorTotal / CDbl(txtNumParc.Text), 2)
If Not IsNumeric(txtAbate.Text) Then txtAbate.Text = 0
'MONTA TRIBUTOS
sDadosLanc = ""
bDesconto = False
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
       ReDim Preserve aTributos(UBound(aTributos) + 1)
       ReDim Preserve aTributosU(UBound(aTributosU) + 1)
       aTributosU(UBound(aTributosU)).nCodTributo = Val(Right$(lvTrib.ListItems(x).Key, 3))
       aTributos(UBound(aTributos)).nCodTributo = Val(Right$(lvTrib.ListItems(x).Key, 3))
       
       If chkUnica.value = 0 Then
          If Not bAbateu Then
             aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber((CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1))) - CDbl(txtAbate.Text), 2)
'             bAbateu = True
          Else
             aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber(CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1)), 2)
          End If
       Else
          aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber(CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1)), 2)
          If aTributosU(UBound(aTributosU)).nCodTributo = 22 Or aTributosU(UBound(aTributosU)).nCodTributo = 23 Or aTributosU(UBound(aTributosU)).nCodTributo = 24 Or aTributosU(UBound(aTributosU)).nCodTributo = 28 Then
          Else
             bDesconto = True 'ja foi dado 5% de desconto para iss fixo para não dar la na frente de novo.
             aTributosU(UBound(aTributosU)).nValorTributo = aTributosU(UBound(aTributosU)).nValorTributo - aTributosU(UBound(aTributosU)).nValorTributo * 0.05
          End If
       End If
       If Not bAbateu Then
            '((CDbl(lvTrib.ListItems(x).SubItems(2)) *  CDbl(lvTrib.ListItems(x).SubItems(1)))   - CDbl(txtAbate.text))   / Val(txtNumParc.text)
            aTributos(UBound(aTributos)).nValorTributo = FormatNumber(((CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1))) - CDbl(txtAbate.Text)) / Val(txtNumParc.Text), 2)
            bAbateu = True
       Else
            aTributos(UBound(aTributos)).nValorTributo = FormatNumber((CDbl(lvTrib.ListItems(x).SubItems(2) * CDbl(lvTrib.ListItems(x).SubItems(1))) / Val(txtNumParc.Text)), 2)
       End If
       sDadosLanc = sDadosLanc & lvTrib.ListItems(x).Text & vbCrLf
    End If
Next

If chkTxExp.value = vbChecked Then
    sDadosLanc = sDadosLanc & "TX.EXPEDIENTE" & vbCrLf
       ReDim Preserve aTributos(UBound(aTributos) + 1)
       ReDim Preserve aTributosU(UBound(aTributosU) + 1)
       aTributosU(UBound(aTributosU)).nCodTributo = 3
       aTributos(UBound(aTributos)).nCodTributo = 3
       aTributosU(UBound(aTributosU)).nValorTributo = nValorExp
       aTributos(UBound(aTributos)).nValorTributo = nValorExp
End If

nValorUnica = 0
For x = 0 To UBound(aTributos)
    If aTributosU(x).nCodTributo <> 3 Then
        nValorUnica = nValorUnica + aTributosU(x).nValorTributo
    End If
Next

'**************************************************
'verifica se é um caso de ISS Fixo/TLL
'se tiver devem ser divididos
If cmbLanc.ItemData(cmbLanc.ListIndex) = 2 Then
    bISSFixo = False: bTLL = False
    For Y = 1 To UBound(aTributos)
        If aTributos(Y).nCodTributo = 11 Then
            bISSFixo = True
            nValorISSFixo = aTributos(Y).nValorTributo
        End If
        If aTributos(Y).nCodTributo = 14 Then bTLL = True
    Next
End If
'**************************************************

'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nLastCod = RdoAux!maximo
RdoAux.Close
nNumDocEspecial = nLastCod + 1 'usado para exportação do PDF
For nNumParc = 0 To Val(txtNumParc.Text)
    If nNumParc = 0 And chkUnica.value = 0 Then GoTo proximo 'VALIDA PARCELA ÚNICA
    If nNumParc > 0 Then
        sVencimento = grdData.TextMatrix(nNumParc, 1)
    Else
        sVencimento = mskDataVencimento.Text
    End If
    If sVencimento = "" Then sVencimento = mskDataVencimento.Text
    nAno = nAnoEmissao
    nLastCod = nLastCod + 1
    If nNumParc = 0 Then
       If bDesconto Then
          'ja foi dado desconto la em cima
          sValorBoleto = FormatNumber(nValorUnica, 2)
       Else
          sValorBoleto = FormatNumber(nValorUnica - (nValorUnica * 0.05), 2)
       End If
    Else
       sValorBoleto = FormatNumber(nValorTotal / Val(txtNumParc.Text), 2)
    End If
    
    If bISSFixo And bTLL Then
        'Uma linha para iss fixo e outra para taxa de licença
         'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 14 & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nNumParc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & 14 & Chr(9) & nSeq & Chr(9) & nNumParc & Chr(9) & nComplemento & Chr(9) & sVencimento & Chr(9) & nValorISSFixo & Chr(9) & nLastCod
         'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 6 & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nNumParc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & 6 & Chr(9) & nSeq & Chr(9) & nNumParc & Chr(9) & nComplemento & Chr(9) & sVencimento & Chr(9) & CDbl(sValorBoleto) - nValorISSFixo & Chr(9) & nLastCod
        grdTemp2.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & 2 & Chr(9) & nSeq & Chr(9) & nNumParc & Chr(9) & nComplemento & Chr(9) & sVencimento & Chr(9) & CDbl(sValorBoleto) & Chr(9) & nLastCod
    Else
        'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nNumParc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & nNumParc & Chr(9) & nComplemento & Chr(9) & sVencimento & Chr(9) & sValorBoleto & Chr(9) & nLastCod
        grdTemp2.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & nNumParc & Chr(9) & nComplemento & Chr(9) & sVencimento & Chr(9) & sValorBoleto & Chr(9) & nLastCod
    End If

proximo:
Next
'Exit Sub
'DADOS CABEÇALHO
sNumProc = Format(txtCod.Text, "000000") & "/" & sTr(Year(Now))
dDataProc = Format(Now, "dd/mm/yyyy")
sDescImposto = lblLanc.Caption
NumBarra1 = Format(ExtraiNumero(sNumProc), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)

'GERAÇÃO DOS DÉBITOS
With grdTemp
    For x = 1 To .Rows - 1
          'GRAVA DEBITOPARCELA    // (STATUS 3 - NAO PAGO)
'          Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
'          Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & ","
'          Sql = Sql & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
'          Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
'          Sql = Sql & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
          Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
          Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & ","
          Sql = Sql & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
          Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
          Sql = Sql & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
          cn.Execute Sql, rdExecDirect
          If Trim(txtObs.Text) <> "" Then
                Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
                Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If IsNull(!maximo) Then
                        nSeq = 1
                    Else
                        nSeq = !maximo + 1
                    End If
                   .Close
                End With
                sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'                Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & .TextMatrix(x, 0) & ","
'                Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(txtObs.Text)) & "','"
'                Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
                Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & .TextMatrix(x, 0) & ","
                Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(txtObs.Text)) & "',"
                Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "')"
                cn.Execute Sql, rdExecDirect
          End If
          
          sFullTrib = ""
          For Y = 1 To UBound(aTributos)
             '*************
             If bISSFixo And bTLL Then
                If Val(.TextMatrix(x, 2)) = 14 Then 'se for iss fixo ignorar outros tributos que não sejam iss fixo(11)
                    If aTributos(Y).nCodTributo <> 11 Then
                        GoTo ProximoTrib
                    End If
                ElseIf Val(.TextMatrix(x, 2)) = 6 Then 'se for taxa licenca ignorar o tributo iss fixo(11)
                    If aTributos(Y).nCodTributo = 11 Then
                        GoTo ProximoTrib
                    End If
                End If
             End If
             '*************
            'GRAVA DEBITOTRIBUTO
             If aTributos(Y).nCodTributo <> 3 Then
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
                Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & aTributos(Y).nCodTributo & "," & IIf(.TextMatrix(x, 4) = "0", Virg2Ponto(CStr(aTributosU(Y).nValorTributo)), Virg2Ponto(CStr(aTributos(Y).nValorTributo))) & ")"
                cn.Execute Sql, rdExecDirect
             End If
             Sql = "SELECT CODTRIBUTO,DESCTRIBUTO FROM TRIBUTO WHERE CODTRIBUTO=" & aTributos(Y).nCodTributo
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                sFullTrib = sFullTrib & !desctributo & " - "
               .Close
             End With
ProximoTrib:
          Next
          If Len(sFullTrib) > 0 Then
            sFullTrib = Left(sFullTrib, Len(sFullTrib) - 2)
          End If
         'GRAVA NUMDOCUMENTO
          sDataEmissao = Format(Day(Now), "00") & "/" & Format(Month(Now), "00") & "/" & Format(nAnoEmissao, "0000")
          
          If bISSFixo And bTLL Then
             nNumDocEspecial = grdTemp.TextMatrix(x, 8)
          Else
             Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             nLastCod = RdoAux2!maximo + 1
             RdoAux2.Close
             nNumDocEspecial = nLastCod  'usado para exportação do PDF
             grdTemp.TextMatrix(x, 8) = nLastCod
             On Error Resume Next
             grdTemp2.TextMatrix(x, 8) = nLastCod
             On Error GoTo Erro
          End If
'          Sql = "SELECT * FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & Val(.TextMatrix(x, 8))
'          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'          If RdoAux2.RowCount = 0 Then
             Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
             Sql = Sql & nLastCod & ",'" & Format(sDataEmissao, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ")"
             On Error Resume Next
             cn.Execute Sql, rdExecDirect
             On Error GoTo Erro
'          End If
          'RdoAux2.Close
         'GRAVA PARCELADOCUMENTO
          Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.Text) & ","
          Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
          Sql = Sql & .TextMatrix(x, 5) & "," & .TextMatrix(x, 8) & ")"
          cn.Execute Sql, rdExecDirect
    Next
End With


'Exit Sub
'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
'

sCodReduz = Format(Val(txtCod.Text), "000000")
sNomeResp = lblProp.Caption
sTipoImposto = lblLanc.Caption
sEndImovel = lblRua.Caption
nNumImovel = lblNumImovel.Caption
sComplImovel = lblCompl.Caption
sBairroImovel = lblBairro.Caption
nCodLogr = 0

sEndEntrega = lblRuaEntrega.Caption
nNumEntrega = lblNumEntrega.Caption
sBairroEntrega = lblBairroEntrega.Caption
sComplEntrega = lblCompl.Caption
sCepEntrega = lblCepEntrega.Caption
sCidadeEntrega = lblCidadeEntrega.Caption
sUFEntrega = lblUF.Caption

'GRAVA TEMPORARIO
With grdTemp2
    For x = 1 To .Rows - 1
        nAno = .TextMatrix(x, 1)
        nCodLanc = .TextMatrix(x, 2)
        nSeq = .TextMatrix(x, 3)
        nNumParc = .TextMatrix(x, 4)
        nComplemento = .TextMatrix(x, 5)
        dDataVencto = CDate(.TextMatrix(x, 6))
        sValorParc = .TextMatrix(x, 7)
        nNumdoc = .TextMatrix(x, 8)
        If CDbl(sValorParc) > 0 Then
'            NumBarra2 = Gera2of5Cod(CDbl(sValorParc) + nValorTxExpParc, dDataVencto, nNumDoc, nNumParc, nCodLanc, nSeq, nComplemento)
        Else
 '           NumBarra2 = Gera2of5Cod(0, dDataVencto, nNumDoc, nNumParc, nCodLanc, nSeq, nComplemento)
        End If
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))

        Sql = "INSERT CARNETMP(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,ENDIMOVEL,NUMIMOVEL,COMPLIMOVEL,"
        Sql = Sql & "BAIRROIMOVEL,ENDENTREGA,NUMENTREGA,COMPLENTREGA,BAIRROENTREGA,CEPENTREGA,CIDADEENTREGA,UFENTREGA,"
        Sql = Sql & "DESCIMPOSTO,EXERCICIO,NUMPROCESSO,DATAPROCESSO,NUMDOCUMENTO,DV,QUADRA,LOTE,DATAVENCTO,NUMPARCELA,"
        Sql = Sql & "NUMTOTPARCELA,VALORPARCELA,STRBARRA1,STRBARRA2,NUMBARRA1,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
        Sql = Sql & "DADOSLANCAMENTO,TAXAEXP,SAIR,OBS) VALUES('" & NomeDoUsuario & "'," & x & ",'" & lblNumInsc.Caption & "','" & sCodReduz & "','"
        Sql = Sql & Left$(sTipoImposto, 15) & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left$(sEndImovel, 40) & "'," & nNumImovel & ",'" & Mask(Left$(sComplImovel, 30)) & "','"
        Sql = Sql & Left(sBairroImovel, 25) & "','" & Left(sEndEntrega, 40) & "'," & nNumEntrega & ",'" & Mask(Left$(sComplEntrega, 30)) & "','" & Left(sBairroEntrega, 25) & "','"
        Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & sDescImposto & "'," & nAno & ",'" & sNumProc & "','"
        Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "','" & CStr(nNumdoc) & "','" & CStr(RetornaDVNumDoc(nNumdoc)) & "','" & sQuadra & "','"
        Sql = Sql & sLote & "','" & Format(dDataVencto, "mm/dd/yyyy") & "'," & IIf(nNumParc = 0, 1, nNumParc) & "," & IIf(nNumParc = 0, 1, IIf(chkUnica.value = 0, Val(grdTemp2.Rows - 1), Val(grdTemp2.Rows - 2))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
        Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & sDadosLanc & "'," & Virg2Ponto(CStr(nValorTxExpParc)) & "," & "0" & ",'" & Mask(Trim(txtObs.Text)) & "')"
        cn.Execute Sql, rdExecDirect

        modLg "Emissão de guia nº " & nNumdoc & " Codigo: " & Val(txtCod.Text) & " - " & lblProp.Caption
    Next
End With
bGerado = True

'EXIBE RELATORIO

frmReport.ShowReport "Carne", frmMdi.HWND, Me.HWND, nNumDocEspecial, nNumDocEspecial
'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
     
     MsgBox rdoErrors(x).Description

Next
Resume Next
End Sub

Private Sub FillTotal()
Dim nTotal As Double, nTotalUnica As Double, nValorTaxa As Double, Sql As String, RdoAux As rdoResultset, nValorAbate As Double
nTotal = 0: nTotalUnica = 0

For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
        If lvTrib.ListItems(x).SubItems(1) = "" Then lvTrib.ListItems(x).SubItems(1) = "0"
        nTotal = nTotal + (CDbl(lvTrib.ListItems(x).SubItems(1)) * CDbl(lvTrib.ListItems(x).SubItems(2)))
        If Val(Right$(lvTrib.ListItems(x).Key, 3)) = 23 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 24 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 28 Or Val(Right$(lvTrib.ListItems(x).Key, 3)) = 22 Then
        Else
            nTotalUnica = nTotalUnica + (CDbl(lvTrib.ListItems(x).SubItems(1)) * CDbl(lvTrib.ListItems(x).SubItems(2)))
        End If
    End If
Next
If chkTxExp.value = vbChecked Then
    Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         nValorTaxa = FormatNumber(!VALORPARCELA, 2)
        .Close
    End With
Else
    nValorTaxa = 0
End If

If txtAbate.Text = "" Or txtAbate.Text = "," Then
   nValorAbate = 0
Else
    nValorAbate = CDbl(txtAbate.Text)
End If
lblTotal.Caption = FormatNumber(nTotal + nValorTaxa - nValorAbate, 2)
lblTotalUnica.Caption = FormatNumber(nTotalUnica + nValorTaxa - nValorAbate, 2)

End Sub

Private Function CalculaISS(sTipo As String, nCodTributo As Integer) As Double
Dim nValorAliquotaISS As Double, nValorEstimado As Double, nAliq As Double
Dim nValorExpediente As Double, nUfirAtual As Double, nValorTotal As Double
Dim x As Integer

'CARREGA VALOR ATIVIDADE
Sql = "SELECT DISTINCT MOBILIARIOATIVIDADEISS.CODMOBILIARIO,MOBILIARIOATIVIDADEISS.CODTRIBUTO,MOBILIARIOATIVIDADEISS.CODATIVIDADE,"
Sql = Sql & "MOBILIARIOATIVIDADEISS.QTDEISS,MOBILIARIOATIVIDADEISS.VALORISS,"
Sql = Sql & "TABELAISS.ALIQUOTA FROM MOBILIARIOATIVIDADEISS INNER JOIN TABELAISS ON MOBILIARIOATIVIDADEISS.CODTRIBUTO = TABELAISS.TIPOISS AND "
Sql = Sql & "MOBILIARIOATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV Where MOBILIARIOATIVIDADEISS.CODMOBILIARIO = " & Val(txtCod.Text)
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        Do Until .EOF
            For x = 1 To UBound(aISS)
                If aISS(x) = !codatividade Then
                    If chkCalculo2010.value = vbUnchecked Then
                        nAliq = RetornaAliquotaISS(!codatividade, Format(Now, "dd/mm/yyyy"))
                    Else
                        nAliq = RetornaAliquotaISS(!codatividade, Format("01/01/2010", "dd/mm/yyyy"))
                    End If
                    'nValorAliquotaISS = nValorAliquotaISS + FormatNumber(!Aliquota * IIf(!QTDEISS = 0, 1, !QTDEISS), 2)
                    nValorAliquotaISS = nValorAliquotaISS + FormatNumber(nAliq * IIf(!QTDEISS = 0, 1, !QTDEISS), 2)
                    'nValorEstimado = nValorEstimado + FormatNumber(!valoriss * !Aliquota * IIf(!QTDEISS = 0, 1, !QTDEISS), 2)
                    nValorEstimado = nValorEstimado + FormatNumber(!valoriss * nAliq * IIf(!QTDEISS = 0, 1, !QTDEISS), 2)
                End If
            Next
           .MoveNext
        Loop
    End If
   .Close
End With
'EXPEDIENTE FUNCIONARIO
Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE ANO=" & Year(Now) & " AND CODTRIBUTO=" & nCodTributo
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        nValorExpediente = FormatNumber(!valoraliq, 2)
    End If
   .Close
End With
'UFIR ATUAL
If IsDate(mskDataInicio.Text) Then
    nUfirAtual = RetornaUFIR(Year(mskDataInicio.Text))
Else
    If chkCalculo2010.value = vbUnchecked Then
        nUfirAtual = RetornaUFIR(Year(Now))
    Else
        nUfirAtual = RetornaUFIR(2010)
    End If
End If

Select Case sTipo
    Case "F"
        nValorTotal = nValorAliquotaISS * nUfirAtual
        nValorTotal = nValorTotal * MesesProporcional / 12
        lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1) = "1"
    Case "E"
        nValorTotal = FormatNumber(nValorEstimado * nUfirAtual * Val(txtNumParc.Text), 2)
        'nValortotal = FormatNumber(nValorEstimado * nUfirAtual, 2)
        lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1) = "1"
    Case "V"
        nValorTotal = 0
        lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(1) = "0"
        lvTrib.ListItems(lvTrib.SelectedItem.Index).SubItems(2) = nValorTotal
End Select
CalculaISS = FormatNumber(nValorTotal, 2)

End Function

Private Function MesesProporcional() As Integer
Dim sDataAtual As String, sDataFim As String
Dim dData As Date

If Not IsDate(mskDataInicio.Text) Then
   dData = CDate(Format(Now, "dd/mm/yyyy"))
Else
   dData = CDate(mskDataInicio.Text)
End If

sDataAtual = Format(dData, "dd/mm/yyyy")
sDataFim = Format(dData, "31/12/yyyy")

If chkCalculo2010.value = vbUnchecked Then
    MesesProporcional = DateDiff("m", CDate(sDataAtual), CDate(sDataFim)) + 1
Else
    MesesProporcional = 12
End If

End Function

Private Sub txtNumParc_Change()
grdData.Rows = 1
If Val(txtNumParc.Text) <= 1 Then
   chkUnica.value = 0
End If
End Sub

Private Sub txtNumParc_KeyPress(KeyAscii As Integer)
Tweak txtNumParc, KeyAscii, IntegerPositive
End Sub

Public Function TemISSnoAno(nTipoIss As Integer) As Boolean

TemISSnoAno = False

Sql = "SELECT COUNT(*) AS TOTAL From DEBITOPARCELA WHERE CODREDUZIDO = " & Val(txtCod.Text)
Sql = Sql & " AND ANOEXERCICIO = " & Year(Now) & " AND CODLANCAMENTO = " & nTipoIss
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If !Total > 0 Then
        TemISSnoAno = True
    End If
End With

End Function

Private Sub CarregaLista()
Dim itmX As ListItem, z As Long, sCnae As String

z = SendMessage(lvISS.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvTL.HWND, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvVS.HWND, LVM_DELETEALLITEMS, 0, 0)


'VS
'Sql = "SELECT  mobiliarioatividadevs2.codmobiliario,mobiliarioatividadevs2.secao,mobiliarioatividadevs2.divisao,mobiliarioatividadevs2.grupo,"
'Sql = Sql & "mobiliarioatividadevs2.classe,mobiliarioatividadevs2.subclasse,cnaesubclasse.descricao,mobiliarioatividadevs2.criterio,"
'Sql = Sql & "mobiliarioatividadevs2.qtde,mobiliarioatividadevs2.valor,cnaecriteriodesc.descricao as desc2 From  mobiliarioatividadevs2 INNER JOIN "
'Sql = Sql & "cnaesubclasse ON (mobiliarioatividadevs2.secao = cnaesubclasse.secao) AND (mobiliarioatividadevs2.divisao = cnaesubclasse.divisao) AND "
'Sql = Sql & "(mobiliarioatividadevs2.grupo = cnaesubclasse.grupo) AND (mobiliarioatividadevs2.classe = cnaesubclasse.classe) AND "
'Sql = Sql & "(mobiliarioatividadevs2.subclasse = cnaesubclasse.subclasse) INNER JOIN cnaecriteriodesc ON (mobiliarioatividadevs2.criterio = cnaecriteriodesc.criterio) "
'Sql = Sql & "WHERE mobiliarioatividadevs2.codmobiliario=" & Val(txtCod.Text)

On Error Resume Next
Sql = "SELECT mobiliariovs.codigo, mobiliariovs.cnae, mobiliariovs.criterio, mobiliariovs.qtde, cnae.descricao, cnaecriteriodesc.valor "
Sql = Sql & "FROM mobiliariovs INNER JOIN cnae_criterio ON mobiliariovs.cnae = cnae_criterio.cnae INNER JOIN cnae ON mobiliariovs.cnae = cnae.cnae "
Sql = Sql & "INNER JOIN cnaecriteriodesc ON mobiliariovs.criterio = cnaecriteriodesc.criterio WHERE mobiliariovs.codigo = " & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
'        sCnae = Format(!divisao, "00") & !grupo & Left(Format(!classe, "00"), 1) & "-" & Right$(Format(!classe, "00"), 1) & "/" & Format(!subclasse, "00")
'        Set itmX = lvVS.ListItems.Add(, "VS" & sCnae & Format(!criterio, "00"), sCnae)
'        itmX.SubItems(1) = Format(!criterio, "00")
'        itmX.SubItems(2) = !descricao
'        itmX.SubItems(3) = !qtde
'        itmX.SubItems(4) = !Valor
        Set itmX = lvVS.ListItems.Add(, "VS" & !Cnae & Format(!criterio, "00"), !Cnae)
        itmX.SubItems(1) = Format(!criterio, "00")
        itmX.SubItems(2) = !descricao
        itmX.SubItems(3) = !QTDE
        itmX.SubItems(4) = !Valor
       .MoveNext
    Loop
   .Close
End With
On Error GoTo 0

'TX.LIC.
Sql = "SELECT MOBILIARIO.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE FROM MOBILIARIO INNER JOIN "
Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE "
Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       Set itmX = lvTL.ListItems.Add(, "IS" & Format(!codatividade, "000"), Format(!codatividade, "000"))
       itmX.SubItems(1) = !descatividade
      .MoveNext
    Loop
End With
On Error Resume Next
Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE FROM ATIVIDADE INNER JOIN "
Sql = Sql & "MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
Sql = Sql & "WHERE MOBILIARIOATIVIDADETL.CODIGOMOB =" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       Set itmX = lvTL.ListItems.Add(, "IS" & Format(!codatividade, "000"), Format(!codatividade, "000"))
       itmX.SubItems(1) = !descatividade
      .MoveNext
    Loop
End With

End Sub

Private Sub txtNumParc_LostFocus()
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
       If Val(Right$(lvTrib.ListItems(x).Key, 3)) = 3 Then 'taxa expediente não divide
          lvTrib.ListItems(x).SubItems(1) = Val(txtNumParc.Text)
          FillTotal
          Exit For
       End If
    End If
Next

If Not IsDate(mskDataVencimento.Text) Then Exit Sub

df = ValidaFeriado(CDate(mskDataVencimento.Text))
If df = 1 Then
    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 2 Then
    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencimento.Text = Format(RetornaDiaUtil(CDate(mskDataVencimento.Text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 3 Then
    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
    Sql = Sql & " Where DIA = " & Day(CDate(mskDataVencimento.Text))
    Sql = Sql & " AND MES=" & Month(CDate(mskDataVencimento.Text)) & " AND ANO=" & Year(CDate(mskDataVencimento.Text))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencimento.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                mskDataVencimento.Text = RetornaDiaUtil(CDate(mskDataVencimento.Text))
            Else
                Exit Sub
            End If
          .Close
        End If
    End With
End If

If grdData.Rows = 1 Then
    grdData.Rows = Val(txtNumParc.Text) + 1
    AutoFillDate2
    For x = 1 To Val(txtNumParc.Text)
        grdData.TextMatrix(x, 0) = "Parcela " & Format(x, "00")
    Next
End If



End Sub

Private Sub AutoFillDate2()
Dim x As Integer, sData As String, sDiaIni As String
sData = mskDataVencimento.Text

If Not IsDate(sData) Then
    MsgBox "Data inválida", vbExclamation, "Atenção"
    Exit Sub
End If


sDiaIni = Left(sData, 2)
For x = 1 To Val(txtNumParc.Text)
    grdData.TextMatrix(x, 1) = sData
    sData = Format(DateAdd("m", 1, sData), "dd/mm/yyyy")

    sData = sDiaIni & "/" & Mid(sData, 4, 2) & "/" & Right(sData, 4)
Inicio:
    If Not IsDate(sData) Then
        sData = Format(Val(Left(sData, 2)) - 1, "00") & "/" & Mid(sData, 4, 2) & "/" & Right(sData, 4)
        GoTo Inicio
    End If
    'sData = RetornaDiaUtil(CDate(sData))
Next


End Sub

Private Sub MontaMenu()

   Set m_cMenuContrib = New cPopupMenu
   With m_cMenuContrib
      .hwndOwner = Me.HWND
      .GradientHighlight = True
      
      i = .AddItem("Mobiliário", "", 1, , , , , "mnuMob")
      .OwnerDraw(i) = True
      i = .AddItem("Imobiliário", "", 1, , , , , "mnuImob")
      .OwnerDraw(i) = True
      i = .AddItem("Outros", "", 1, , , , , "mnuOutros")
      .OwnerDraw(i) = True
   End With
   
End Sub

Private Sub GravaMulta()

Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataGuia As String, sDataVencto As String
Dim nCodReduz As Long, sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, nLinha As Integer
Dim sUsuario As String, nNumdoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date
Dim sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long, sDescImposto As String, nQtdeParc As Integer
Dim aTributos() As TRIBUTO, aTributosU() As TRIBUTO, nNumDocEspecial As Long, bAbateu As Boolean, nAnoEmissao As Integer
Dim bDesconto As Boolean, bISSFixo As Boolean, bTLL As Boolean, nValorISSFixo As Double, sDadosLanc As String
Dim nValorUnica As Double, nValorParc As Double, sValorBoleto As String, nValorParcUnica As Double

'RETORNA VALOR EXPEDIENTE
If chkTxExp.value = vbChecked Then
    Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nValorTaxa = RdoAux!VALORDAM
    RdoAux.Close
Else
    nValorTaxa = 0
End If

nCodReduz = Val(txtCod.Text)
nCodLanc = 16 'MULTA DE INFRAÇÃO
nAno = Val(txtAno.Text)
nCompl = 0

ReDim aTributos(0): ReDim aTributosU(0)
nQtdeParc = 1
nValorParc = 0
nValorParcUnica = 0
bAbateu = False
grdTemp.Rows = 1: grdTemp2.Rows = 1

'CALCULA O VALOR PARCELADO
'If CDbl(lblArea.Caption) <= 125 Then
'    nValorTotal = 65.75
'ElseIf CDbl(lblArea.Caption) > 125 And CDbl(lblArea.Caption) <= 250 Then
'    nValorTotal = 164.37
'ElseIf CDbl(lblArea.Caption) > 250 And CDbl(lblArea.Caption) <= 500 Then
'    nValorTotal = 262.99
'ElseIf CDbl(lblArea.Caption) > 500 Then
'    nValorTotal = 394.49
'End If
nValorTotal = 500

'nValorParc = FormatNumber(nValorTotal, 2)

'nValorTotal = CDbl(lblTotal.Caption) - nValorTaxa
nValorUnica = 0

'MONTA TRIBUTOS
sDadosLanc = "MULTA DE INFRAÇÃO"
bDesconto = False
ReDim Preserve aTributos(UBound(aTributos) + 1)
aTributos(UBound(aTributos)).nCodTributo = 20
aTributos(UBound(aTributos)).nValorTributo = nValorTotal

sDadosLanc = sDadosLanc & "TX.EXPEDIENTE"
ReDim Preserve aTributos(UBound(aTributos) + 1)
aTributos(UBound(aTributos)).nCodTributo = 3
aTributos(UBound(aTributos)).nValorTributo = nValorTaxa

'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nLastCod = RdoAux!maximo
nNumGuia = nLastCod
RdoAux.Close
nNumDocEspecial = nLastCod + 1 'usado para exportação do PDF

nNumGuia = nNumGuia + 1
sDataVencto = grdData.TextMatrix(nParc, 1)
If sDataVencto = "" Or sDataVencto = "Data" Then sDataVencto = mskDataVencimento.Text

sValorBoleto = FormatNumber(nValorTotal / nQtdeParc, 2) + nValorTaxa

'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & 1
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux2!SEQMAXIMA) Then
   nSeq = 0
Else
   nSeq = RdoAux2!SEQMAXIMA + 1
End If
grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & 1 & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & sValorBoleto & Chr(9) & nNumGuia
grdTemp2.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & 1 & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & sValorBoleto & Chr(9) & nNumGuia

'GERAÇÃO DOS DÉBITOS
With grdTemp
    For x = 1 To .Rows - 1
       'GRAVA DEBITOPARCELA
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
'        Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
'        Sql = Sql & 3 & ",'" & Format(mskDataVencimento.Text, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
'        Sql = Sql & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
        Sql = Sql & 3 & ",'" & Format(mskDataVencimento.Text, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
        Sql = Sql & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect

        If Trim(txtObs.Text) <> "" Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
            Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            sDataGuia = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & .TextMatrix(x, 0) & ","
'            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(txtObs.Text)) & "','"
'            Sql = Sql & NomeDeLogin & "','" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & .TextMatrix(x, 0) & ","
            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(txtObs.Text)) & "',"
            Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
        End If

        If Trim(sObsIss) <> "" Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
            Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            sDataGuia = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & .TextMatrix(x, 0) & ","
'            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(sObsIss)) & "','"
'            Sql = Sql & NomeDeLogin & "','" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & .TextMatrix(x, 0) & ","
            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(sObsIss)) & "',"
            Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
        
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodigoImovel
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            
'            Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'            Sql = Sql & nCodigoImovel & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "','GTI','" & Format(Now, "mm/dd/yyyy") & "')"
            Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
            Sql = Sql & nCodigoImovel & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
        
        End If

        sFullTrib = ""
        For Y = 1 To UBound(aTributos)
           'GRAVA DEBITOTRIBUTO
            If aTributos(Y).nCodTributo <> 3 Then
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
                Sql = Sql & aTributos(Y).nCodTributo & "," & Virg2Ponto(CStr(aTributos(Y).nValorTributo)) & ")"
                cn.Execute Sql, rdExecDirect
             End If
             Sql = "SELECT CODTRIBUTO,DESCTRIBUTO FROM TRIBUTO WHERE CODTRIBUTO=" & aTributos(Y).nCodTributo
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                sFullTrib = sFullTrib & !desctributo & " - "
               .Close
             End With
ProximoTrib:
        Next
        If Len(sFullTrib) > 0 Then
          sFullTrib = Left(sFullTrib, Len(sFullTrib) - 2)
        End If

       'GRAVA NUMDOCUMENTO
        sDataGuia = Format(Day(Now), "00") & "/" & Format(Month(Now), "00") & "/" & Format(nAno, "0000")

        If bISSFixo And bTLL Then
           nNumDocEspecial = grdTemp.TextMatrix(x, 8)
        Else
           Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           nNumGuia = RdoAux2!maximo + 1
           RdoAux2.Close
           nNumDocEspecial = nLastCod  'usado para exportação do PDF
           grdTemp.TextMatrix(x, 8) = nNumGuia
           grdTemp2.TextMatrix(x, 8) = nNumGuia
        End If

        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
        Sql = Sql & grdTemp.TextMatrix(x, 8) & ",'" & Format(sDataGuia, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTaxa)) & ")"
        On Error Resume Next
        cn.Execute Sql, rdExecDirect
        On Error GoTo Erro

       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.Text) & ","
        Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
        Sql = Sql & .TextMatrix(x, 5) & "," & .TextMatrix(x, 8) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

Select Case nCodReduz
    Case 1 To 99999
        Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !Inscricao
            sNome = !nomecidadao
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
                If sDoc = "" Then
                    sDoc = SubNull(!rg)
                End If
            End If
            sEnd = SubNull(!Logradouro)
            nNum = Val(SubNull(!Li_Num))
            sCompl = Left(SubNull(!Li_Compl), 30)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!li_uf)
           .Close
        End With
    Case 100000 To 500000
        Sql = "SELECT * FROM vwFULLEMPRESA3 WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !inscestadual
            sNome = !razaosocial
            If Not IsNull(!CPF) Then
               sDoc = !CPF
            ElseIf Not IsNull(!Cnpj) Then
               sDoc = !Cnpj
            Else
                sDoc = ""
            End If
            sEnd = !Logradouro
            nNum = !Numero
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!SiglaUF)
         End With
     Case 500000 To 800000
        Sql = "SELECT * from vwFULLCIDADAO WHERE CODCIDADAO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = ""
            sNome = !nomecidadao
            If SubNull(!Cnpj) <> "" Then
               sDoc = !Cnpj
            ElseIf SubNull(!CPF) <> "" Then
               sDoc = !CPF

            Else
                sDoc = SubNull(!rg)
            End If
            sEnd = SubNull(!Endereco)
            nNum = Val(SubNull(!NUMIMOVEL))
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!SiglaUF)
        End With
End Select

nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect


'GRAVA TEMPORARIO
With grdTemp2
    For nLinha = 1 To .Rows - 1
'        nAno = .TextMatrix(nLinha, 1)
'        nCodLanc = .TextMatrix(nLinha, 2)
'        nSeq = .TextMatrix(nLinha, 3)
'        nParc = .TextMatrix(nLinha, 4)
'        nCompl = .TextMatrix(nLinha, 5)
'        dDataVencto = CDate(mskDataVencimento.Text)
'        sValorParc = .TextMatrix(nLinha, 7)
 '       nValorGuia = CDbl(sValorParc) / 2
 '       nValorDoc = nValorGuia
'        nNumGuia = .TextMatrix(nLinha, 8)'
'
'        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
 '       sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
  '      sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)

'****************
        'ARRECADAÇÃO
        sValor = CDbl(.TextMatrix(nLinha, 7)) / 2
        dDataVencto = CDate(mskDataVencimento.Text)
        nNumGuia = .TextMatrix(nLinha, 8)
        nNumdoc = nNumGuia
        sDadosLanc = "MULTA DE ROÇADA"
        NumBarra2 = Gera2of5Cod(CStr(sValor), CDate(dDataVencto), nNumdoc, nCodReduz)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
        
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        sBarra = StrBarra2



        '**** GERADOR DE CÓDIGO DE BARRAS ********
'        sNossoNumero = "2678478"
'        sDigitavel = "001900000"
'        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'        sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
'        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'        sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
'        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'        sDigitavel = sDigitavel & sDv
'
'        dDataBase = "07/10/1997"
'        nFatorVencto = dDataVencto - dDataBase
'        'nFatorVencto = CDate(sDataDam) - dDataBase
'        sQuintoGrupo = Format(nFatorVencto, "0000")
'        sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
'        sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "00000026784780"
'        sBarra = sBarra & sNumDoc3 & "18"
'        sDv = Trim(Calculo_DV11(sBarra))
'        sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
'
'        sDigitavel = sDigitavel & sDv & sQuintoGrupo
'
'        sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
'        sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
'        sBarra = Gera2of5Str(sBarra)

        
        Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,cep,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nLinha & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & lblNumInsc.Caption & "','"
        Sql = Sql & Left(Mask(sEnd), 80) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Mask(sCidade) & "','" & sCep & "','" & sUF & "','" & Mask(sDadosLanc) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & 1 & "," & 1 & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & nNumdoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(sValor, "#0.00")) & ",'" & Mask(txtObs.Text) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
        cn.Execute Sql, rdExecDirect
     
        Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nLinha & ","
        Sql = Sql & 16 & ",'" & "MULTA DE INFRAÇÃO" & "'," & 250 & ")"
        cn.Execute Sql, rdExecDirect
        
     '   Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
     '   Sql = Sql & "digitavel,codbarra,valorguia,obs) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nLinha & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & "" & "','"
     '   Sql = Sql & Left(Mask(sEnd), 80) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Mask(sCidade) & "','" & sUF & "','" & Mask(sDadosLanc) & "','"
     '   Sql = Sql & CStr(nNumGuia) & "'," & 1 & "," & 1 & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & nNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
     '   Sql = Sql & Virg2Ponto(Format((nValorTotal / 2) + nValorTaxa, "#0.00")) & ",'" & Mask(txtObs.Text) & "')"
     '   cn.Execute Sql, rdExecDirect

        modLg "Emissão de guia nº " & nNumGuia & " Codigo: " & Val(txtCod.Text) & " - " & lblProp.Caption
    Next
End With

If chkTxExp.value = vbChecked Then
    Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & 1 & ","
    Sql = Sql & 3 & ",'" & "MULTA INF" & "'," & Virg2Ponto(RemovePonto(Format(nValorTotal / 2, "#0.00"))) & ")"
    cn.Execute Sql, rdExecDirect
End If

bGerado = True

'EXIBE RELATORIO

'frmReport.ShowReport2 "BOLETOGUIA", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
frmReport.ShowReport2 "BOLETOGUIA_V4", frmMdi.HWND, Me.HWND, nSid, nNumGuia
Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(x).Description
Next
Resume Next

End Sub

Private Sub GravaMultaOld()
On Error GoTo Erro

Dim x As Integer, nAnoEmissao As Integer, sDataEmissao As String, nNumDocEspecial As Long, bAbateu As Boolean
Dim RdoAux2 As rdoResultset
Dim sNumInsc As String
Dim nCodReduz As Long
Dim sNomeResp As String
Dim sValorParc As String
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
Dim dDataVencto As Date
Dim nNumdoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim sVencimento As String
Dim nCodLanc As Integer
Dim nSeq As Integer, nSeq2 As Integer
Dim nComplemento As Integer
Dim nValorTotal As Double, nValorUnica As Double
Dim nValorParc As Double, sValorBoleto As String
Dim nValorParcUnica As Double
Dim aTributos() As TRIBUTO
Dim aTributosU() As TRIBUTO
Dim NumBarra1 As String, sCPF As String
Dim StrBarra1 As String
Dim NumBarra2 As String
Dim NumBarra2a As String
Dim NumBarra2b As String
Dim NumBarra2c As String
Dim NumBarra2d As String
Dim StrBarra2 As String
Dim nLastCod As Long
Dim sDadosLanc As String
Dim sFullTrib As String, bDesconto As Boolean, bISSFixo As Boolean, bTLL As Boolean, nValorISSFixo As Double

nAnoEmissao = Val(txtAno.Text)
nValorParc = 0

nCodReduz = Val(txtCod.Text)
Sql = "SELECT CODREDUZIDO,CPF,CNPJ,RG,ORGAO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        If Not IsNull(!CPF) Then
           sCPF = !CPF
        ElseIf Not IsNull(!Cnpj) Then
           sCPF = !Cnpj
        ElseIf Not IsNull(!rg) Then
           sCPF = !rg
        Else
           sCPF = ""
        End If
    End If
End With

nCodLanc = 16 'MULTA DE INFRAÇÃO

ReDim aTributos(0)

If chkTxExp.value = vbChecked Then
    Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         nValorTxExpParc = FormatNumber(!VALORPARCELA, 2)
        .Close
    End With
Else
    nValorTxExpParc = 0
End If
'CALCULA O VALOR PARCELADO
If CDbl(lblArea.Caption) <= 125 Then
    nValorTotal = 65.75
ElseIf CDbl(lblArea.Caption) > 125 And CDbl(lblArea.Caption) <= 250 Then
    nValorTotal = 164.37
ElseIf CDbl(lblArea.Caption) > 250 And CDbl(lblArea.Caption) <= 500 Then
    nValorTotal = 262.99
ElseIf CDbl(lblArea.Caption) > 500 Then
    nValorTotal = 394.49
End If
nValorParc = FormatNumber(nValorTotal, 2)
'MONTA TRIBUTOS
sDadosLanc = ""
bDesconto = False
sDadosLanc = "MULTA DE INFRAÇÃO"
sVencimento = mskDataVencimento.Text
nNumParc = 1
nComplemento = 0
nAno = nAnoEmissao
sValorBoleto = FormatNumber(nValorTotal, 2)
txtObs.Text = txtObs.Text & " Multa ja reduzida em 50% até o vencimento. NÃO RECEBER APÓS O VENCIMENTO."
    
    
'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nNumParc
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux2!SEQMAXIMA) Then
   nSeq = 0
Else
   nSeq = RdoAux2!SEQMAXIMA + 1
End If

'DADOS CABEÇALHO
sNumProc = NumeroProcesso
dDataProc = Format(Now, "dd/mm/yyyy")
sDescImposto = "MULTA INF."
NumBarra1 = Format(ExtraiNumero(sNumProc), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)


'GRAVA DEBITOPARCELA    // (STATUS 3 - NAO PAGO)
'Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
'Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & nCodReduz & "," & nAno & ","
'Sql = Sql & nCodLanc & "," & nSeq & "," & 1 & "," & nComplemento & ","
'Sql = Sql & 3 & ",'" & Format(sVencimento, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
'Sql = Sql & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & nCodReduz & "," & nAno & ","
Sql = Sql & nCodLanc & "," & nSeq & "," & 1 & "," & nComplemento & ","
Sql = Sql & 3 & ",'" & Format(sVencimento, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
Sql = Sql & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect

If Trim(txtObs.Text) <> "" Then
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
    Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & 1 & " AND CODCOMPLEMENTO=" & nComplemento
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeq2 = 1
        Else
            nSeq2 = !maximo + 1
        End If
        .Close
    End With
    sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduz & ","
'    Sql = Sql & nAno & "," & nCodLanc & "," & nSeq & "," & 1 & "," & nComplemento & "," & nSeq2 & ",'" & Mask(Trim(txtObs.Text)) & "','"
'    Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & ","
    Sql = Sql & nAno & "," & nCodLanc & "," & nSeq & "," & 1 & "," & nComplemento & "," & nSeq2 & ",'" & Mask(Trim(txtObs.Text)) & "',"
    Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
End If
          
sFullTrib = ""
'GRAVA DEBITOTRIBUTO
Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
Sql = Sql & nCodReduz & "," & nAno & "," & nCodLanc & "," & nSeq & ","
Sql = Sql & 1 & "," & nComplemento & "," & 20 & "," & Virg2Ponto(CStr(nValorParc)) & ")"
cn.Execute Sql, rdExecDirect
sFullTrib = "MULTA DE INFRAÇÃO"

sDataEmissao = Format(Day(Now), "00") & "/" & Format(Month(Now), "00") & "/" & Format(nAnoEmissao, "0000")
          
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nLastCod = RdoAux2!maximo + 1
RdoAux2.Close
nNumDocEspecial = nLastCod  'usado para exportação do PDF
          
Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
Sql = Sql & nLastCod & ",'" & Format(sDataEmissao, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ")"
cn.Execute Sql, rdExecDirect
'GRAVA PARCELADOCUMENTO
Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
Sql = Sql & nAno & "," & nCodLanc & "," & nSeq & "," & 1 & ","
Sql = Sql & nComplemento & "," & nLastCod & ")"
cn.Execute Sql, rdExecDirect

'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

sCodReduz = Format(Val(txtCod.Text), "000000")
sNomeResp = lblProp.Caption
sTipoImposto = "MULTA DE INF."
sEndImovel = lblRua.Caption
nNumImovel = lblNumImovel.Caption
sComplImovel = lblCompl.Caption
sBairroImovel = lblBairro.Caption
nCodLogr = 0

sEndEntrega = lblRuaEntrega.Caption
nNumEntrega = lblNumEntrega.Caption
sBairroEntrega = lblBairroEntrega.Caption
sComplEntrega = lblCompl.Caption
sCepEntrega = lblCepEntrega.Caption
sCidadeEntrega = lblCidadeEntrega.Caption
sUFEntrega = lblUF.Caption

'GRAVA TEMPORARIO
nValorParc = FormatNumber(nValorParc / 2, 2)
nNumParc = 1
nComplemento = 0
sValorParc = nValorParc
nNumdoc = nLastCod
If CDbl(sValorParc) > 0 Then
    'NumBarra2 = Gera2of5Cod(CDbl(sValorParc) + nValorTxExpParc, CDate(sVencimento), nNumDoc, nNumParc, nCodLanc, nSeq, nComplemento)
Else
    'NumBarra2 = Gera2of5Cod(0, dDataVencto, nNumDoc, nNumParc, nCodLanc, nSeq, nComplemento)
End If
NumBarra2a = Left$(NumBarra2, 13)
NumBarra2b = Mid$(NumBarra2, 14, 13)
NumBarra2c = Mid$(NumBarra2, 27, 13)
NumBarra2d = Right$(NumBarra2, 13)
StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))

Sql = "INSERT CARNETMP(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,ENDIMOVEL,NUMIMOVEL,COMPLIMOVEL,"
Sql = Sql & "BAIRROIMOVEL,ENDENTREGA,NUMENTREGA,COMPLENTREGA,BAIRROENTREGA,CEPENTREGA,CIDADEENTREGA,UFENTREGA,"
Sql = Sql & "DESCIMPOSTO,EXERCICIO,NUMPROCESSO,DATAPROCESSO,NUMDOCUMENTO,DV,QUADRA,LOTE,DATAVENCTO,NUMPARCELA,"
Sql = Sql & "NUMTOTPARCELA,VALORPARCELA,STRBARRA1,STRBARRA2,NUMBARRA1,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
Sql = Sql & "DADOSLANCAMENTO,TAXAEXP,SAIR,OBS) VALUES('" & NomeDoUsuario & "'," & 1 & ",'" & lblNumInsc.Caption & "','" & sCodReduz & "','"
Sql = Sql & Left$(sTipoImposto, 15) & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left$(sEndImovel, 40) & "'," & nNumImovel & ",'" & Mask(Left$(sComplImovel, 30)) & "','"
Sql = Sql & Left(sBairroImovel, 25) & "','" & Left(sEndEntrega, 40) & "'," & nNumEntrega & ",'" & Mask(Left$(sComplEntrega, 30)) & "','" & Left(sBairroEntrega, 25) & "','"
Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & sDescImposto & "'," & nAno & ",'" & sNumProc & "','"
Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "','" & CStr(nNumdoc) & "','" & "" & "','" & sQuadra & "','"
Sql = Sql & sLote & "','" & Format(sVencimento, "mm/dd/yyyy") & "'," & nNumParc & "," & 1 & ","
Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & sDadosLanc & "'," & Virg2Ponto(CStr(nValorTxExpParc)) & "," & "0" & ",'" & Mask(Trim(txtObs.Text)) & "')"
cn.Execute Sql, rdExecDirect

modLg "Emissão de guia (Multa) nº " & nNumdoc & " Codigo: " & Val(txtCod.Text) & " - " & lblProp.Caption
bGerado = True

'EXIBE RELATORIO

frmReport.ShowReport "Carne", frmMdi.HWND, Me.HWND, nNumDocEspecial
'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
     
     MsgBox rdoErrors(x).Description
Next
Resume Next
End Sub

Private Sub EmiteBoleto()

Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataGuia As String, sDataVencto As String
Dim nCodReduz As Long, sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, nLinha As Integer, sCep As String
Dim sUsuario As String, nNumdoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date
Dim sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long, sDescImposto As String, nQtdeParc As Integer
Dim aTributos() As TRIBUTO, aTributosU() As TRIBUTO, nNumDocEspecial As Long, bAbateu As Boolean, nAnoEmissao As Integer
Dim bDesconto As Boolean, bISSFixo As Boolean, bTLL As Boolean, nValorISSFixo As Double, sDadosLanc As String
Dim nValorUnica As Double, nValorParc As Double, sValorBoleto As String, nValorParcUnica As Double
Dim sValor As String, dDataVencto As Date, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String

If MsgBox("Confirma emissão da guia?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

'RETORNA VALOR EXPEDIENTE
If chkTxExp.value = vbChecked Then
    Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nValorTaxa = RdoAux!VALORDAM
    RdoAux.Close
Else
    nValorTaxa = 0
End If

nCodReduz = Val(txtCod.Text)
nCodLanc = cmbLanc.ItemData(cmbLanc.ListIndex)
nAno = Val(txtAno.Text)
nCompl = 0

ReDim aTributos(0): ReDim aTributosU(0)
nQtdeParc = Val(txtNumParc.Text)
nValorParc = 0
nValorParcUnica = 0
bAbateu = False
grdTemp.Rows = 1: grdTemp2.Rows = 1

'CALCULA O VALOR PARCELADO
nValorTotal = CDbl(lblTotal.Caption) - nValorTaxa
nValorUnica = CDbl(lblTotalUnica.Caption) - nValorTaxa
nValorParc = FormatNumber(nValorTotal / CDbl(txtNumParc.Text), 2)
If Not IsNumeric(txtAbate.Text) Then txtAbate.Text = 0

'MONTA TRIBUTOS
sDadosLanc = ""
bDesconto = False
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
       ReDim Preserve aTributos(UBound(aTributos) + 1)
       ReDim Preserve aTributosU(UBound(aTributosU) + 1)
       aTributosU(UBound(aTributosU)).nCodTributo = Val(Right$(lvTrib.ListItems(x).Key, 3))
       aTributos(UBound(aTributos)).nCodTributo = Val(Right$(lvTrib.ListItems(x).Key, 3))

       If chkUnica.value = 0 Then
          If Not bAbateu Then
             aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber((CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1))) - CDbl(txtAbate.Text), 2)
          Else
             aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber(CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1)), 2)
          End If
       Else
          aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber(CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1)), 2)
          If aTributosU(UBound(aTributosU)).nCodTributo = 22 Or aTributosU(UBound(aTributosU)).nCodTributo = 23 Or aTributosU(UBound(aTributosU)).nCodTributo = 24 Or aTributosU(UBound(aTributosU)).nCodTributo = 28 Then
          Else
'              nValorUnica = nValorUnica + aTributosU(x).nValorTributo
             aTributosU(UBound(aTributosU)).nValorTributo = aTributosU(UBound(aTributosU)).nValorTributo - aTributosU(UBound(aTributosU)).nValorTributo * 0.05
          End If
       End If
       If Not bAbateu Then
            aTributos(UBound(aTributos)).nValorTributo = FormatNumber(((CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1))) - CDbl(txtAbate.Text)) / Val(txtNumParc.Text), 2)
            bAbateu = True
       Else
            aTributos(UBound(aTributos)).nValorTributo = FormatNumber((CDbl(lvTrib.ListItems(x).SubItems(2) * CDbl(lvTrib.ListItems(x).SubItems(1))) / Val(txtNumParc.Text)), 2)
       End If
       sDadosLanc = sDadosLanc & lvTrib.ListItems(x).Text & vbCrLf
    End If
Next

If chkTxExp.value = vbChecked Then
    sDadosLanc = sDadosLanc & "TX.EXPEDIENTE" & vbCrLf
    ReDim Preserve aTributos(UBound(aTributos) + 1)
    ReDim Preserve aTributosU(UBound(aTributosU) + 1)
    aTributosU(UBound(aTributosU)).nCodTributo = 3
    aTributos(UBound(aTributos)).nCodTributo = 3
    aTributosU(UBound(aTributosU)).nValorTributo = nValorTaxa
    aTributos(UBound(aTributos)).nValorTributo = nValorTaxa
End If

nValorUnica = 0
For x = 0 To UBound(aTributos)
    If aTributosU(x).nCodTributo <> 3 Then
        nValorUnica = nValorUnica + aTributosU(x).nValorTributo
    End If
Next

'**************************************************
'verifica se é um caso de ISS Fixo/TLL
'se tiver devem ser divididos
If cmbLanc.ItemData(cmbLanc.ListIndex) = 2 Then
    bISSFixo = False: bTLL = False
    For Y = 1 To UBound(aTributos)
        If aTributos(Y).nCodTributo = 11 Then
            bISSFixo = True
            nValorISSFixo = aTributos(Y).nValorTributo
        End If
        If aTributos(Y).nCodTributo = 14 Then bTLL = True
    Next
End If
'**************************************************

'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nLastCod = RdoAux!maximo
nNumGuia = nLastCod
RdoAux.Close
nNumDocEspecial = nLastCod + 1 'usado para exportação do PDF

For nParc = 0 To Val(txtNumParc.Text)
    nNumGuia = nNumGuia + 1
    If nParc = 0 And chkUnica.value = 0 Then GoTo proximo 'VALIDA PARCELA ÚNICA
    If nParc > 0 Then
        If grdData.Rows = 1 Then
            sDataVencto = mskDataVencimento.Text
        Else
            sDataVencto = grdData.TextMatrix(nParc, 1)
        End If
    Else
        sDataVencto = mskDataVencimento.Text
    End If
    If sDataVencto = "" Then sDataVencto = mskDataVencimento.Text

    If nParc = 0 Then
     '  If bDesconto Then
          'ja foi dado desconto la em cima
          sValorBoleto = FormatNumber(nValorUnica, 2)
      ' Else
       '   sValorBoleto = FormatNumber(nValorUnica - (nValorUnica * 0.05), 2)
       'End If
    Else
       If chkTxExp.value = vbChecked Then
            sValorBoleto = FormatNumber(nValorTotal / nQtdeParc, 2) + nValorTaxa
       Else
            sValorBoleto = FormatNumber(nValorTotal / nQtdeParc, 2)
       End If
    End If

'    nNumGuia = nNumGuia + 1
    If bISSFixo And bTLL Then
        'Uma linha para iss fixo e outra para taxa de licença
         'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 14 & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & 14 & Chr(9) & nSeq & Chr(9) & nParc & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & nValorISSFixo & Chr(9) & nNumGuia

         'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 6 & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & 6 & Chr(9) & nSeq & Chr(9) & nParc & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & CDbl(sValorBoleto) - nValorISSFixo & Chr(9) & nNumGuia
        grdTemp2.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & 2 & Chr(9) & nSeq & Chr(9) & nParc & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & CDbl(sValorBoleto) & Chr(9) & nNumGuia
    Else
        'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
        grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & nParc & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & sValorBoleto & Chr(9) & nNumGuia
        grdTemp2.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & nParc & Chr(9) & nCompl & Chr(9) & sDataVencto & Chr(9) & sValorBoleto & Chr(9) & nNumGuia
    End If

proximo:
Next

'GERAÇÃO DOS DÉBITOS
With grdTemp
    For x = 1 To .Rows - 1
       'GRAVA DEBITOPARCELA
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
'        Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
'        Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
'        Sql = Sql & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
        Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
        Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
        Sql = Sql & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect

        If Trim(txtObs.Text) <> "" Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
            Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            sDataGuia = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & .TextMatrix(x, 0) & ","
'            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(txtObs.Text)) & "','"
'            Sql = Sql & NomeDeLogin & "','" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & .TextMatrix(x, 0) & ","
            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(txtObs.Text)) & "',"
            Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
        End If

        If Trim(sObsIss) <> "" Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
            Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            sDataGuia = Right$(frmMdi.Sbar.Panels(6).Text, 10)
            z = 2
            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & .TextMatrix(x, 0) & ","
            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(sObsIss)) & "',"
            Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataGuia, "mm/dd/yyyy") & "')"
'            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & .TextMatrix(x, 0) & ","
'            Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeq & ",'" & Mask(Trim(sObsIss)) & "','"
'            Sql = Sql & NomeDeLogin & "','" & Format(sDataGuia, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
        
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodigoImovel
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 1
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            
'            Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'            Sql = Sql & nCodigoImovel & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "','GTI','" & Format(Now, "mm/dd/yyyy") & "')"
            Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
            Sql = Sql & nCodigoImovel & "," & nSeq & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
            sObsIss = ""
        
        End If


        sFullTrib = ""
        For Y = 1 To UBound(aTributos)
           '*************
            If bISSFixo And bTLL Then
                If Val(.TextMatrix(x, 2)) = 14 Then 'se for iss fixo ignorar outros tributos que não sejam iss fixo(11)
                    If aTributos(Y).nCodTributo <> 11 Then
                        GoTo ProximoTrib
                    End If
                ElseIf Val(.TextMatrix(x, 2)) = 6 Then 'se for taxa licenca ignorar o tributo iss fixo(11)
                    If aTributos(Y).nCodTributo = 11 Then
                        GoTo ProximoTrib
                    End If
                End If
            End If
           '*************

           'GRAVA DEBITOTRIBUTO
            If aTributos(Y).nCodTributo <> 3 Then
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
                Sql = Sql & aTributos(Y).nCodTributo & "," & IIf(.TextMatrix(x, 4) = "0", Virg2Ponto(CStr(aTributosU(Y).nValorTributo)), Virg2Ponto(CStr(aTributos(Y).nValorTributo))) & ")"
                cn.Execute Sql, rdExecDirect
             End If
             Sql = "SELECT CODTRIBUTO,DESCTRIBUTO FROM TRIBUTO WHERE CODTRIBUTO=" & aTributos(Y).nCodTributo
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                sFullTrib = sFullTrib & !desctributo & " - "
               .Close
             End With
ProximoTrib:
        Next
        If Len(sFullTrib) > 0 Then
          sFullTrib = Left(sFullTrib, Len(sFullTrib) - 2)
        End If

       'GRAVA NUMDOCUMENTO
        sDataGuia = Format(Day(Now), "00") & "/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")

        If bISSFixo And bTLL Then
           nNumDocEspecial = grdTemp.TextMatrix(x, 8)
        Else
           Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           nNumGuia = RdoAux2!maximo + 1
           RdoAux2.Close
           nNumDocEspecial = nLastCod  'usado para exportação do PDF
           grdTemp.TextMatrix(x, 8) = nNumGuia
           grdTemp2.TextMatrix(x, 8) = nNumGuia
        End If

        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,EMISSOR) VALUES("
        Sql = Sql & grdTemp.TextMatrix(x, 8) & ",'" & Format(sDataGuia, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTaxa)) & ",'" & NomeDeLogin & "')"
        On Error Resume Next
        cn.Execute Sql, rdExecDirect
        On Error GoTo Erro


       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.Text) & ","
        Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
        Sql = Sql & .TextMatrix(x, 5) & "," & .TextMatrix(x, 8) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

sInsc = ""
sNome = lblProp.Caption
sEnd = lblRuaEntrega.Caption
nNum = lblNumEntrega.Caption
sCompl = lblComplentrega.Caption
sBairro = lblBairroEntrega.Caption
sCidade = lblCidadeEntrega.Caption
sUF = lblUF.Caption
sCep = lblCEP.Caption
nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect



'GRAVA TEMPORARIO
With grdTemp2
    For nLinha = 1 To .Rows - 1
        nAno = .TextMatrix(nLinha, 1)
        nCodLanc = .TextMatrix(nLinha, 2)
        nSeq = .TextMatrix(nLinha, 3)
        nParc = .TextMatrix(nLinha, 4)
        nCompl = .TextMatrix(nLinha, 5)
        dDataVencto = CDate(.TextMatrix(nLinha, 6))
        sValorParc = .TextMatrix(nLinha, 7)
        nValorGuia = CDbl(sValorParc)
        nNumGuia = .TextMatrix(nLinha, 8)
        nValorDoc = nValorGuia
        sDataDam = Format(dDataVencto, "dd/mm/yyyy")
        '*****
        'nNumGuia = 13982196
        'nValorDoc = 116.46
        'sDataDam = "15/05/2015"
        '******
        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)
        
'        bBoleto = True
 '       If NomeDeLogin = "SCHWARTZ" Then
            bBoleto = False
  '      End If
            
        If bBoleto Then
            '**** GERADOR DE CÓDIGO DE BARRAS ********
            sNossoNumero = "2873532"
            sDigitavel = "001900000"
            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
            sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
            sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
            sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
            sDigitavel = sDigitavel & sDv
            
            dDataBase = "07/10/1997"
            nFatorVencto = CDate(sDataDam) - dDataBase
            sQuintoGrupo = Format(nFatorVencto, "0000")
            sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
            sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "00000028735320"
            sBarra = sBarra & sNumDoc3 & "18"
            sDv = Trim(Calculo_DV11(sBarra))
            sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
            
            sDigitavel = sDigitavel & sDv & sQuintoGrupo
            
            sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
            sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
            sBarra = Gera2of5Str(sBarra)
            
            '*******************************************
        Else
            
            sValor = nValorDoc
            dDataVencto = CDate(sDataDam)
            nNumdoc = nNumGuia
            sDadosLanc = cmbLanc.Text
            NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumdoc, nCodReduz)
            NumBarra2a = Left$(NumBarra2, 13)
            NumBarra2b = Mid$(NumBarra2, 14, 13)
            NumBarra2c = Mid$(NumBarra2, 27, 13)
            NumBarra2d = Right$(NumBarra2, 13)
        
            StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
            sBarra = StrBarra2
        '    Sql = "update boletoguia set numbarra2a='" & NumBarra2a & "',numbarra2b='" & NumBarra2b & "',numbarra2c='" & NumBarra2c & "',numbarra2d='" & NumBarra2d & "',codbarra='" & Mask(StrBarra2) & "' where sid=" & nSid
        '    cn.Execute Sql, rdExecDirect
               
        End If

        Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,cep,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nLinha & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & "" & "','"
        Sql = Sql & Left(Mask(sEnd), 80) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Mask(sCidade) & "','" & sCep & "','" & sUF & "','" & Mask(sDadosLanc) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & nParc & "," & nQtdeParc & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(sValorParc, "#0.00")) & ",'" & Mask(txtObs.Text) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
        cn.Execute Sql, rdExecDirect

        modLg "Emissão de guia nº " & nNumGuia & " Codigo: " & Val(txtCod.Text) & " - " & lblProp.Caption
    Next
End With

nLinha = 1
For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked Then
        Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nLinha & ","
        Sql = Sql & Val(Right(lvTrib.ListItems(x).Key, 3)) & ",'" & Left(lvTrib.ListItems(x).SubItems(3), 50) & "'," & Virg2Ponto(RemovePonto(lvTrib.ListItems(x).SubItems(1) * lvTrib.ListItems(x).SubItems(2))) & ")"
        cn.Execute Sql, rdExecDirect
        nLinha = nLinha + 1
    End If
Next

If chkTxExp.value = vbChecked Then
    Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nLinha & ","
    Sql = Sql & 3 & ",'" & "TAXA DE EXPEDIENTE" & "'," & Virg2Ponto(RemovePonto(Format(nValorTaxa, "#0.00"))) & ")"
    cn.Execute Sql, rdExecDirect
End If

bGerado = True

'EXIBE RELATORIO
If bBoleto Then
    frmReport.ShowReport2 "BOLETOGUIA", frmMdi.HWND, Me.HWND, nSid, nNumGuia
Else
    frmReport.ShowReport2 "BOLETOGUIA_V4", frmMdi.HWND, Me.HWND, nSid, nNumGuia
End If

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect


Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1

     
     MsgBox rdoErrors(x).Description

Next
Resume Next

End Sub

Private Sub EmiteBoletoRegistrado()
Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String, v8 As String, v9 As String, V10 As String
Dim RdoAux As rdoResultset, Sql As String, nNumdoc As Long, nSeq As Integer, nCodReduz As Long, nLanc As Integer, nParc As Integer, nAno As Integer
Dim sDataBase As String, sDataVencto As String, nSeqObs As Integer, bAbateu As Boolean, aTributos() As TRIBUTO, aTributosU() As TRIBUTO

nCodReduz = Val(txtCod.Text)
nLanc = cmbLanc.ItemData(cmbLanc.ListIndex)
nAno = Val(txtAno.Text)
sDataBase = Right$(frmMdi.Sbar.Panels(6).Text, 10)
ReDim aTributos(0)
ReDim aTributosU(0)

For x = 1 To lvTrib.ListItems.Count
    If lvTrib.ListItems(x).Checked = True Then
       ReDim Preserve aTributos(UBound(aTributos) + 1)
       ReDim Preserve aTributosU(UBound(aTributosU) + 1)
       aTributosU(UBound(aTributosU)).nCodTributo = Val(Right$(lvTrib.ListItems(x).Key, 3))
       aTributos(UBound(aTributos)).nCodTributo = Val(Right$(lvTrib.ListItems(x).Key, 3))

       If chkUnica.value = 0 Then
          If Not bAbateu Then
             aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber((CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1))) - CDbl(txtAbate.Text), 2)
          Else
             aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber(CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1)), 2)
          End If
       Else
          aTributosU(UBound(aTributosU)).nValorTributo = FormatNumber(CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1)), 2)
          If aTributosU(UBound(aTributosU)).nCodTributo = 22 Or aTributosU(UBound(aTributosU)).nCodTributo = 23 Or aTributosU(UBound(aTributosU)).nCodTributo = 24 Or aTributosU(UBound(aTributosU)).nCodTributo = 28 Then
          Else
             aTributosU(UBound(aTributosU)).nValorTributo = aTributosU(UBound(aTributosU)).nValorTributo - aTributosU(UBound(aTributosU)).nValorTributo * 0.05
          End If
       End If
       If Not bAbateu Then
            aTributos(UBound(aTributos)).nValorTributo = FormatNumber(((CDbl(lvTrib.ListItems(x).SubItems(2)) * CDbl(lvTrib.ListItems(x).SubItems(1))) - CDbl(txtAbate.Text)) / Val(txtNumParc.Text), 2)
            bAbateu = True
       Else
            aTributos(UBound(aTributos)).nValorTributo = FormatNumber((CDbl(lvTrib.ListItems(x).SubItems(2) * CDbl(lvTrib.ListItems(x).SubItems(1))) / Val(txtNumParc.Text)), 2)
       End If
    End If
Next


For nParc = 0 To Val(txtNumParc.Text)
    If nParc = 0 And chkUnica.value = 0 Then GoTo proximo
    
    If nParc > 0 Then
        If grdData.Rows = 1 Then
            sDataVencto = mskDataVencimento.Text
        Else
            sDataVencto = grdData.TextMatrix(nParc, 1)
        End If
    Else
        sDataVencto = mskDataVencimento.Text
    End If
    
    Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nNumdoc = RdoAux!maximo + 1
    RdoAux.Close

    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,ISENTOMJ,PERCISENCAO,TIPODOC,emissor,registrado) VALUES("
    Sql = Sql & nNumdoc & ",'" & Format(Now, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",3,'" & NomeDeLogin & "',1)"
    cn.Execute Sql, rdExecDirect
    
    Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux2!SEQMAXIMA) Then
       nSeq = 0
    Else
       nSeq = RdoAux2!SEQMAXIMA + 1
    End If

    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
    Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & nNumdoc & ")"
    cn.Execute Sql, rdExecDirect

    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
    Sql = Sql & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect

    If Trim(txtObs.Text) <> "" Then
        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=0"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!maximo) Then
                nSeqObs = 1
            Else
                nSeqObs = !maximo + 1
            End If
           .Close
        End With
        Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & nSeqObs & ",'" & Mask(Trim(txtObs.Text)) & "',"
        Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    End If

    If Trim(sObsIss) <> "" Then
        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=0"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!maximo) Then
                nSeqObs = 1
            Else
                nSeqObs = !maximo + 1
            End If
           .Close
        End With
        Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & nSeqObs & ",'" & Mask(Trim(sObsIss)) & "',"
        Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    End If

    sFullTrib = ""
    For Y = 1 To UBound(aTributos)
       '*************
        If bISSFixo And bTLL Then
            If nLanc = 14 Then 'se for iss fixo ignorar outros tributos que não sejam iss fixo(11)
                If aTributos(Y).nCodTributo <> 11 Then
                    GoTo ProximoTrib
                End If
            ElseIf nLanc = 6 Then 'se for taxa licenca ignorar o tributo iss fixo(11)
                If aTributos(Y).nCodTributo = 11 Then '
                    GoTo ProximoTrib
                End If
            End If
        End If
      '*************

        If aTributos(Y).nCodTributo <> 3 Then
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
            Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & ","
            Sql = Sql & aTributos(Y).nCodTributo & "," & IIf(nParc = "0", Virg2Ponto(CStr(aTributosU(Y).nValorTributo)), Virg2Ponto(CStr(aTributos(Y).nValorTributo))) & ")"
            cn.Execute Sql, rdExecDirect
         End If
ProximoTrib:
    Next


proximo:

Next

If Trim(sObsIss) <> "" Then
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodigoImovel
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqObs = 1
        Else
            nSeqObs = !maximo + 1
        End If
       .Close
    End With
        
    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
    Sql = Sql & nCodigoImovel & "," & nSeqObs & ",'" & Format(Now, "dd/mm/yyyy") & "','" & sObsIss & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
    sObsIss = ""
End If

v1 = lblProp.Caption
v2 = Left(lblRua.Caption & ", " & lblNumImovel.Caption & IIf(lblCompl.Caption <> "", " " & lblCompl.Caption, "") & " - " & lblBairro.Caption, 60)
v3 = mskDataVencimento.Text
v4 = RetornaNumero(lblNumInsc.Caption)
v5 = "287353200" & Format(nNumdoc, "00000000")
v6 = RetornaNumero(lblTotal.Caption)
v7 = Left("JABOTICABAL", 18)
v8 = "SP"
v9 = lblCEP.Caption
V10 = NomeDeLogin
If Trim(lblCEP.Caption) = "" Or Trim(lblCEP.Caption) = "-" Then
    v9 = "14870-000"
End If
ShellExecute HWND, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10, vbNullString, vbNullString, conSwNormal

Limpa

End Sub

