VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCalcGeral 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Execução do Cálculo Geral de ITU/IPTU"
   ClientHeight    =   5910
   ClientLeft      =   3825
   ClientTop       =   2550
   ClientWidth     =   10680
   Icon            =   "frmCalcGeral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5910
   ScaleWidth      =   10680
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   3870
      TabIndex        =   87
      Top             =   8100
      Width           =   1140
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   525
      TabIndex        =   9
      Top             =   3630
      Width           =   1275
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Cálculo Geral"
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   3060
      Width           =   1590
   End
   Begin VB.OptionButton opt1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Cálculo Individual:"
      Height          =   210
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   3375
      Value           =   -1  'True
      Width           =   1590
   End
   Begin VB.TextBox txtNumParc 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   1650
      TabIndex        =   3
      Top             =   930
      Width           =   675
   End
   Begin VB.TextBox txtAnoCalculo 
      Appearance      =   0  'Flat
      Height          =   280
      Left            =   1635
      MaxLength       =   4
      TabIndex        =   1
      Top             =   180
      Width           =   1035
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   2295
      Left            =   2355
      TabIndex        =   0
      Top             =   3240
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   4048
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Orientation     =   1
   End
   Begin esMaskEdit.esMaskedEdit mskDataBase 
      Height          =   285
      Left            =   1650
      TabIndex        =   2
      Top             =   525
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmCalcGeral.frx":030A
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
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   165
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   5115
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frmCalcGeral.frx":0326
      PICN            =   "frmCalcGeral.frx":0342
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
      Left            =   165
      TabIndex        =   7
      ToolTipText     =   "Cancelar Edição"
      Top             =   4755
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
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
      MICON           =   "frmCalcGeral.frx":03B0
      PICN            =   "frmCalcGeral.frx":03CC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCalculo 
      Height          =   315
      Left            =   165
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   4035
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Calcular"
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
      MICON           =   "frmCalcGeral.frx":0526
      PICN            =   "frmCalcGeral.frx":0542
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   165
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   4395
      Width           =   1350
      _ExtentX        =   2381
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCalcGeral.frx":05E1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRel 
      Height          =   315
      Left            =   150
      TabIndex        =   81
      ToolTipText     =   "Sair da Tela"
      Top             =   5490
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Relatório"
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
      MICON           =   "frmCalcGeral.frx":05FD
      PICN            =   "frmCalcGeral.frx":0619
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdLista 
      Height          =   315
      Left            =   285
      TabIndex        =   82
      ToolTipText     =   "Sair da Tela"
      Top             =   8085
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Lista"
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
      MICON           =   "frmCalcGeral.frx":0704
      PICN            =   "frmCalcGeral.frx":0720
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdArray 
      Height          =   315
      Left            =   1695
      TabIndex        =   83
      ToolTipText     =   "Sair da Tela"
      Top             =   8085
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Array"
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
      MICON           =   "frmCalcGeral.frx":080B
      PICN            =   "frmCalcGeral.frx":0827
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
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto Única %.:"
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   95
      Top             =   1980
      Width           =   1575
   End
   Begin VB.Label lblPercUnica3 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   94
      Top             =   1995
      Width           =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto Única %.:"
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   93
      Top             =   1770
      Width           =   1575
   End
   Begin VB.Label lblPercUnica2 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   92
      Top             =   1785
      Width           =   570
   End
   Begin VB.Label lblUnica3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   7080
      TabIndex        =   91
      Top             =   5580
      Width           =   1065
   End
   Begin VB.Label lblUnica2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   7080
      TabIndex        =   90
      Top             =   5340
      Width           =   1065
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Única 3.:"
      Height          =   225
      Index           =   2
      Left            =   6030
      TabIndex        =   89
      Top             =   5580
      Width           =   1065
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Única 2.:"
      Height          =   225
      Index           =   1
      Left            =   6030
      TabIndex        =   88
      Top             =   5340
      Width           =   1065
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   10320
      TabIndex        =   86
      Top             =   2850
      Width           =   135
   End
   Begin VB.Label lblPerc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   9480
      TabIndex        =   85
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentual de Isenção.......:"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   6810
      TabIndex        =   84
      Top             =   2850
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplicação de  Fatores...:"
      Height          =   225
      Index           =   15
      Left            =   2865
      TabIndex        =   80
      Top             =   3375
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Venal Territorial.........:"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   16
      Left            =   6810
      TabIndex        =   79
      Top             =   1740
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Venal Predial.............:"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   17
      Left            =   6810
      TabIndex        =   78
      Top             =   1950
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Agrupamento............:"
      Height          =   225
      Index           =   21
      Left            =   2880
      TabIndex        =   77
      Top             =   3615
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Pedologia.................:"
      Height          =   225
      Index           =   22
      Left            =   2865
      TabIndex        =   76
      Top             =   1980
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Situação...................:"
      Height          =   225
      Index           =   23
      Left            =   2880
      TabIndex        =   75
      Top             =   2205
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Profundidade............:"
      Height          =   225
      Index           =   24
      Left            =   2865
      TabIndex        =   74
      Top             =   2445
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Categoria..................:"
      Height          =   225
      Index           =   25
      Left            =   2865
      TabIndex        =   73
      Top             =   2670
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Distrito.......................:"
      Height          =   225
      Index           =   26
      Left            =   2865
      TabIndex        =   72
      Top             =   2910
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Topografia................:"
      Height          =   225
      Index           =   27
      Left            =   2865
      TabIndex        =   71
      Top             =   1740
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fator Gleba........................:"
      Height          =   225
      Index           =   28
      Left            =   2865
      TabIndex        =   70
      Top             =   3150
      Width           =   1980
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do ITU/IPTU.............:"
      ForeColor       =   &H00008000&
      Height          =   225
      Index           =   30
      Left            =   6810
      TabIndex        =   69
      Top             =   2385
      Width           =   1980
   End
   Begin VB.Label lblFatorT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   68
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label lblFatorP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   67
      Top             =   1980
      Width           =   1515
   End
   Begin VB.Label lblFatorS 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   66
      Top             =   2205
      Width           =   1515
   End
   Begin VB.Label lblFatorF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   65
      Top             =   2445
      Width           =   1515
   End
   Begin VB.Label lblFatorC 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   64
      Top             =   2670
      Width           =   1515
   End
   Begin VB.Label lblFatorD 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   63
      Top             =   2910
      Width           =   1515
   End
   Begin VB.Label lblFatorG 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   62
      Top             =   3150
      Width           =   1515
   End
   Begin VB.Label lblMulF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   61
      Top             =   3375
      Width           =   1515
   End
   Begin VB.Label lblAgrup 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   5025
      TabIndex        =   60
      Top             =   3615
      Width           =   1515
   End
   Begin VB.Label lblVVT 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   8955
      TabIndex        =   59
      Top             =   1740
      Width           =   1515
   End
   Begin VB.Label lblVVP 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   8955
      TabIndex        =   58
      Top             =   1965
      Width           =   1515
   End
   Begin VB.Label lblIPTU 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   8955
      TabIndex        =   57
      Top             =   2385
      Width           =   1515
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Venal do Imóvel........:"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   37
      Left            =   6810
      TabIndex        =   56
      Top             =   2175
      Width           =   1980
   End
   Begin VB.Label lblVVI 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   8955
      TabIndex        =   55
      Top             =   2175
      Width           =   1515
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPTU Corrigido...........:"
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   38
      Left            =   6810
      TabIndex        =   54
      Top             =   2610
      Width           =   1980
   End
   Begin VB.Label lblIPTUCorrigido 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   8955
      TabIndex        =   53
      Top             =   2610
      Width           =   1515
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      Caption         =   "Cálculo 2004"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   2850
      TabIndex        =   52
      Top             =   1470
      Width           =   1140
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do ITU/IPTU..:"
      Height          =   225
      Index           =   33
      Left            =   3015
      TabIndex        =   51
      Top             =   5115
      Width           =   1545
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Única...:"
      Height          =   225
      Index           =   34
      Left            =   6045
      TabIndex        =   50
      Top             =   5100
      Width           =   1065
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Parcela...:"
      Height          =   225
      Index           =   35
      Left            =   8340
      TabIndex        =   49
      Top             =   5085
      Width           =   1200
   End
   Begin VB.Label lblValorFinal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   4605
      TabIndex        =   48
      Top             =   5100
      Width           =   1155
   End
   Begin VB.Label lblUnica 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   7080
      TabIndex        =   47
      Top             =   5100
      Width           =   1065
   End
   Begin VB.Label lblParcela 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00C00000&
      Height          =   210
      Left            =   9510
      TabIndex        =   46
      Top             =   5100
      Width           =   885
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   44
      Left            =   2970
      TabIndex        =   45
      Top             =   4860
      Width           =   1140
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Fração Ideal.........:"
      Height          =   225
      Index           =   7
      Left            =   2850
      TabIndex        =   44
      Top             =   990
      Width           =   1470
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Área do Terreno..:"
      Height          =   225
      Index           =   8
      Left            =   5340
      TabIndex        =   43
      Top             =   765
      Width           =   1440
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Testada Principal.:"
      Height          =   225
      Index           =   9
      Left            =   8070
      TabIndex        =   42
      Top             =   765
      Width           =   1440
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Área Construida...:"
      Height          =   225
      Index           =   10
      Left            =   5340
      TabIndex        =   41
      Top             =   990
      Width           =   1440
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tem Predial..........:"
      Height          =   225
      Index           =   11
      Left            =   2850
      TabIndex        =   40
      Top             =   765
      Width           =   1440
   End
   Begin VB.Label lblPredial 
      BackStyle       =   0  'Transparent
      Caption         =   "Sim"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4260
      TabIndex        =   39
      Top             =   750
      Width           =   570
   End
   Begin VB.Label lblFracao 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4260
      TabIndex        =   38
      Top             =   1005
      Width           =   705
   End
   Begin VB.Label lblAreaTerreno 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6705
      TabIndex        =   37
      Top             =   780
      Width           =   1200
   End
   Begin VB.Label lblAreaPrincipal 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6705
      TabIndex        =   36
      Top             =   1005
      Width           =   1200
   End
   Begin VB.Label lblTestada 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   9465
      TabIndex        =   35
      Top             =   780
      Width           =   930
   End
   Begin VB.Label lblIC 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6705
      TabIndex        =   34
      Top             =   75
      Width           =   3735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insc.Cadastral.....:"
      Height          =   225
      Index           =   39
      Left            =   5340
      TabIndex        =   33
      Top             =   60
      Width           =   1350
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço.............:"
      Height          =   225
      Index           =   40
      Left            =   2850
      TabIndex        =   32
      Top             =   510
      Width           =   1440
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário...........:"
      Height          =   225
      Index           =   41
      Left            =   2850
      TabIndex        =   31
      Top             =   270
      Width           =   1440
   End
   Begin VB.Label lblProp 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4260
      TabIndex        =   30
      Top             =   270
      Width           =   6120
   End
   Begin VB.Label lblRua 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4260
      TabIndex        =   29
      Top             =   510
      Width           =   6150
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Parâmetros do Cálculo"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   45
      Left            =   2820
      TabIndex        =   28
      Top             =   30
      Width           =   2040
   End
   Begin VB.Label lblTestadaMedia 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   9465
      TabIndex        =   27
      Top             =   1005
      Width           =   930
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Testada Média.....:"
      Height          =   225
      Index           =   46
      Left            =   8070
      TabIndex        =   26
      Top             =   990
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de Cálculo.....:"
      Height          =   225
      Left            =   120
      TabIndex        =   25
      Top             =   255
      Width           =   1455
   End
   Begin VB.Label lblEstimado 
      BackStyle       =   0  'Transparent
      Caption         =   "14350"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   24
      Top             =   2220
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Imóveis Estimados.:"
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   23
      Top             =   2220
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Parcelas......:"
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   22
      Top             =   975
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Base.............:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   90
      TabIndex        =   21
      Top             =   570
      Width           =   1455
   End
   Begin VB.Label lblTemUnica 
      BackStyle       =   0  'Transparent
      Caption         =   "Sim"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   20
      Top             =   1260
      Width           =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela Única........:"
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   19
      Top             =   1260
      Width           =   1455
   End
   Begin VB.Label lblPercUnica 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   18
      Top             =   1560
      Width           =   570
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Desconto Única %.:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   17
      Top             =   1545
      Width           =   1575
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2250
      TabIndex        =   16
      Top             =   2940
      Width           =   480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2355
      TabIndex        =   15
      Top             =   5580
      Width           =   270
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota Predial......:"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   2490
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "1,5 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   13
      Top             =   2475
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota Territorial..:"
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "3 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1680
      TabIndex        =   11
      Top             =   2760
      Width           =   810
   End
End
Attribute VB_Name = "frmCalcGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Lista
    nCodReduzido As Long
    nCodCidadao As Long
End Type

Dim bExec As Boolean
Dim aParc() As Date
Dim xImovel As clsImovel

'PARAMETROS
Dim nUfirCalc As Double
Dim nUfir1999 As Double
Dim nAliquotaPredial As Double
Dim nAliquotaTerritorial As Double
Dim bTemPredial As Boolean
Dim bFracaoIdeal As Boolean
Dim nAreaTerreno As Double
Dim nAreaPrincipal As Double
Dim nCodAgrupamento As Integer
Dim nValorAgrupamento As Double
Dim nNumTestadas As Integer
Dim nTestadaPrincipal As Double
Dim nCodGleba As Integer
Dim nFatorGleba As Double
Dim nCodProfundidade As Integer
Dim nValorProfundidade As Double
Dim nFatorProfundidade As Double
Dim nCodSituacao As Integer
Dim nFatorSituacao As Double
Dim nCodPedologia As Integer
Dim nFatorPedologia As Double
Dim nCodTopografia As Integer
Dim nFatorTopografia As Double
Dim nFatorDistrito As Double
Dim nValorFatores As Double
Dim nFatorCategoria As Double
Dim nValorVenalTerritorial As Double
Dim nValorVenalPredial As Double
Dim nCodTributo As Integer
Dim nValorVenalImovel As Double
Dim nValorVenalImovel98 As Double
Dim nValorIptu As Double, nValorITU As Double
Dim nValorFinal As Double, nNumParc As Integer
'GERAL
Dim nCodReduz As Long
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim Sql As String
Dim nAnoCalculo As Integer
'TIPOS
Private Type PROFUNDIDADE
    Distrito As Integer
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORPROFUN
    Distrito As Integer
    Codigo As Integer
    Fator As Double
End Type
Private Type GLEBA
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORCATEG
    Uso As Integer
    Tipo As Integer
    Categoria As Integer
    Fator As Double
End Type

Private Type LaserIPTU
    nCodReduz As Long
    nVVT1 As Double
    nVVC1 As Double
    nVVI1 As Double
    nVVT2 As Double
    nVVC2 As Double
    nVVI2 As Double
    nImpPre1 As Double
    nImpTer1 As Double
    nImpPre2 As Double
    nImpTer2 As Double
    sNatureza As String
    nAreaConstruida As Double
    nTestadaPrincipal As Double
    nValorParcela1 As Double
    nValorUnica1 As Double
    nValorParcela2 As Double
    nValorUnica2 As Double
    nQtdeParc As Integer
    nAreaTerreno As Double
    nFatorCat As Double
    nFatorPed As Double
    nFatorSit As Double
    nFatorPro As Double
    nFatorTop As Double
    nFatorDis As Double
    nfatorGle As Double
    nAgrupamento As Double
    nAliquota As Double
    nFracaoIdeal As Double
    nDistrito As Integer
    nSetor As Integer
    nQuadra As Integer
    nLote As Integer
    nFace As Integer
    nUnidade As Integer
    nSubUnidade As Integer
    sProprietario As String
    nCodLogradouro As Integer
    sEndereco As String
    nNumero As Integer
    sComplemento As String
    sBairro As String
    sEndEntrega As String
    sComplEntrega As String
    sBairroEntrega As String
    sCidadeEntrega As String
    sCepEntrega As String
    sUFEntrega As String
End Type

'MATRIZES
Dim aFatorD() As Double
Dim aFatorD98() As Double
Dim aFatorP() As Double
Dim aFatorP98() As Double
Dim aFatorT() As Double
Dim aFatorT98() As Double
Dim aFatorS() As Double
Dim aFatorS98() As Double
Dim aFatorG() As Double
Dim aFatorG98() As Double
Dim aFatorR() As Double
Dim aFatorR98() As Double
Dim aProf() As PROFUNDIDADE
Dim aFatorF() As FATORPROFUN
Dim aFatorF98() As FATORPROFUN
Dim aFatorC() As FATORCATEG
Dim aFatorC98() As FATORCATEG
Dim aGleba() As GLEBA
Dim aLaser() As LaserIPTU

Private Sub cmdArray_Click()
Dim x As Long, RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, nAno1 As Integer, nAno2 As Integer
Dim nCodReduz As Long, nCodReduz2 As Long, nVVT As Double, nVVC As Double, nVVI As Double, nImpostoPredial As Double, nImpostoTerritorial As Double, sNatureza As String, nAreaPredial As Double
Dim nTestada As Double, nValorParc As Double, nValorUnica As Double, nQtdeParc As Double, nAreaTerreno As Double, nFatorCat As Double, nFatorPed As Double, nFatorSit As Double
Dim nFatorPro As Double, nFatorTop As Double, nFatorDis As Double, nfatorGle As Double, nAgrupamento As Double, nFracao As Double, nAliquota As Double
Dim nVVT2 As Double, nVVC2 As Double, nVVI2 As Double, nImpostoPredial2 As Double, nImpostoTerritorial2 As Double, sNatureza2 As String, nAreaPredial2 As Double
Dim nTestada2 As Double, nValorParc2 As Double, nValorUnica2 As Double, nQtdeparc2 As Double, nAreaTerreno2 As Double, nFatorCat2 As Double, nFatorPed2 As Double, nFatorSit2 As Double
Dim nFatorPro2 As Double, nFatorTop2 As Double, nFatorDis2 As Double, nfatorGle2 As Double, nAgrupamento2 As Double, nFracao2 As Double, nAliquota2 As Double

nAno1 = 2010: nAno2 = 2011
Sql = "TRUNCATE TABLE RELIPTU"
cn.Execute Sql, rdExecDirect

For x = 1 To 40000
    Sql = "SELECT * FROM LASERIPTU WHERE ANO=" & nAno1 & " AND CODREDUZIDO=" & x & " ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    With RdoAux
        If .RowCount > 0 Then
            nCodReduz = !CODREDUZIDO
            nVVT = FormatNumber(!vvt, 2)
            nVVC = FormatNumber(!vvc, 2)
            nVVI = FormatNumber(!VVI, 2)
            nImpostoPredial = FormatNumber(!impostopredial, 2)
            nImpostoTerritorial = FormatNumber(!IMPOSTOTERRITORIAL, 2)
            sNatureza = !Natureza
            nAreaPredial = FormatNumber(!areaconstrucao, 2)
            nTestada = FormatNumber(!TESTADAPRINC, 2)
            nValorParc = FormatNumber(!valortotalparc, 2)
            nValorUnica = FormatNumber(!VALORTOTALUNICA, 2)
            nAreaTerreno = FormatNumber(!AreaTerreno, 2)
            nFatorCat = FormatNumber(!FATORCAT, 2)
            nFatorPed = FormatNumber(!FATORPED, 2)
            nFatorSit = FormatNumber(!FATORSIT, 2)
            nFatorPro = FormatNumber(!FATORPRO, 2)
            nFatorTop = FormatNumber(!FATORTOP, 2)
            nFatorDis = FormatNumber(!FATORDIS, 2)
            nfatorGle = FormatNumber(!FATORGLE, 2)
            nAgrupamento = FormatNumber(!Agrupamento, 2)
            nFracao = FormatNumber(!FracaoIdeal, 2)
            nAliquota = FormatNumber(!Aliquota, 2)
        Else
            nCodReduz = 0
            nVVT = 0
            nVVC = 0
            nVVI = 0
            nImpostoPredial = 0
            nImpostoTerritorial = 0
            sNatureza = ""
            nAreaPredial = 0
            nTestada = 0
            nValorParc = 0
            nValorUnica = 0
            nAreaTerreno = 0
            nFatorCat = 0
            nFatorPed = 0
            nFatorSit = 0
            nFatorPro = 0
            nFatorTop = 0
            nFatorDis = 0
            nfatorGle = 0
            nAgrupamento = 0
            nFracao = 0
            nAliquota = 0
        End If
       .Close
    End With
    DoEvents
    Sql = "SELECT * FROM LASERIPTU WHERE ANO=" & nAno2 & " AND CODREDUZIDO=" & x & " ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    With RdoAux
        If .RowCount > 0 Then
            nCodReduz2 = !CODREDUZIDO
            nVVT2 = FormatNumber(!vvt, 2)
            nVVC2 = FormatNumber(!vvc, 2)
            nVVI2 = FormatNumber(!VVI, 2)
            nImpostoPredial2 = FormatNumber(!impostopredial, 2)
            nImpostoTerritorial2 = FormatNumber(!IMPOSTOTERRITORIAL, 2)
            sNatureza2 = !Natureza
            nAreaPredial2 = FormatNumber(!areaconstrucao, 2)
            nTestada2 = FormatNumber(!TESTADAPRINC, 2)
            nValorParc2 = FormatNumber(!valortotalparc, 2)
            nValorUnica2 = FormatNumber(!VALORTOTALUNICA, 2)
            nAreaTerreno2 = FormatNumber(!AreaTerreno, 2)
            nFatorCat2 = FormatNumber(!FATORCAT, 2)
            nFatorPed2 = FormatNumber(!FATORPED, 2)
            nFatorSit2 = FormatNumber(!FATORSIT, 2)
            nFatorPro2 = FormatNumber(!FATORPRO, 2)
            nFatorTop2 = FormatNumber(!FATORTOP, 2)
            nFatorDis2 = FormatNumber(!FATORDIS, 2)
            nfatorGle2 = FormatNumber(!FATORGLE, 2)
            nAgrupamento = FormatNumber(!Agrupamento, 2)
            nFracao = FormatNumber(!FracaoIdeal, 2)
            nAliquota = FormatNumber(!Aliquota, 2)
        Else
            nCodReduz2 = 0
            nVVT2 = 0
            nVVC2 = 0
            nVVI2 = 0
            nImpostoPredial2 = 0
            nImpostoTerritorial2 = 0
            sNatureza2 = ""
            nAreaPredial2 = 0
            nTestada2 = 0
            nValorParc2 = 0
            nValorUnica2 = 0
            nAreaTerreno2 = 0
            nFatorCat = 0
            nFatorPed = 0
            nFatorSit = 0
            nFatorPro = 0
            nFatorTop = 0
            nFatorDis = 0
            nfatorGle = 0
            nAgrupamento2 = 0
            nFracao2 = 0
            nAliquota2 = 0
        End If
       .Close
    End With
    If nCodReduz > 0 Or nCodReduz2 > 0 Then
        Sql = "INSERT RELIPTU (CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,AREATERRENO,FATORCAT,FATORPED,FATORSIT,FATORPRO,FATORTOP,FATORDIS,FATORGLE,AGRUPAMENTO,FRACAOIDEAL,ALIQUOTA,VVT2,VVC2,VVI2,IMPOSTOPREDIAL2,"
        Sql = Sql & "IMPOSTOTERRITORIAL2,NATUREZA2,AREACONSTRUCAO2,TESTADAPRINC2,VALORTOTALPARC2,VALORTOTALUNICA2,AREATERRENO2,FATORCAT2,FATORPED2,FATORSIT2,FATORPRO2,FATORTOP2,FATORDIS2,FATORGLE2,AGRUPAMENTO2,FRACAOIDEAL2,ALIQUOTA2) VALUES(" & x & "," & Virg2Ponto(CStr(nVVT)) & "," & Virg2Ponto(CStr(nVVC)) & ","
        Sql = Sql & Virg2Ponto(CStr(nVVI)) & "," & Virg2Ponto(CStr(nImpostoPredial)) & "," & Virg2Ponto(CStr(nImpostoTerritorial)) & ",'" & sNatureza & "'," & Virg2Ponto(CStr(nAreaPredial)) & "," & Virg2Ponto(CStr(nTestada)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorParc)) & "," & Virg2Ponto(CStr(nValorUnica)) & "," & Virg2Ponto(CStr(nAreaTerreno)) & "," & Virg2Ponto(CStr(nFatorCat)) & "," & Virg2Ponto(CStr(nFatorPed)) & "," & Virg2Ponto(CStr(nFatorSit)) & "," & Virg2Ponto(CStr(nFatorPro)) & ","
        Sql = Sql & Virg2Ponto(CStr(nFatorTop)) & "," & Virg2Ponto(CStr(nFatorDis)) & "," & Virg2Ponto(CStr(nfatorGle)) & "," & Virg2Ponto(CStr(nAgrupamento)) & "," & Virg2Ponto(CStr(nFracao)) & "," & Virg2Ponto(CStr(nAliquota)) & ","
        Sql = Sql & Virg2Ponto(CStr(nVVT2)) & "," & Virg2Ponto(CStr(nVVC2)) & "," & Virg2Ponto(CStr(nVVI2)) & "," & Virg2Ponto(CStr(nImpostoPredial2)) & "," & Virg2Ponto(CStr(nImpostoTerritorial2)) & ",'" & sNatureza2 & "'," & Virg2Ponto(CStr(nAreaPredial2)) & "," & Virg2Ponto(CStr(nTestada2)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorParc2)) & "," & Virg2Ponto(CStr(nValorUnica2)) & "," & Virg2Ponto(CStr(nAreaTerreno2)) & "," & Virg2Ponto(CStr(nFatorCat2)) & "," & Virg2Ponto(CStr(nFatorPed2)) & "," & Virg2Ponto(CStr(nFatorSit2)) & "," & Virg2Ponto(CStr(nFatorPro2)) & ","
        Sql = Sql & Virg2Ponto(CStr(nFatorTop2)) & "," & Virg2Ponto(CStr(nFatorDis2)) & "," & Virg2Ponto(CStr(nfatorGle2)) & "," & Virg2Ponto(CStr(nAgrupamento2)) & "," & Virg2Ponto(CStr(nFracao2)) & "," & Virg2Ponto(CStr(nAliquota2)) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

MsgBox "FIM"
End Sub

Private Sub cmdCalculo_Click()

Limpa

CarregaImovel

If Val(txtNumParc.Text) = 0 Then
    MsgBox "Digite a qtde de parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

'If opt1(0).Value = True Then
'    MsgBox "Você não tem permissão para realizar o cálculo geral de IPTU.", vbExclamation, "ALERTA DE SEGURANÇA !!"
    
'    Exit Sub
'End If
nAnoCalculo = Val(txtAnoCalculo.Text)
LoadMatrix
'CARREGA PARAMETROS
'nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3
bExec = True
If opt1(1).value = True Then
    If Val(txtCod.Text) = 0 Then
       MsgBox "Digite o código do imóvel.", vbExclamation, "Atenção"
    Else
       Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCod.Text)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
            If .RowCount = 0 Then
                MsgBox "Imóvel não cadastrado.", vbExclamation, "Atenção"
            Else
'                CalculoIndividual (Val(txtCod.Text))
                CalculoInd
            End If
       End With
    End If
Else
'    If frmMdi.frTeste.Visible = False Then
'        MsgBox "Calculo geral apenas para base de testes."
'        Exit Sub
'    End If
'    If Not IsDate(mskDataBase.Text) Then
'        MsgBox "Data Base Inválida.", vbExclamation, "atenção"
'        Exit Sub
'    End If
    
'    If NomeDeLogin <> "SCHWARTZ" Then
'        MsgBox "Cálculo geral bloqueado até o término do cálculo 2012.", vbCritical, "ERRO"
'        Exit Sub
'    End If
    
'    If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "FACTORE" And NomeDeLogin <> "MARIELA" And NomeDeLogin <> "REGINA" Then Exit Sub
    
'    If MsgBox("Executar o cálculo de IPTU para todos os imóveis?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
'    Ocupado
'     CalculoGeral
'    CalculoGeralEstimado
'    Liberado
 '   If bExec Then
  '     MsgBox "Calculo efetuado", vbExclamation, "atenção"
   ' End If
End If

End Sub

Private Sub CalculoIndividual(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double, RdoAux4 As rdoResultset, RdoAux5 As rdoResultset, RdoAux6 As rdoResultset
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer, bCalcProc As Boolean
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

bCalcProc = False
bIsento = False
lblPerc.Caption = "0"

Sql = "SELECT * FROM CALCPROC WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        bCalcProc = True
    End If
   .Close
End With

If bCalcProc Then GoTo FASE1

Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO,PERCISENCAO "
Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Val(txtAnoCalculo.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If !percisencao > 0 And !percisencao < 100 Then
            lblPerc.Caption = !percisencao
        Else
            MsgBox "Este imóvel esta classificado como: " & !DESCTIPO, vbExclamation, "Atenção"
            bIsento = True
        End If
    End If
   .Close
End With

Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND CODISENCAO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Este imóvel esta classificado como: " & !DESCTIPO, vbExclamation, "Atenção"
'        Exit Sub
        bIsento = True
    End If
   .Close
End With

FASE1:

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where (CADIMOB.CODREDUZIDO = " & nCodReduz & ") GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
    lblIC.Caption = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00")
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    lblFracao.Caption = FormatNumber(!Dt_FracaoIdeal, 2)
    lblAreaTerreno.Caption = FormatNumber(nAreaTerreno, 2)
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    lblAreaPrincipal.Caption = FormatNumber(nAreaPrincipal, 2)
    lblPredial.Caption = IIf(bTemPredial, "Sim", "Não")
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
    lblTestada.Caption = FormatNumber(nTestada1, 2)
    lblTestadaMedia.Caption = FormatNumber(nTestadaPrincipal, 2)
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    
    'BUSCA ÁREA PRINCIPAL
    'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 1 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) And RdoAux2!QTDEPAV < 2 And nAreaTerreno < 600 Then
                        If nAnoCalculo > 2006 Then
                            If bCalcProc = False Then
                                Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                                Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                                If RdoAux4.RowCount = 0 Then
                                    bIsento = True
                                    MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                Else
                                    If ImovelAreaUnica(RdoAux4!CODPROPRIETARIO) Then
                                        MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                        bIsento = True
                                    End If
                                End If
                                RdoAux4.Close
                            End If
                        Else
                            If bCalcProc = False Then
                                bIsento = True
                                MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                Limpa
                            End If
                        End If
                    End If
                End If
               .Close
            End With
        Else
            bIsento = False
            
'            Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
'            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
 '           If RdoAux4.RowCount > 0 Then
 '               If ImovelAreaUnica(RdoAux4!CODPROPRIETARIO) Then
 '                   MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
 '                   bIsento = True
 '               End If
 '           End If
 '           RdoAux4.Close
        End If

        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
        
        'novo VVP ***********************************
        If nAnoCalculo > 2007 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            If bTemPredial Then
                Do Until .EOF
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                    nFatorCategoria = 0
                    For x = 1 To UBound(aFatorC)
                        If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                           nFatorCategoria = aFatorC(x).Fator
                           Exit For
                        End If
                    Next
                    nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
                   .MoveNext
                Loop
            End If
        Else
            If bTemPredial Then
                 nUso = !USOCONSTR
                 nTipo = !TIPOCONSTR
                 nCat = !CATCONSTR
            End If
        End If
       .Close
    End With
    
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    lblAgrup.Caption = FormatNumber(nValorAgrupamento, 2)
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    For x = 1 To UBound(aGleba)
        If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
             Exit For
        ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
             Exit For
        End If
    Next
    nCodGleba = aGleba(x).Codigo
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
    nFatorGleba = aFatorG(nCodGleba)
    lblFatorG.Caption = FormatNumber(nFatorGleba, 2)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(x).Min, 2) And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
        lblFatorF.Caption = FormatNumber(nFatorProfundidade, 2)
     Else
        nFatorProfundidade = 1
        lblFatorF.Caption = "1,00"
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    lblFatorS.Caption = FormatNumber(nFatorSituacao, 2)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    lblFatorP.Caption = FormatNumber(nFatorPedologia, 2)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    lblFatorT.Caption = FormatNumber(nFatorTopografia, 2)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba, 2)
    lblMulF.Caption = FormatNumber(nValorFatores, 2)
    'CÁLCULO VALOR VENAL TERRITORIAL
    nFatorDistrito = aFatorD(!Distrito)
    nValorFatores = nValorFatores * nFatorDistrito
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
'        nFatorDistrito = aFatorD(!Distrito)
'        nValorFatores = nValorFatores * nFatorDistrito
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2)
        lblMulF.Caption = FormatNumber(nValorFatores, 2)
        lblFatorD.Caption = FormatNumber(nFatorDistrito, 2)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        If nAnoCalculo < 2008 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
        End If
        lblFatorC.Caption = FormatNumber(nFatorCategoria, 2)
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        lblVVP.Caption = FormatNumber(nValorVenalPredial, 2)
    Else
        nFatorDistrito = 0
        nFatorCategoria = 0
        lblFatorD.Caption = FormatNumber(nFatorDistrito, 2)
        lblFatorC.Caption = FormatNumber(nFatorCategoria, 2)
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)  'reajuste 2004-2005 (TIRADO)
        lblIPTU.Caption = FormatNumber(nValorVenalImovel * (nAliquotaPredial / 100), 2)
        lblIPTUCorrigido.Caption = FormatNumber(nValorIptu, 2)
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)  'reajuste 2004-2005 (TIRADO)
        lblIPTU.Caption = FormatNumber(nValorVenalImovel * (nAliquotaTerritorial / 100), 2)
        lblIPTUCorrigido.Caption = FormatNumber(nValorITU, 2)
    End If
    lblVVI.Caption = FormatNumber(nValorVenalImovel, 2)
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    If bTemPredial Then
       nValorFinal = nValorIptu
    Else
       nValorFinal = nValorITU
    End If
    
    'PERCENTUAL ISENÇÃO
    If Val(lblPerc.Caption) > 0 Then
        nValorFinal = nValorFinal - (nValorFinal * Val(lblPerc.Caption) / 100)
    End If
    
    
    If bIsento Then
        lblValorFinal.Caption = FormatNumber(0, 2)
        lblUnica.Caption = FormatNumber(0, 2)
        lblParcela.Caption = FormatNumber(0, 2)
    Else
        lblValorFinal.Caption = FormatNumber(nValorFinal, 2)
        lblUnica.Caption = FormatNumber(nValorFinal - (nValorFinal * CDbl(lblPercUnica.Caption) / 100), 2)
        lblParcela.Caption = FormatNumber(nValorFinal / CDbl(txtNumParc.Text), 2)
    End If
End With

End Sub

Private Sub CalculoGeral()

Dim xId As Long, nNumRec As Long
Dim nValorExpDocParc As Double, nValorExpDocUnica As Double, nLastDoc As Long, nAreaTerrenoReal As Double
Dim ax As String, sDataBase As String, nAliquota As Double, RdoAux4 As rdoResultset
Dim nValorUnica As Double, nValorUnica2 As Double, nValorUnica3 As Double, nValorParcela As Double, nTestada1 As Double, nFracaoIdeal As Double
'Relatorio
Dim nValorTotalIptu As Double, nNumImovelCalc As Integer, nNumImovelOK As Integer, nNumImovelBloqueio As Integer

nValorTotalIptu = 0: nNumImovelBloqueio = 0: nNumImovelCalc = 0: nNumImovelOK = 0
nAnoCalculo = Val(txtAnoCalculo.Text)
cn.QueryTimeout = 0
cmdCalculo.Enabled = False
nAnoCalculo = 2013
If NomeDeLogin = "SCHWARTZ" Then
    Sql = "DELETE FROM LASERIPTU WHERE ANO=" & nAnoCalculo
    cn.Execute Sql, rdExecDirect
End If

nNumParc = Val(txtNumParc.Text)

TESTE:
If cGetInputState() <> 0 Then DoEvents

sDataBase = mskDataBase.Text

Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTO.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTO.TXT" For Output As #3
Open sPathBin & "\NUMDOCUMENTO.TXT" For Output As #4

'********************************
' TAXA DE EXPEDIÇÃO DE DOCUMENTO
'********************************
Calculo:
Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
        nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
        nValorExpDocUnica = FormatNumber(!ValorUnica, 2)
     Else
        MsgBox "Taxa de Expediente não cadastrada.", vbCritical, "Atenção"
        GoTo FIM
     End If
    .Close
End With
'ULTIMO Nº DE DOCUMENTO
'Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nLastDoc = !ULTIMO + 3000
'   .Close
'End With

nLastDoc = 12854944

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,CADIMOB.INATIVO,LI_CODBAIRRO,PAVIMENTO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,"
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE  AND INATIVO=0  GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.INATIVO,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
Sql = Sql & "HAVING      (cadimob.codreduzido = 12237) "
Sql = Sql & " ORDER BY CADIMOB.CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    
    nNumRec = .RowCount
    Do Until .EOF
'        If !CODREDUZIDO <> 5397 Then GoTo proximo
        'GAUGE
        If xId Mod 100 = 0 Then
           CallPb xId, nNumRec
        End If
        If !Inativo = True Then GoTo proximo
        If Not bExec Then
           MsgBox "Cálculo Interrompido pelo usuário", vbCritical, "Atenção"
           Exit Do
        End If
        'DADOS DO IMOVEL
        nCodReduz = !CODREDUZIDO
        
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & nAnoCalculo
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                GoTo proximo
            End If
           .Close
        End With
                
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND CODISENCAO=1"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                GoTo proximo
            End If
           .Close
        End With
FASE1:
        nCodBairro = !Li_CodBairro
        nAreaTerreno = !Dt_AreaTerreno
        nAreaTerrenoReal = nAreaTerreno
        nCodSituacao = !Dt_CodSituacao
        nCodPedologia = !Dt_CodPedol
        nCodTopografia = !Dt_CodTopog
        nCodAgrupamento = !CODAGRUPA
                
        nTestadaPrincipal = 0
        nFracaoIdeal = !Dt_FracaoIdeal
        bFracaoIdeal = IIf(nFracaoIdeal > 0, True, False)
        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
        'TESTADAS
        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumTestadas = .RowCount
            If nNumTestadas = 0 Then
                nTestadaPrincipal = 1
                nTestada1 = 1
            Else
                If nNumTestadas = 1 Then
                    nTestadaPrincipal = !AREATESTADA
                    nTestada1 = !AREATESTADA
                Else
                    nSomaTestada = 0
                    Do Until .EOF
                       If !NUMFACE = RdoAux!Seq Then
                          nTestada1 = !AREATESTADA
                       End If
                       nSomaTestada = nSomaTestada + !AREATESTADA
                      .MoveNext
                    Loop
                    If nNumTestadas > 0 Then
                       nTestadaPrincipal = nSomaTestada / nNumTestadas
                    Else
                       nTestadaPrincipal = 1
                    End If
                End If
            End If
           .Close
        End With
        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
        
        'BUSCA ÁREA PRINCIPAL
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
        'TEM ÁREA?
            If .RowCount > 0 Then
                If Not IsNull(RdoAux!SOMAAREA) Then
                    If RdoAux!SOMAAREA <= 65 And !USOCONSTR = 1 And (!CATCONSTR = 4 Or !CATCONSTR = 7) And !QTDEPAV < 2 And nAreaTerreno < 600 Then
                        Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                        If RdoAux4.RowCount = 0 Then
                            GoTo proximo
                        End If
                    End If
                    bTemPredial = True
                    nAreaPrincipal = FormatNumber(RdoAux!SOMAAREA, 2)
                Else
                    bTemPredial = False
                    nAreaPrincipal = 0
                End If
                If bFracaoIdeal Then
                    If nAreaPrincipal > 0 Then
                       nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
                    Else
                       nTestadaPrincipal = 1
                    End If
                End If
                If bTemPredial Then
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                End If
            Else
                bTemPredial = False
                nAreaPrincipal = 0
            End If
            
            'novo VVP ***********************************
            If nAnoCalculo > 2007 Then
                nValorVenalPredial = 0
                nFatorCategoria = 0
                If bTemPredial Then
                    Do Until .EOF
                        nUso = !USOCONSTR
                        nTipo = !TIPOCONSTR
                        nCat = !CATCONSTR
                        For x = 1 To UBound(aFatorC)
                            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                               nFatorCategoria = aFatorC(x).Fator
                               Exit For
                            End If
                        Next
                        nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
                       .MoveNext
                    Loop
                End If
            Else
                If bTemPredial Then
                     nUso = !USOCONSTR
                     nTipo = !TIPOCONSTR
                     nCat = !CATCONSTR
                End If
            End If
           
           .Close
        End With
        
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
        End If
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For x = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
               nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
            Else
               nValorProfundidade = 1
            End If
            'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
            For x = 1 To UBound(aProf)
                If aProf(x).Distrito = !Distrito Then
                   If nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And nValorProfundidade <= FormatNumber(aProf(x).Max, 2) Then
                      Exit For
                   ElseIf nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And FormatNumber(aProf(x).Max, 2) = 0 Then
                      Exit For
                   End If
                End If
            Next
            On Error Resume Next
            nCodProfundidade = aProf(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
            nFatorProfundidade = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade = aFatorF(x).Fator
                   Exit For
                End If
            Next
        Else
            nFatorProfundidade = 1
        End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nFatorDistrito = aFatorD(!Distrito)
        nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito, 2)
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
'            nFatorDistrito = aFatorD(!Distrito)
'            nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito, 2)
            'CÁLCULO VALOR VENAL TERRITORIAL
            nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
            'FATOR DISTRITO 98
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            If nAnoCalculo < 2008 Then
                nValorVenalPredial = 0
                nFatorCategoria = 0
                For x = 1 To UBound(aFatorC)
                    If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                       nFatorCategoria = aFatorC(x).Fator
                       Exit For
                    End If
                Next
                nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
            End If
           'FATOR CATEGORIA 98
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        Else
            nValorVenalPredial = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100) '1.125
            nValorFinal = nValorIptu
            nValorITU = 0
            nAliquota = nAliquotaPredial
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
            nValorFinal = nValorITU
            nValorIptu = 0
            nAliquota = nAliquotaTerritorial
        End If
        
        '** NOVA ROTINA DE QTDE DE PARCELAS ***
        If nValorFinal > 0 And nValorFinal <= 10 Then
            nNumParc = 1
        ElseIf nValorFinal > 10 And nValorFinal <= 20 Then nNumParc = 1
        ElseIf nValorFinal > 20 And nValorFinal <= 30 Then nNumParc = 2
        ElseIf nValorFinal > 30 And nValorFinal <= 40 Then nNumParc = 3
        ElseIf nValorFinal > 40 And nValorFinal <= 50 Then nNumParc = 4
        ElseIf nValorFinal > 50 And nValorFinal <= 60 Then nNumParc = 5
        ElseIf nValorFinal > 60 And nValorFinal <= 70 Then nNumParc = 6
        ElseIf nValorFinal > 70 And nValorFinal <= 80 Then nNumParc = 7
        ElseIf nValorFinal > 80 And nValorFinal <= 90 Then nNumParc = 8
        ElseIf nValorFinal > 90 And nValorFinal <= 100 Then nNumParc = 9
        Else
            nNumParc = Val(txtNumParc.Text)
        End If
        '**************************************
        
        nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
        nValorUnica2 = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica2.Caption) / 100)), 2)
        nValorUnica3 = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica3.Caption) / 100)), 2)
        nValorParcela = Round(nValorFinal / nNumParc, 2)
        
        'GRAVA TABELA LASERIPTU
        If NomeDeLogin = "SCHWARTZ" Then
            Sql = "INSERT LASERIPTU (ANO,CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,"
            Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA,AREATERRENO,FATORCAT,FATORPED,FATORSIT,"
            Sql = Sql & "FATORPRO,FATORTOP,FATORDIS,FATORGLE,AGRUPAMENTO,FRACAOIDEAL,ALIQUOTA) VALUES("
            Sql = Sql & nAnoCalculo & "," & nCodReduz & "," & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
            Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nValorIptu)) & "," & Virg2Ponto(CStr(nValorITU)) & ",'"
            Sql = Sql & IIf(bTemPredial, "Predial", "Territorial") & "'," & Virg2Ponto(CStr(nAreaPrincipal)) & "," & Virg2Ponto(CStr(nTestada1)) & ","
            Sql = Sql & Virg2Ponto(CStr(nValorParcela)) & "," & Virg2Ponto(CStr(nValorUnica)) & "," & nNumParc & ","
            Sql = Sql & Virg2Ponto(CStr(nValorExpDocParc) * Val(txtNumParc.Text)) & "," & Virg2Ponto(CStr(nValorExpDocUnica)) & "," & Virg2Ponto(CStr(nAreaTerreno)) & ","
            Sql = Sql & Virg2Ponto(CStr(nFatorCategoria)) & "," & Virg2Ponto(CStr(nFatorPedologia)) & "," & Virg2Ponto(CStr(nFatorSituacao)) & "," & Virg2Ponto(CStr(nFatorProfundidade)) & ","
            Sql = Sql & Virg2Ponto(CStr(nFatorTopografia)) & "," & Virg2Ponto(CStr(nFatorDistrito)) & "," & Virg2Ponto(CStr(nFatorGleba)) & "," & Virg2Ponto(CStr(nValorAgrupamento)) & ","
            Sql = Sql & Virg2Ponto(CStr(nFracaoIdeal)) & "," & Virg2Ponto(CStr(nAliquota)) & ")"
            cn.Execute Sql, rdExecDirect
        
        End If
        
        nValorTotalIptu = nValorTotalIptu + nValorFinal 'relatorio
        nNumImovelCalc = nNumImovelCalc + 1 'relatorio
        
        For x = 0 To nNumParc
            DoEvents
            If x = 0 And lblUnica.Caption = "Não" Then GoTo proximo
            'GRAVA NA TABELA DEBITOPARCELA
            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
            ax = ax & 3 & "," & Format(aParc(x), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
            ax = ax & 1 & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
            ax = ax & Null & "," & 0
            Print #1, ax
            'GRAVA NA TABELA DEBITO TRIBUTO
            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
            ax = ax & nCodTributo & "," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
            ax = ax & 0 & "," & 0 & "," & 0
            Print #2, ax
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & 3 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
            'GRAVA NA TABELA NUMDOCUMENTO
            nLastDoc = nLastDoc + 1
            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2)))
            Print #4, ax
            'GRAVA NA TABELA PARCELADOCUMENTO
            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & ","
            ax = ax & x & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
            Print #3, ax
        Next
proximo:
        xId = xId + 1
       .MoveNext
    Loop
End With

FIM:
Close #4
Close #3
Close #2
Close #1

MsgBox "Valor total do IPTU: " & FormatNumber(nValorTotalIptu, 2)
MsgBox "No de Imóveis Calculados: " & nNumImovelCalc
MsgBox "No de Imóveis OK: " & nNumImovelOK
MsgBox "No de Imóveis Bloqueados: " & nNumImovelBloqueio
cmdCalculo.Enabled = True
End Sub

Private Sub CalculoGeralIsentos()

Dim xId As Long, nNumRec As Long
Dim nValorExpDocParc As Double, nValorExpDocUnica As Double, nLastDoc As Long, nAreaTerrenoReal As Double
Dim ax As String, sDataBase As String, nAliquota As Double, RdoAux4 As rdoResultset
Dim nValorUnica As Double, nValorParcela As Double, nTestada1 As Double, nFracaoIdeal As Double
Dim bIsento As Boolean, nTipoIsento As Integer
'Relatorio
Dim nValorTotalIptu As Double, nNumImovelCalc As Integer, nNumImovelOK As Integer, nNumImovelBloqueio As Integer

nValorTotalIptu = 0: nNumImovelBloqueio = 0: nNumImovelCalc = 0: nNumImovelOK = 0
cn.QueryTimeout = 0
cmdCalculo.Enabled = False
nAnoCalculo = 2017

nNumParc = 12

TESTE:
If cGetInputState() <> 0 Then DoEvents

sDataBase = mskDataBase.Text

Sql = "TRUNCATE TABLE ISENTOIPTUREL"
cn.Execute Sql, rdExecDirect
'********************************
' TAXA DE EXPEDIÇÃO DE DOCUMENTO
'********************************
Calculo:
'Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'     If .RowCount > 0 Then
        nValorExpDocParc = 0
        nValorExpDocUnica = 0
'     Else
'        MsgBox "Taxa de Expediente não cadastrada.", vbCritical, "Atenção"
'        GoTo fim
'     End If
'    .Close
'End With

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,CADIMOB.INATIVO,LI_CODBAIRRO,PAVIMENTO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,"
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE  AND INATIVO=0  GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.INATIVO,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
Sql = Sql & " ORDER BY CADIMOB.CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    
    nNumRec = .RowCount
    Do Until .EOF
'        If !CODREDUZIDO <> 5397 Then GoTo proximo
        'GAUGE
        If xId Mod 100 = 0 Then
           CallPb xId, nNumRec
        End If
        If !Inativo = True Then GoTo proximo
'        If Not bExec Then
'           MsgBox "Cálculo Interrompido pelo usuário", vbCritical, "Atenção"
'           Exit Do
'        End If
        bIsento = False
        nTipoIsento = 0
        'DADOS DO IMOVEL
        nCodReduz = !CODREDUZIDO
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & nAnoCalculo
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                nTipoIsento = 1
                bIsento = True
            End If
           .Close
        End With
                
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND CODISENCAO=1"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                nTipoIsento = 2
                bIsento = True
            End If
           .Close
        End With
                
        nCodBairro = !Li_CodBairro
        nAreaTerreno = !Dt_AreaTerreno
        nAreaTerrenoReal = nAreaTerreno
        nCodSituacao = !Dt_CodSituacao
        nCodPedologia = !Dt_CodPedol
        nCodTopografia = !Dt_CodTopog
        nCodAgrupamento = !CODAGRUPA
                
        nTestadaPrincipal = 0
        nFracaoIdeal = !Dt_FracaoIdeal
        bFracaoIdeal = IIf(nFracaoIdeal > 0, True, False)
        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
        'TESTADAS
        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumTestadas = .RowCount
            If nNumTestadas = 0 Then
                nTestadaPrincipal = 1
                nTestada1 = 1
            Else
                If nNumTestadas = 1 Then
                    nTestadaPrincipal = !AREATESTADA
                    nTestada1 = !AREATESTADA
                Else
                    nSomaTestada = 0
                    Do Until .EOF
                       If !NUMFACE = RdoAux!Seq Then
                          nTestada1 = !AREATESTADA
                       End If
                       nSomaTestada = nSomaTestada + !AREATESTADA
                      .MoveNext
                    Loop
                    If nNumTestadas > 0 Then
                       nTestadaPrincipal = nSomaTestada / nNumTestadas
                    Else
                       nTestadaPrincipal = 1
                    End If
                End If
            End If
           .Close
        End With
        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
        
        'BUSCA ÁREA PRINCIPAL
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
        'TEM ÁREA?
            If .RowCount > 0 Then
                If Not IsNull(RdoAux!SOMAAREA) Then
                    If RdoAux!SOMAAREA <= 65 Then
                       bIsento = False
                    End If
                    bTemPredial = True
                    nAreaPrincipal = FormatNumber(RdoAux!SOMAAREA, 2)
                Else
                    bTemPredial = False
                    nAreaPrincipal = 0
                End If
                If bFracaoIdeal Then
                    If nAreaPrincipal > 0 Then
                       nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
                    Else
                       nTestadaPrincipal = 1
                    End If
                End If
                If bTemPredial Then
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                End If
            Else
                bTemPredial = False
                nAreaPrincipal = 0
            End If
            
            If Not bIsento Then
                GoTo proximo
            End If
            
            nValorVenalPredial = 0
            nFatorCategoria = 0
            If bTemPredial Then
                Do Until .EOF
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                    For x = 1 To UBound(aFatorC)
                        If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                           nFatorCategoria = aFatorC(x).Fator
                           Exit For
                        End If
                    Next
                    nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
                   .MoveNext
                Loop
            End If
           
           .Close
        End With
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
        End If
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For x = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
               nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
            Else
               nValorProfundidade = 1
            End If
            'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
            For x = 1 To UBound(aProf)
                If aProf(x).Distrito = !Distrito Then
                   If nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And nValorProfundidade <= FormatNumber(aProf(x).Max, 2) Then
                      Exit For
                   ElseIf nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And FormatNumber(aProf(x).Max, 2) = 0 Then
                      Exit For
                   End If
                End If
            Next
            On Error Resume Next
            nCodProfundidade = aProf(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
            nFatorProfundidade = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade = aFatorF(x).Fator
                   Exit For
                End If
            Next
        Else
            nFatorProfundidade = 1
        End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nFatorDistrito = aFatorD(!Distrito)
        nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito, 2)
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
'            nFatorDistrito = aFatorD(!Distrito)
'            nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito, 2)
            'CÁLCULO VALOR VENAL TERRITORIAL
            nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
            'FATOR DISTRITO 98
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            If nAnoCalculo < 2008 Then
                nValorVenalPredial = 0
                nFatorCategoria = 0
                For x = 1 To UBound(aFatorC)
                    If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                       nFatorCategoria = aFatorC(x).Fator
                       Exit For
                    End If
                Next
                nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
            End If
           'FATOR CATEGORIA 98
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        Else
            nValorVenalPredial = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100) '1.125
            nValorFinal = nValorIptu
            nValorITU = 0
            nAliquota = nAliquotaPredial
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
            nValorFinal = nValorITU
            nValorIptu = 0
            nAliquota = nAliquotaTerritorial
        End If
        nValorTotalIptu = nValorTotalIptu + nValorFinal 'relatorio
        nNumImovelCalc = nNumImovelCalc + 1 'relatorio
        
        nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
        nValorParcela = Round(nValorFinal / nNumParc, 2)
        
        Sql = "INSERT ISENTOIPTUREL(CODREDUZIDO,VVT,VVP,VVI,AREAT,AREAC,VALORIPTU,TIPOISENCAO) VALUES(" & nCodReduz & ","
        Sql = Sql & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nAreaTerreno)) & ","
        Sql = Sql & Virg2Ponto(CStr(nAreaPrincipal)) & "," & Virg2Ponto(CStr(nValorFinal)) & "," & nTipoIsento & ")"
        cn.Execute Sql, rdExecDirect
proximo:
        xId = xId + 1
       .MoveNext
    Loop
End With

FIM:

End Sub

Private Sub CalculoEspecial()
'Dim aCodigo() As String, sCodigo As String, x As Integer, nAreaTerreno As Double, nAreaConstrucao As Double, y As Integer
'Dim nVVP As Double, nVVT As Double, nVVI As Double, nValorIptu As Double, nCodReduz As Long, nValorParcela As Double
'Dim aVencto() As String, sVencto As String, nLastDoc As Long, dDataBase As Date, nAnoCalculo As Integer, nValorExp As Double, sObs As String, nSeq As Integer
'
'Exit Sub
'
'
'sCodigo = 25242
'
''datas de vencimento
'sVencto = "16/06/2006,17/07/2006,15/08/2006,15/09/2006,16/10/2006,16/11/2006,15/12/2006,15/01/2007,15/02/2007,15/03/2007,22/05/2007,15/06/2007"
'
''outras variaveis
'dDataBase = Format(Now, "dd/mm/yyyy"): nAnoCalculo = 2003: nValorExp = 1.13: nSeq = 3
'sObs = "Recalculo do IPTU 2003 efetuado conforme solicitação do cadastro técnico (Factore) em 25/04/2006"
'
''carrega as matrizes
'aCodigo = Split(sCodigo, ","): aVencto = Split(sVencto, ",")
'
''calcula cada código
'For x = 0 To UBound(aCodigo)
'    nCodReduz = CLng(aCodigo(x))
'    'por garantia vemos se os códigos podem ser recalculados
'    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=2003 AND CODLANCAMENTO=29  AND STATUSLANC<>3 AND STATUSLANC<>6 AND STATUSLANC<>8"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If .RowCount > 0 Then
''            MsgBox "Favor verificar o código " & nCodReduz, vbExclamation, "Atenção"
''            GoTo PROXIMO
'        End If
'       .Close
'    End With
'
'    'efetua o cálculo
'    Sql = "SELECT CADIMOB.CODREDUZIDO,CADIMOB.DT_AREATERRENO,SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
'    Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
'    Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where (CADIMOB.CODREDUZIDO = " & nCodReduz & ") GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
'    Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        nAreaTerreno = !Dt_AreaTerreno
'        nAreaConstrucao = !SOMAAREA
'        nVVT = nAreaTerreno * 2.11 'v.venal territorial
'        nVVP = nAreaConstrucao * 77.28 'v.venal predial
'        nVVI = nVVT + nVVP 'v.venal imovel
'        nValorIptu = nVVI * 0.015 'v.total do iptu
'        nValorParcela = nValorIptu / 12 'v.da parcela
'       .Close
'    End With
'
'    'busca ultimo numero de documento
'    Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        nLastDoc = !ULTIMO + 2
'       .Close
'    End With
'
'    'cancela todo débito que estiver nao pago em 2003 de iptu
'    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=2003 AND (CODLANCAMENTO=1 OR CODLANCAMENTO=29) AND STATUSLANC=3"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        Do Until .EOF
'            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=8 WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
'            Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
'            cn.Execute Sql, rdExecDirect
'            'grava observacao na parcela
''            On Error Resume Next
'            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & !AnoExercicio
'            Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
'            Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
'            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux
'                If IsNull(!MAXIMO) Then
'                    nSeq = 1
'                Else
'                    nSeq = !MAXIMO + 1
'                End If
'               .Close
'            End With
'            sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & !CODREDUZIDO & ","
'            Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeq & ",'" & sObs & "','"
'            Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
'            cn.Execute Sql, rdExecDirect
'            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & 2003
'            Sql = Sql & " AND CODLANCAMENTO=" & 29 & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
'            Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
'            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux
'                If IsNull(!MAXIMO) Then
'                    nSeq = 1
'                Else
'                    nSeq = !MAXIMO + 1
'                End If
'               .Close
'            End With
'            Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & !CODREDUZIDO & ","
'            Sql = Sql & 2003 & "," & 29 & "," & nSeq & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeq & ",'" & sObs & "','"
'            Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
'            cn.Execute Sql, rdExecDirect
'            'On Error GoTo 0
'           .MoveNext
'        Loop
'    End With
'
'
'    'gravar o débito (os debitos serao gerados com sequencia 2)
'    For y = 1 To 12
'       'GRAVA NA TABELA DEBITOPARCELA
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodReduz & "," & nAnoCalculo & ",29," & nSeq & "," & y & ",0,3,'"
'        Sql = Sql & Format(aVencto(y - 1), "mm/dd/yyyy") & "','" & Format(dDataBase, "mm/dd/yyyy") & "',1,'" & Left$(NomeDeLogin, 25) & "')"
'        cn.Execute Sql, rdExecDirect
'        'GRAVA NA TABELA DEBITO TRIBUTO
'        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
'        Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAnoCalculo & ",29," & nSeq & "," & y & ",0,1," & Virg2Ponto(CStr(nValorParcela)) & ")"
'        cn.Execute Sql, rdExecDirect
'        'GRAVA NA TABELA DEBITO TRIBUTO
'        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
'        Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAnoCalculo & ",29," & nSeq & "," & y & ",0,3," & Virg2Ponto(CStr(nValorExp)) & ")"
'        cn.Execute Sql, rdExecDirect
'        'GRAVA NA TABELA NUMDOCUMENTO
'        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor) VALUES("
'        Sql = Sql & nLastDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0," & Virg2Ponto(CStr(nValorExp)) & ",'" & NomeDeLogin & " (CALCGERAL-ESP)" & "')"
'        cn.Execute Sql, rdExecDirect
'        'GRAVA NA TABELA PARCELADOCUMENTO
'        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
'        Sql = Sql & nCodReduz & "," & nAnoCalculo & ",29," & nSeq & "," & y & ",0," & nLastDoc & ")"
'        cn.Execute Sql, rdExecDirect
'        nLastDoc = nLastDoc + 1
'    Next
'proximo:
'Next
'
''fim do cálculo
'MsgBox "Cálculo efetuado", vbInformation, "Aviso"

End Sub

Private Sub CalculoGeral2()

'Dim xId As Long, nNumRec As Long
'Dim nValorExpDocParc As Double, nValorExpDocUnica As Double, nLastDoc As Long, nAreaTerrenoReal As Double
'Dim ax As String, sDataBase As String, nAliquota As Double, RdoAux4 As rdoResultset
'Dim nValorUnica As Double, nValorParcela As Double, nTestada1 As Double, nFracaoIdeal As Double
''Relatorio
'Dim nValorTotalIptu As Double, nNumImovelCalc As Integer, nNumImovelOK As Integer, nNumImovelBloqueio As Integer
'Exit Sub
'nValorTotalIptu = 0: nNumImovelBloqueio = 0: nNumImovelCalc = 0: nNumImovelOK = 0
''nAnoCalculo = Val(txtAnoCalculo.text)
'cn.QueryTimeout = 0
'cmdCalculo.Enabled = False
'Sql = "DELETE FROM LASERIPTU WHERE ANO=" & nAnoCalculo
''cn.Execute Sql, rdExecDirect
'If cGetInputState() <> 0 Then DoEvents
'
'nAnoCalculo = 2008
'nNumParc = Val(txtNumParc.Text)
'
'Sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If !contador > 0 Then
'
'        'REMOVE RELACIONAMENTO
'        Sql = "BEGIN TRANSACTION SET QUOTED_IDENTIFIER ON "
' '       cn.Execute Sql, rdExecDirect
'        Sql = "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE COMMIT"
' '       cn.Execute Sql, rdExecDirect
'        Sql = "BEGIN TRANSACTION ALTER TABLE dbo.PARCELADOCUMENTO  DROP CONSTRAINT FK_PARCELADOCUMENTO_NUMDOCUMENTO COMMIT"
' '       cn.Execute Sql, rdExecDirect
'
'        'TABELA NUMDOCUMENTO
'        Sql = "DELETE FROM NUMDOCUMENTO WHERE NUMDOCUMENTO in ("
'        Sql = Sql & "SELECT NumDocumento From PARCELADOCUMENTO WHERE ANOEXERCICIO =" & nAnoCalculo & " AND CODLANCAMENTO = 1)"
''        cn.Execute Sql, rdExecDirect
'        If cGetInputState() <> 0 Then DoEvents
'
'        'RECRIA O RELACIONAMENTO
'        Sql = "BEGIN TRANSACTION SET QUOTED_IDENTIFIER ON"
' '       cn.Execute Sql, rdExecDirect
'        Sql = "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE COMMIT"
' '       cn.Execute Sql, rdExecDirect
'        Sql = "BEGIN TRANSACTION ALTER TABLE dbo.PARCELADOCUMENTO WITH NOCHECK ADD CONSTRAINT FK_PARCELADOCUMENTO_NUMDOCUMENTO FOREIGN KEY "
'        Sql = Sql & " ( NUMDOCUMENTO) REFERENCES dbo.NUMDOCUMENTO ( NUMDOCUMENTO ) COMMIT"
''        cn.Execute Sql, rdExecDirect
'
'        'TABELA PARCELA DOCUMENTO
'        Sql = "DELETE FROM PARCELADOCUMENTO WHERE ANOEXERCICIO =" & nAnoCalculo & " AND CODLANCAMENTO = 1"
' '       cn.Execute Sql, rdExecDirect
'        If cGetInputState() <> 0 Then DoEvents
'        'TABELA DEBITOTRIBUTO
'        Sql = "DELETE FROM DEBITOTRIBUTO WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
' '       cn.Execute Sql, rdExecDirect
'        If cGetInputState() <> 0 Then DoEvents
'        'TABELA DEBITOPAGO
'        Sql = "DELETE FROM DEBITOPAGO WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
' '       cn.Execute Sql, rdExecDirect
'        'TABELA DEBITOPARCELA
'        Sql = "DELETE FROM DEBITOPARCELA WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
' '       cn.Execute Sql, rdExecDirect
'    End If
'End With
'TESTE:
'If cGetInputState() <> 0 Then DoEvents
'
'sDataBase = mskDataBase.Text
'
'Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1
'Open sPathBin & "\DEBITOTRIBUTO.TXT" For Output As #2
'Open sPathBin & "\PARCELADOCUMENTO.TXT" For Output As #3
'Open sPathBin & "\NUMDOCUMENTO.TXT" For Output As #4
'
''********************************
'' TAXA DE EXPEDIÇÃO DE DOCUMENTO
''********************************
'Calculo:
'Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'     If .RowCount > 0 Then
'        nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
'        nValorExpDocUnica = FormatNumber(!ValorUnica, 2)
'     Else
'        MsgBox "Taxa de Expediente não cadastrada.", vbCritical, "Atenção"
'        GoTo fim
'     End If
'    .Close
'End With
''ULTIMO Nº DE DOCUMENTO
'Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    nLastDoc = !ULTIMO + 1000
'   .Close
'End With
'
''CÁLCULO
'Sql = "SELECT CADIMOB.CODREDUZIDO,CADIMOB.INATIVO,LI_CODBAIRRO,PAVIMENTO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
'Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,"
'Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
'Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
'Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE WHERE  INATIVO=0  GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
'Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.INATIVO,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
'Sql = Sql & " ORDER BY CADIMOB.CODREDUZIDO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    xId = 1
'    nNumRec = .RowCount
'    Do Until .EOF
'        'GAUGE
'        If xId Mod 100 = 0 Then
'           CallPb xId, nNumRec
'        End If
'        If !Inativo = True Then GoTo proximo
'        If Not bExec Then
'           MsgBox "Cálculo Interrompido pelo usuário", vbCritical, "Atenção"
'           Exit Do
'        End If
'        'DADOS DO IMOVEL
'        nCodReduz = !CODREDUZIDO
'        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
'        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & nAnoCalculo
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            If .RowCount > 0 Then
'                GoTo proximo
'            End If
'           .Close
'        End With
'
'        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
'        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND CODISENCAO=1"
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            If .RowCount > 0 Then
'                GoTo proximo
'            End If
'           .Close
'        End With
'
'        nCodBairro = !Li_CodBairro
'        nAreaTerreno = !Dt_AreaTerreno
'        nAreaTerrenoReal = nAreaTerreno
'        nCodSituacao = !Dt_CodSituacao
'        nCodPedologia = !Dt_CodPedol
'        nCodTopografia = !Dt_CodTopog
'        nCodAgrupamento = !CODAGRUPA
'
'        nTestadaPrincipal = 0
'        nFracaoIdeal = !Dt_FracaoIdeal
'        bFracaoIdeal = IIf(nFracaoIdeal > 0, True, False)
'        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
'        'TESTADAS
'        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            nNumTestadas = .RowCount
'            If nNumTestadas = 0 Then
'                nTestadaPrincipal = 1
'                nTestada1 = 1
'            Else
'                If nNumTestadas = 1 Then
'                    nTestadaPrincipal = !AREATESTADA
'                    nTestada1 = !AREATESTADA
'                Else
'                    nSomaTestada = 0
'                    Do Until .EOF
'                       If !NUMFACE = RdoAux!Seq Then
'                          nTestada1 = !AREATESTADA
'                       End If
'                       nSomaTestada = nSomaTestada + !AREATESTADA
'                      .MoveNext
'                    Loop
'                    If nNumTestadas > 0 Then
'                       nTestadaPrincipal = nSomaTestada / nNumTestadas
'                    Else
'                       nTestadaPrincipal = 1
'                    End If
'                End If
'            End If
'           .Close
'        End With
'        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
'        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
'
'        'BUSCA ÁREA PRINCIPAL
'        'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P' AND YEAR(DATAAPROVA) < " & nAnoCalculo
'        'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P' "
'        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'        'TEM ÁREA?
'            If .RowCount > 0 Then
'                If Not IsNull(RdoAux!SOMAAREA) Then
'                    If RdoAux!SOMAAREA <= 65 And !USOCONSTR = 1 And (!CATCONSTR = 4 Or !CATCONSTR = 7) And !QTDEPAV < 2 And nAreaTerreno < 600 Then
'                        Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO2 FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
'                        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
'                        If RdoAux4.RowCount = 0 Then
'                            GoTo proximo
'                        End If
'                    End If
'                    bTemPredial = True
'                    nAreaPrincipal = FormatNumber(RdoAux!SOMAAREA, 2)
'                Else
'                    bTemPredial = False
'                    nAreaPrincipal = 0
'                End If
'                If bFracaoIdeal Then
'                    If nAreaPrincipal > 0 Then
'                       nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
'                    Else
'                       nTestadaPrincipal = 1
'                    End If
'                End If
'                If bTemPredial Then
'                    nUso = !USOCONSTR
'                    nTipo = !TIPOCONSTR
'                    nCat = !CATCONSTR
'                End If
'            Else
'                bTemPredial = False
'                nAreaPrincipal = 0
'            End If
'
'            'novo VVP ***********************************
'            If nAnoCalculo > 2008 Then
'                nValorVenalPredial = 0
'                nFatorCategoria = 0
'                If bTemPredial Then
'                    Do Until .EOF
'                        nUso = !USOCONSTR
'                        nTipo = !TIPOCONSTR
'                        nCat = !CATCONSTR
'                        For x = 1 To UBound(aFatorC)
'                            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
'                               nFatorCategoria = aFatorC(x).Fator
'                               Exit For
'                            End If
'                        Next
'                        nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
'                       .MoveNext
'                    Loop
'                End If
'            Else
'                If bTemPredial Then
'                     nUso = !USOCONSTR
'                     nTipo = !TIPOCONSTR
'                     nCat = !CATCONSTR
'                End If
'            End If
'
'           .Close
'        End With
'
'        'VALOR DOS AGRUPAMENTOS
'        If !Dt_CodUsoTerreno = 6 Then
'           nValorAgrupamento = aFatorR(7)
'        Else
'           nValorAgrupamento = aFatorR(nCodAgrupamento)
'        End If
'        '**************************
'        'CÁLCULO DOS FATORES
'        '**************************
'        '**************************
'        '### FATOR GLEBA ###
'        '**************************
'        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
'        For x = 1 To UBound(aGleba)
'            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
'                 Exit For
'            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
'                 Exit For
'            End If
'        Next
'        nCodGleba = aGleba(x).Codigo
'        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
'        nFatorGleba = aFatorG(nCodGleba)
'        '**************************
'        '### FATOR PROFUNDIDADE ###
'        '**************************
'        If !Dt_CodUsoTerreno <> 6 Then
'            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
'            If nTestadaPrincipal > 0 Then
'               nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
'            Else
'               nValorProfundidade = 1
'            End If
'            'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
'            For x = 1 To UBound(aProf)
'                If aProf(x).Distrito = !Distrito Then
'                   If nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And nValorProfundidade <= FormatNumber(aProf(x).Max, 2) Then
'                      Exit For
'                   ElseIf nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And FormatNumber(aProf(x).Max, 2) = 0 Then
'                      Exit For
'                   End If
'                End If
'            Next
'            On Error Resume Next
'            nCodProfundidade = aProf(x).Codigo
'            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
'            nFatorProfundidade = 0
'            For x = 1 To UBound(aFatorF)
'                If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
'                   nFatorProfundidade = aFatorF(x).Fator
'                   Exit For
'                End If
'            Next
'        Else
'            nFatorProfundidade = 1
'        End If
'        '**************************
'        '### FATOR SITUAÇÃO ###
'        '**************************
'        nFatorSituacao = aFatorS(nCodSituacao)
'        '**************************
'        '### FATOR PEDOLOGIA ###
'        '**************************
'        nFatorPedologia = aFatorP(nCodPedologia)
'        '**************************
'        '### FATOR TOPOGRAFIA ###
'        '**************************
'        nFatorTopografia = aFatorT(nCodTopografia)
'        '**************************
'        'FIM DO CÁLCULO DOS FATORES
'        '**************************
'        'MULTIPLICA OS FATORES
'        nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba, 2)
'        'CÁLCULO VALOR VENAL TERRITORIAL
'        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
'        'CÁLCULO VALOR VENAL PREDIAL
'        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
'        If bTemPredial Then
'            '**************************
'            '### FATOR DISTRITO ###
'            '**************************
'            nFatorDistrito = aFatorD(!Distrito)
'            nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito, 2)
'            'CÁLCULO VALOR VENAL TERRITORIAL
'            nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
'            'FATOR DISTRITO 98
'            '**************************
'            '### FATOR CATEGORIA ###
'            '**************************
'            If nAnoCalculo < 2009 Then
'                nValorVenalPredial = 0
'                nFatorCategoria = 0
'                For x = 1 To UBound(aFatorC)
'                    If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
'                       nFatorCategoria = aFatorC(x).Fator
'                       Exit For
'                    End If
'                Next
'                nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
'            End If
''            nValorVenalPredial = 0
''            For x = 1 To UBound(aFatorC)
''                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
''                   nFatorCategoria = aFatorC(x).Fator
''                   Exit For
''                End If
''            Next
''            nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
'           'FATOR CATEGORIA 98
'            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
'        Else
'            nValorVenalPredial = 0
'        End If
'        'VALOR ITU/IPTU
'        If bTemPredial Then
'            nCodTributo = 1
'            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
'            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100) '1.125
'            nValorFinal = nValorIptu
'            nValorITU = 0
'            nAliquota = nAliquotaPredial
'        Else
'            nCodTributo = 2
'            nValorVenalImovel = nValorVenalTerritorial
'            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
'            nValorFinal = nValorITU
'            nValorIptu = 0
'            nAliquota = nAliquotaTerritorial
'        End If
'        nValorTotalIptu = nValorTotalIptu + nValorFinal 'relatorio
'        nNumImovelCalc = nNumImovelCalc + 1 'relatorio
'
'        '** NOVA ROTINA DE QTDE DE PARCELAS ***
'        If nValorFinal > 0 And nValorFinal <= 10 Then
'            nNumParc = 1
'        ElseIf nValorFinal > 10 And nValorFinal <= 20 Then nNumParc = 1
'        ElseIf nValorFinal > 20 And nValorFinal <= 30 Then nNumParc = 2
'        ElseIf nValorFinal > 30 And nValorFinal <= 40 Then nNumParc = 3
'        ElseIf nValorFinal > 40 And nValorFinal <= 50 Then nNumParc = 4
'        ElseIf nValorFinal > 50 And nValorFinal <= 60 Then nNumParc = 5
'        ElseIf nValorFinal > 60 And nValorFinal <= 70 Then nNumParc = 6
'        ElseIf nValorFinal > 70 And nValorFinal <= 80 Then nNumParc = 7
'        ElseIf nValorFinal > 80 And nValorFinal <= 90 Then nNumParc = 8
'        ElseIf nValorFinal > 90 And nValorFinal <= 100 Then nNumParc = 9
'        Else
'            nNumParc = Val(txtNumParc.Text)
'        End If
'        '**************************************
'
'        nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
'        nValorParcela = Round(nValorFinal / nNumParc, 2)
'
'        'GRAVA TABELA LASERIPTU
'        Sql = "INSERT LASERIPTU (ANO,CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,"
'        Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA,AREATERRENO,FATORCAT,FATORPED,FATORSIT,"
'        Sql = Sql & "FATORPRO,FATORTOP,FATORDIS,FATORGLE,AGRUPAMENTO,FRACAOIDEAL,ALIQUOTA) VALUES("
'        Sql = Sql & nAnoCalculo & "," & nCodReduz & "," & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nValorIptu)) & "," & Virg2Ponto(CStr(nValorITU)) & ",'"
'        Sql = Sql & IIf(bTemPredial, "Predial", "Territorial") & "'," & Virg2Ponto(CStr(nAreaPrincipal)) & "," & Virg2Ponto(CStr(nTestada1)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorParcela)) & "," & Virg2Ponto(CStr(nValorUnica)) & "," & nNumParc & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorExpDocParc) * Val(txtNumParc.Text)) & "," & Virg2Ponto(CStr(nValorExpDocUnica)) & "," & Virg2Ponto(CStr(nAreaTerreno)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nFatorCategoria)) & "," & Virg2Ponto(CStr(nFatorPedologia)) & "," & Virg2Ponto(CStr(nFatorSituacao)) & "," & Virg2Ponto(CStr(nFatorProfundidade)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nFatorTopografia)) & "," & Virg2Ponto(CStr(nFatorDistrito)) & "," & Virg2Ponto(CStr(nFatorGleba)) & "," & Virg2Ponto(CStr(nValorAgrupamento)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nFracaoIdeal)) & "," & Virg2Ponto(CStr(nAliquota)) & ")"
'        cn.Execute Sql, rdExecDirect
'
'        For x = 0 To nNumParc
'            If x = 0 And lblUnica.Caption = "Não" Then GoTo proximo
'            'GRAVA NA TABELA DEBITOPARCELA
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & 3 & "," & Format(aParc(x), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
'            ax = ax & 1 & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
'            ax = ax & Null & "," & 0
'            Print #1, ax
'            'GRAVA NA TABELA DEBITO TRIBUTO
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & nCodTributo & "," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & 3 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
'            'GRAVA NA TABELA NUMDOCUMENTO
'            nLastDoc = nLastDoc + 1
'            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2)))
'            Print #4, ax
'            'GRAVA NA TABELA PARCELADOCUMENTO
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & ","
'            ax = ax & x & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
'            Print #3, ax
'        Next
'proximo:
'        xId = xId + 1
'       .MoveNext
'    Loop
'End With
'
'fim:
'Close #4
'Close #3
'Close #2
'Close #1
'
'MsgBox "Valor total do IPTU: " & FormatNumber(nValorTotalIptu, 2)
'MsgBox "No de Imóveis Calculados: " & nNumImovelCalc
'MsgBox "No de Imóveis OK: " & nNumImovelOK
'MsgBox "No de Imóveis Bloqueados: " & nNumImovelBloqueio
'cmdCalculo.Enabled = True

End Sub


Private Sub cmdGravar_Click()
Dim nValorParcela As Double, nValorUnica As Double, nValorPago As Double, nSeq As Integer
Dim nValorFinal As Double, nNumParc As Integer

If Not IsDate(mskDataBase.Text) Then
    MsgBox "Data base inválida", vbCritical, "atenção"
    Exit Sub
End If

Sql = "select * from debitoparcela where codreduzido=" & Val(txtCod.Text) & " and anoexercicio=" & Val(txtAnoCalculo.Text) & " and codlancamento=1 and statuslanc<>5 and statuslanc<>8"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    MsgBox "Este código já possui IPTU para " & txtAnoCalculo.Text
    Exit Sub
End If
RdoAux.Close

'Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'     nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
'
'     nValorExpDocUnica = FormatNumber(!ValorUnica, 2)
'    .Close
'End With

'busca o valorpago
Sql = "SELECT SUM(debitotributo.valortributo) AS TOTAL FROM debitotributo INNER JOIN debitoparcela ON debitotributo.codreduzido = debitoparcela.codreduzido AND "
Sql = Sql & "debitotributo.anoexercicio = debitoparcela.anoexercicio AND debitotributo.codlancamento = debitoparcela.codlancamento AND "
Sql = Sql & "debitotributo.seqlancamento = debitoparcela.seqlancamento AND debitotributo.NumParcela = debitoparcela.NumParcela And debitotributo.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
Sql = Sql & "WHERE debitotributo.codreduzido = " & Val(txtCod.Text) & " AND debitotributo.anoexercicio = " & nAnoCalculo & " AND "
Sql = Sql & "debitotributo.codlancamento = 1 AND (debitoparcela.statuslanc = 1 OR debitoparcela.statuslanc = 2)"
'Sql = "SELECT SUM(VALORPAGO) AS TOTAL FROM DEBITOPAGO WHERE CODREDUZIDO=" & Val(txtCod) & " AND ANOEXERCICIO =' " & nAnoCalculo & " AND (CODLANCAMENTO = 1 OR CODLANCAMENTO = 29)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!Total) Then
        nValorPago = 0
    Else
        nValorPago = !Total
    End If
   .Close
End With

nCodReduz = Val(txtCod.Text)
nAnoCalculo = Val(txtAnoCalculo.Text)
sDataBase = mskDataBase.Text
nValorParcela = (CDbl(lblValorFinal.Caption) - nValorPago) / Val(txtNumParc.Text)
nValorUnica = CDbl(lblUnica.Caption)

'busca ultima sequencia
Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAnoCalculo & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 0
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With

'cancela as parcelas nao pagas
Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & nAnoCalculo & " AND CODLANCAMENTO=1 AND STATUSLANC=3"
cn.Execute Sql, rdExecDirect

'APAGA

'TABELA PARCELA DOCUMENTO
'Sql = "SELECT  parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento, "
'Sql = Sql & "parceladocumento.numparcela, parceladocumento.codcomplemento, parceladocumento.numdocumento, parceladocumento.valorjuros,"
'Sql = Sql & "parceladocumento.codbanco, parceladocumento.valormulta, parceladocumento.valorcorrecao, parceladocumento.intacto,"
'Sql = Sql & "debitoparcela.statuslanc FROM parceladocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND "
'Sql = Sql & "parceladocumento.anoexercicio = debitoparcela.anoexercicio AND parceladocumento.codlancamento = debitoparcela.codlancamento AND "
'Sql = Sql & "parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.NumParcela = debitoparcela.NumParcela And "
'Sql = Sql & "parceladocumento.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO WHERE parceladocumento.codreduzido = " & Val(txtCod.text) & " AND "
'Sql = Sql & "parceladocumento.anoexercicio = " & nAnoCalculo & " AND parceladocumento.codlancamento = 1 AND debitoparcela.statuslanc = 3"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        Sql = "DELETE FROM NUMDOCUMENTO WHERE NUMDOCUMENTO=" & !NumDocumento
'        cn.Execute Sql, rdExecDirect
'       .MoveNext
'    Loop
'End With
'
'Sql = "DELETE FROM parceladocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND "
'Sql = Sql & "parceladocumento.anoexercicio = debitoparcela.anoexercicio AND parceladocumento.codlancamento = debitoparcela.codlancamento AND "
'Sql = Sql & "parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.NumParcela = debitoparcela.NumParcela And "
'Sql = Sql & "parceladocumento.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO WHERE parceladocumento.codreduzido = " & Val(txtCod.text) & " AND "
'Sql = Sql & "parceladocumento.anoexercicio = " & nAnoCalculo & " AND parceladocumento.codlancamento = 1 AND debitoparcela.statuslanc = 3"
'cn.Execute Sql, rdExecDirect
'
''TABELA DEBITOTRIBUTO
'Sql = "DELETE FROM debitotributo FROM  debitoparcela INNER JOIN  debitotributo ON debitoparcela.codcomplemento = debitotributo.codcomplemento AND "
'Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
'Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.AnoExercicio = debitotributo.AnoExercicio And "
'Sql = Sql & "debitoparcela.CODREDUZIDO = debitotributo.CODREDUZIDO Where debitoparcela.statuslanc = 3 And debitoTRIBUTO.CODREDUZIDO = " & Val(txtCod.text)
'Sql = Sql & " And debitotributo.ANOEXERCICIO = " & nAnoCalculo & " and debitotributo.codlancamento=1"
'cn.Execute Sql, rdExecDirect
'
''TABELA DEBITOPARCELA
'Sql = "DELETE FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod) & " AND ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1 AND STATUSLANC=3"
'cn.Execute Sql, rdExecDirect


'ULTIMO Nº DE DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nLastDoc = !ULTIMO + 20
   .Close
End With


nValorFinal = (CDbl(lblValorFinal.Caption) - nValorPago)
'** NOVA ROTINA DE QTDE DE PARCELAS ***
If nValorFinal > 0 And nValorFinal <= 10 Then
    nNumParc = 1
ElseIf nValorFinal > 10 And nValorFinal <= 20 Then nNumParc = 1
ElseIf nValorFinal > 20 And nValorFinal <= 30 Then nNumParc = 2
ElseIf nValorFinal > 30 And nValorFinal <= 40 Then nNumParc = 3
ElseIf nValorFinal > 40 And nValorFinal <= 50 Then nNumParc = 4
ElseIf nValorFinal > 50 And nValorFinal <= 60 Then nNumParc = 5
ElseIf nValorFinal > 60 And nValorFinal <= 70 Then nNumParc = 6
ElseIf nValorFinal > 70 And nValorFinal <= 80 Then nNumParc = 7
ElseIf nValorFinal > 80 And nValorFinal <= 90 Then nNumParc = 8
ElseIf nValorFinal > 90 And nValorFinal <= 100 Then nNumParc = 9
Else
    nNumParc = Val(txtNumParc.Text)
End If
'**************************************

nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
nValorParcela = Round(nValorFinal / nNumParc, 2)



For x = 0 To nNumParc
    If lblTemUnica = "Não" And x = 0 Then x = 1
    'GRAVA NA TABELA DEBITOPARCELA
'    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & x & ",0,3,'"
'    Sql = Sql & Format(aParc(x), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1,'" & Left$(NomeDeLogin, 25) & "')"
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & x & ",0,3,'"
    Sql = Sql & Format(aParc(x), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1," & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA DEBITO TRIBUTO
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
    Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & x & ",0,1," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ")"
    cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA DEBITO TRIBUTO
'    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
'    Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.text) & "," & nAnoCalculo & ",1," & nSeq & "," & x & ",0,3," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ")"
'    cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA NUMDOCUMENTO
  '  Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,emissor) VALUES("
  '  Sql = Sql & nLastDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0," & Virg2Ponto(CStr(Round(nValorExpDocParc, 2))) & "," & "0" & ",'" & NomeDeLogin & " (CÁLCULO IND.)" & "')"
  '  cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA PARCELADOCUMENTO
  '  Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
  '  Sql = Sql & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & x & ",0," & nLastDoc & ")"
  '  cn.Execute Sql, rdExecDirect
  '  nLastDoc = nLastDoc + 1
Next
'Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & 0 & ",91,3,'"
'Sql = Sql & Format(aParc(2), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1,'" & Left$(NomeDeLogin, 25) & "')"
Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & 0 & ",91,3,'"
Sql = Sql & Format(aParc(2), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect
'GRAVA NA TABELA DEBITO TRIBUTO
Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & 0 & ",91,1," & Virg2Ponto(Round(CDbl(lblUnica2.Caption), 2)) & ")"
cn.Execute Sql, rdExecDirect

'Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & 0 & ",92,3,'"
'Sql = Sql & Format(aParc(3), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1,'" & Left$(NomeDeLogin, 25) & "')"
Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & 0 & ",92,3,'"
Sql = Sql & Format(aParc(3), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect
'GRAVA NA TABELA DEBITO TRIBUTO
Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.Text) & "," & nAnoCalculo & ",1," & nSeq & "," & 0 & ",92,1," & Virg2Ponto(Round(CDbl(lblUnica3.Caption), 2)) & ")"
cn.Execute Sql, rdExecDirect



MsgBox "Debito recalculado."

End Sub

Private Sub cmdLista_Click()
Dim aCidadao1() As Lista, aCidadao2() As Long, x As Long, Y As Long, z As Long, bAchou As Boolean, aDup() As Lista, bAchou2 As Boolean, t As Long

ReDim aCidadao1(0): ReDim aCidadao2(0): ReDim aDup(0): ReDim aEspelho(0)
'CARREGA A MATRIZ 1 COM TODOS OS CODIGOS
Sql = "SELECT codreduzido,codcidadao From Proprietario WHERE tipoprop = 'P' AND principal = 1 ORDER BY codcidadao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux2.RowCount > 0 Then
            ReDim Preserve aCidadao1(UBound(aCidadao1) + 1)
            aCidadao1(UBound(aCidadao1)).nCodReduzido = !CODREDUZIDO
            aCidadao1(UBound(aCidadao1)).nCodCidadao = !CodCidadao
        End If
        RdoAux2.Close
       .MoveNext
    Loop
   .Close
End With

'INTERAÇÃO PELA MATRIZ 1
For x = 1 To UBound(aCidadao1)
    bAchou = False
    'VERIFICA SE O ITEM EXISTE EM MATRIZ 2
    For Y = 1 To UBound(aCidadao2)
        If aCidadao2(Y) = aCidadao1(x).nCodCidadao Then
            bAchou = True
            Exit For
        End If
    Next Y
    If bAchou Then
       'SE ACHAR VERIFICA SE JA EXISTE NA MATRIZ DE DUPLICADOS
        bAchou2 = False
        For z = 1 To UBound(aDup)
            If aDup(z).nCodCidadao = aCidadao1(x).nCodCidadao Then
                bAchou2 = True
                Exit For
            End If
        Next z
       'SE NÃO ACHAR NA MATRIZ DE DUPLICADOS, INCLUI NA MATRIZ DUPLICADOS
       'MAS ANTES É NECESSARIO PEGAR TODOS OS CÓDIGOS DE IMÓVEIS DESTE PROPRIETÁRIO DA MATRIZ 1
        If Not bAchou2 Then
            For t = 1 To UBound(aCidadao1)
                If aCidadao1(t).nCodCidadao = aCidadao1(x).nCodCidadao Then
                    ReDim Preserve aDup(UBound(aDup) + 1)
                    aDup(UBound(aDup)).nCodReduzido = aCidadao1(t).nCodReduzido
                    aDup(UBound(aDup)).nCodCidadao = aCidadao1(t).nCodCidadao
                End If
            Next
        End If
    Else
       'SE NÃO ACHAR INCLUI O ITEM DA MATRIZ 1 NA MATR1Z 2
        ReDim Preserve aCidadao2(UBound(aCidadao2) + 1)
        aCidadao2(UBound(aCidadao2)) = aCidadao1(x).nCodCidadao
    End If
proximo:
Next x

Sql = "TRUNCATE TABLE PROPRIETARIODUPLICADO"
cn.Execute Sql, rdExecDirect

On Error Resume Next
For x = 1 To UBound(aDup)
    Sql = "INSERT PROPRIETARIODUPLICADO(CODREDUZIDO,CODCIDADAO) VALUES(" & aDup(x).nCodReduzido & "," & aDup(x).nCodCidadao & ")"
    cn.Execute Sql, rdExecDirect
Next

MsgBox "fim"

End Sub

Private Sub cmdPrint_Click()
Me.PrintForm
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim x As Long, Sql As String
nAnoCalculo = 2012
LoadMatrix
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3

CalculoGeralIsentos
'For x = 30156 To 30258
'    CalculoIndividualporLista (x)
'    Sql = "UPDATE LASERIPTU SET VVT=" & Virg2Ponto(RemovePonto(lblVVT.Caption)) & ",VVI=" & Virg2Ponto(RemovePonto(lblVVI.Caption)) & ","
'    Sql = Sql & "TESTADAPRINC=" & Virg2Ponto(RemovePonto(lblTestada.Caption)) & ",FRACAOIDEAL=" & Virg2Ponto(RemovePonto(lblFracao.Caption)) & ","
'    Sql = Sql & "AREATERRENO=" & Virg2Ponto(RemovePonto(lblAreaTerreno.Caption)) & ",FATORCAT=" & Virg2Ponto(RemovePonto(lblFatorC.Caption)) & ","
'    Sql = Sql & "FATORPED=" & Virg2Ponto(RemovePonto(lblFatorP.Caption)) & ",FATORSIT=" & Virg2Ponto(RemovePonto(lblFatorS.Caption)) & ","
'    Sql = Sql & "FATORPRO=" & Virg2Ponto(RemovePonto(lblFatorF.Caption)) & ",FATORTOP=" & Virg2Ponto(RemovePonto(lblFatorT.Caption)) & ","
'    Sql = Sql & "FATORDIS=" & Virg2Ponto(RemovePonto(lblFatorD.Caption)) & ",FATORGLE=" & Virg2Ponto(RemovePonto(lblFatorG.Caption))
'    Sql = Sql & " WHERE CODREDUZIDO=" & x & " AND ANO=2010"
'    cn.Execute Sql, rdExecDirect
'Next

MsgBox "fim"
End Sub

Private Sub Form_Load()

Ocupado

If NomeDeLogin = "ANA" Or NomeDeLogin = "JOSIANE" Or NomeDeLogin = "SCHWARTZ" Then
   cmdGravar.Enabled = True
Else
   cmdGravar.Enabled = False
End If

Set xImovel = New clsImovel

Centraliza Me
Pb.value = 0
lblPB.Caption = "0 %"
If Val(txtAnoCalculo.Text) = 0 Then txtAnoCalculo.Text = Year(Now)
nAnoCalculo = txtAnoCalculo.Text
CarregaTela
lblAno.Caption = "Cálculo " & txtAnoCalculo.Text
Sql = "SELECT COUNT(CODREDUZIDO) AS TOTAL FROM CADIMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
       lblEstimado.Caption = !Total
     End If
      .Close
End With
FIM:
'LoadMatrix

Liberado

sRet = RetEventUserForm(Me.Name)
sTitOld = Me.Caption
End Sub

Private Sub LoadMatrix()

ReDim aFatorD(3)
ReDim aFatorD98(3)
ReDim aFatorP(6)
ReDim aFatorP98(6)
ReDim aFatorT(6)
ReDim aFatorT98(6)
ReDim aFatorS(6)
ReDim aFatorS98(6)
ReDim aFatorG(23)
ReDim aFatorG98(23)
ReDim aFatorR(8)
ReDim aFatorR98(8)

Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .Close
End With

ReDim aProf(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aProf(UBound(aProf) + 1)
        aProf(UBound(aProf)).Distrito = !CODDISTRITO
        aProf(UBound(aProf)).Codigo = !CODPROFUN
        aProf(UBound(aProf)).Min = !MINPROFUN
        aProf(UBound(aProf)).Max = !MAXPROFUN
       .MoveNext
     Loop
    .Close
End With


ReDim aFatorF(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aGleba(0)
Sql = "SELECT CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY CODGLEBA "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aGleba(UBound(aGleba) + 1)
        aGleba(UBound(aGleba)).Codigo = !CODGLEBA
        aGleba(UBound(aGleba)).Min = !MINGLEBA
        aGleba(UBound(aGleba)).Max = !MAXGLEBA
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorC(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !CODUSO
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !CODCATEG
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

'Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub mskDataBase_GotFocus()
mskDataBase.SetFocus
End Sub

Private Sub Pb_Click()
CalculoEspecial
End Sub

Private Sub txtAnoCalculo_KeyPress(KeyAscii As Integer)
Tweak txtAnoCalculo, KeyAscii, IntegerPositive
End Sub

Private Sub txtAnoCalculo_LostFocus()
If Val(txtAnoCalculo.Text) >= 2004 And Val(txtAnoCalculo.Text) <= 2019 Then
    lblAno.Caption = "Cálculo " & txtAnoCalculo.Text
   CarregaTela
Else
    MsgBox "Ano Inválido.", vbCritical, "atenção"
End If
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   cmdCalculo_Click
Else
   Tweak txtCod, KeyAscii, IntegerPositive
End If
End Sub

Private Sub Limpa()

lblVVP.Caption = "0,00"

End Sub

Private Sub CarregaImovel()

With xImovel
    .CarregaImovel Val(txtCod.Text)
    lblProp.Caption = .NomePropPrincipal
    lblRua.Caption = .EnderecoCompleto
End With

End Sub

Private Sub CarregaTela()
nAnoCalculo = Val(txtAnoCalculo.Text)
Sql = "SELECT ANO,QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,DESCONTOUNICA2,DESCONTOUNICA3,VENCUNICA,VENC01,VENC02,VENC03,VENC04,VENC05,"
Sql = Sql & "VENC06,VENC07,VENC08,VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA WHERE CODTIPO=1 AND ANO=" & nAnoCalculo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount = 0 Then GoTo FIM
     txtNumParc.Text = !qtdeparcela
     lblTemUnica.Caption = IIf(!PARCELAUNICA = "S", "Sim", "Não")
     lblPercUnica.Caption = FormatNumber(!DESCONTOUNICA, 2)
     lblPercUnica2.Caption = FormatNumber(!DESCONTOUNICA2, 2)
     lblPercUnica3.Caption = FormatNumber(!DESCONTOUNICA3, 2)
     ReDim aParc(!qtdeparcela)
     Do Until .EOF
'        If lblTemUnica.Caption = "Sim" Then
            If Not IsNull(!vencunica) Then aParc(0) = Format(!vencunica, "dd/mm/yyyy")
 '       End If
        If Not IsNull(!venc01) Then aParc(1) = Format(!venc01, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc02) Then aParc(2) = Format(!venc02, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc03) Then aParc(3) = Format(!venc03, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc04) Then aParc(4) = Format(!venc04, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc05) Then aParc(5) = Format(!venc05, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc06) Then aParc(6) = Format(!venc06, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc07) Then aParc(7) = Format(!venc07, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc08) Then aParc(8) = Format(!venc08, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc09) Then aParc(9) = Format(!venc09, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc10) Then aParc(10) = Format(!venc10, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc11) Then aParc(11) = Format(!venc11, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc12) Then aParc(12) = Format(!venc12, "dd/mm/yyyy") Else Exit Do
        x = x + 1
       .MoveNext
     Loop
    .Close
End With
FIM:
End Sub

Private Sub cmdRel_Click()
Dim nPos As Long, nTot As Long, nCodReduz As Long, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim sEnd As String, sCompl As String, sBairro As String, sCid As String, sCep As String, sUF As String

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
If MsgBox("Executar o relatório de IPTU?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

ReDim aLaser(1)
Open sPathBin & "\DETALHEIPTU.TXT" For Output As #1
Sql = "SELECT laseriptu.ano,laseriptu.codreduzido,laseriptu.vvt, laseriptu.vvc, laseriptu.vvi, laseriptu.impostopredial, laseriptu.impostoterritorial, laseriptu.natureza,"
Sql = Sql & "laseriptu.areaconstrucao, laseriptu.testadaprinc, laseriptu.valortotalparc, laseriptu.valortotalunica, laseriptu.qtdeparc, laseriptu.txexpparc,"
Sql = Sql & "laseriptu.txexpunica, laseriptu.areaterreno, laseriptu.fatorcat, laseriptu.fatorped, laseriptu.fatorsit, laseriptu.fatorpro, laseriptu.fatortop,"
Sql = Sql & "laseriptu.fatordis, laseriptu.fatorgle, laseriptu.agrupamento, laseriptu.fracaoideal, laseriptu.aliquota, cadimob.distrito, cadimob.setor, cadimob.quadra,"
Sql = Sql & "cadimob.lote, cadimob.seq, cadimob.unidade, cadimob.subunidade, cadimob.codcondominio, facequadra.codlogr, vwLOGRADOURO.ABREVTIPOLOG,vwLOGRADOURO.ABREVTITLOG,"
Sql = Sql & "vwLOGRADOURO.NOMELOGRADOURO, cadimob.li_num, cadimob.li_compl, bairro.descbairro,Cidadao.nomecidadao FROM laseriptu INNER JOIN cadimob ON laseriptu.codreduzido = cadimob.codreduzido INNER JOIN "
Sql = Sql & "facequadra ON cadimob.distrito = facequadra.coddistrito AND cadimob.setor = facequadra.codsetor AND cadimob.quadra = facequadra.codquadra AND "
Sql = Sql & "cadimob.seq = facequadra.codface INNER JOIN  vwLOGRADOURO ON facequadra.codlogr = vwLOGRADOURO.CODLOGRADOURO INNER JOIN bairro ON cadimob.li_uf = bairro.siglauf AND cadimob.li_codcidade = bairro.codcidade AND "
Sql = Sql & "cadimob.li_codbairro = bairro.codbairro INNER JOIN proprietario ON laseriptu.codreduzido = proprietario.codreduzido INNER JOIN cidadao ON proprietario.codcidadao = cidadao.codcidadao "
Sql = Sql & "WHERE (proprietario.tipoprop = 'P') AND (proprietario.principal = 1) AND ANO=" & Val(txtAnoCalculo.Text) & "  ORDER BY LASERIPTU.CODREDUZIDO,LASERIPTU.ANO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Pb.value = 0: nTot = .RowCount: nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
            CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        
        If nCodReduz <> aLaser(1).nCodReduz Then
            If aLaser(1).nCodReduz > 0 Then
                With aLaser(1)
                    ax = .nCodReduz & "@" & .nDistrito & "@" & .nSetor & "@" & .nQuadra & "@" & .nLote & "@" & .nFace & "@" & .nUnidade & "@" & .nSubUnidade & "@" & .sProprietario & "@"
                    ax = ax & .nCodLogradouro & "@" & .sEndereco & "@" & .nNumero & "@" & .sComplemento & "@" & .sBairro & "@" & .sEndEntrega & "@" & .sComplEntrega & "@"
                    ax = ax & .sCepEntrega & "@" & .sBairroEntrega & "@" & .sCidadeEntrega & "@" & .sUFEntrega & "@" & .sNatureza & "@" & FormatNumber(.nTestadaPrincipal, 2) & "@"
                    ax = ax & FormatNumber(.nAreaTerreno, 2) & "@" & FormatNumber(.nAreaConstruida, 2) & "@" & FormatNumber(.nFracaoIdeal, 2) & "@" & FormatNumber(.nAliquota, 2) & "@" & FormatNumber(.nAgrupamento, 2) & "@" & FormatNumber(.nFatorCat, 2) & "@" & FormatNumber(.nFatorPed, 2) & "@"
                    ax = ax & FormatNumber(.nFatorSit, 2) & "@" & FormatNumber(.nFatorPro, 2) & "@" & FormatNumber(.nFatorDis, 2) & "@" & FormatNumber(.nfatorGle, 2) & "@" & .nQtdeParc & "@"
                    ax = ax & FormatNumber(.nVVT1, 2) & "@" & FormatNumber(.nVVC1, 2) & "@" & FormatNumber(.nVVI1, 2) & "@" & FormatNumber(.nVVT2, 2) & "@" & FormatNumber(.nVVC2, 2) & "@" & FormatNumber(.nVVI2, 2) & "@"
                    ax = ax & FormatNumber(.nImpPre1, 2) & "@" & FormatNumber(.nImpTer1, 2) & "@" & FormatNumber(.nImpPre2, 2) & "@" & FormatNumber(.nImpTer2, 2) & "@"
                    ax = ax & FormatNumber(.nValorParcela1, 2) & "@" & FormatNumber(.nValorUnica1, 2) & "@" & FormatNumber(.nValorParcela2, 2) & "@" & FormatNumber(.nValorUnica2, 2)
                End With
                Print #1, ax
            End If
            ReDim aLaser(0): ReDim aLaser(1)
            aLaser(1).nCodReduz = !CODREDUZIDO
            aLaser(1).nDistrito = !Distrito
            aLaser(1).nSetor = !Setor
            aLaser(1).nQuadra = !Quadra
            aLaser(1).nLote = !Lote
            aLaser(1).nFace = !Seq
            aLaser(1).nUnidade = !Unidade
            aLaser(1).nSubUnidade = !SubUnidade
            aLaser(1).sProprietario = !NomeCidadao
            aLaser(1).nCodLogradouro = !CodLogr
            aLaser(1).sEndereco = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
            aLaser(1).nNumero = !Li_Num
            aLaser(1).sComplemento = SubNull(!Li_Compl)
            aLaser(1).sBairro = !DescBairro
            aLaser(1).sNatureza = !Natureza
            aLaser(1).nTestadaPrincipal = !TESTADAPRINC
            aLaser(1).nAreaTerreno = !AreaTerreno
            aLaser(1).nAreaConstruida = !areaconstrucao
            aLaser(1).nFracaoIdeal = !FracaoIdeal
            aLaser(1).nAliquota = !Aliquota
            aLaser(1).nAgrupamento = !Agrupamento
            aLaser(1).nFatorCat = !FATORCAT
            aLaser(1).nFatorPed = !FATORPED
            aLaser(1).nFatorSit = !FATORSIT
            aLaser(1).nFatorPro = !FATORPRO
            aLaser(1).nFatorTop = !FATORTOP
            aLaser(1).nFatorDis = !FATORDIS
            aLaser(1).nfatorGle = !FATORGLE
            aLaser(1).nQtdeParc = Val(SubNull(!qtdeparc))
           '**BUSCA ENDEREÇO DE ENTREGA***
            Sql = "SELECT vwCnsImovel.*,proprietario.codcidadao, cidadao.nomecidadao FROM vwCnsImovel INNER JOIN proprietario ON vwCnsImovel.codreduzido = proprietario.codreduzido INNER JOIN "
            Sql = Sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao WHERE (proprietario.tipoprop = 'P') AND (proprietario.principal = 1) AND VWCNSIMOVEL.CODREDUZIDO=" & nCodReduz
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If !Ee_TipoEnd = 0 Then
                    '***ENDEREÇO DO IMÓVEL***
                    If Not IsNull(!NomeLogradouro) Then
                        sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro) & ", " & SubNull(!Li_Num)
                    Else
                        sEnd = ""
                    End If
                    sCompl = SubNull(!Li_Compl)
                    sBairro = SubNull(!DescBairro)
                    sCid = SubNull(!descCidade)
                    sCep = RetornaCEP(!CodLogr, !Li_Num)
                    sUF = SubNull(!li_uf)
                ElseIf !Ee_TipoEnd = 1 Then
                    '***ENDEREÇO DO PROPRIETÁRIO
                    Sql = "SELECT cidadao.nomecidadao, cidadao.codcidadao, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, cidadao.siglauf, "
                    Sql = Sql & "cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG, vwLOGRADOURO.NOMELOGRADOURO,bairro.DescBairro , Cidade.desccidade "
                    Sql = Sql & "FROM cidade INNER JOIN bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade RIGHT OUTER JOIN "
                    Sql = Sql & "cidadao LEFT OUTER JOIN vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO ON bairro.siglauf = cidadao.siglauf AND "
                    Sql = Sql & "bairro.codcidade = Cidadao.codcidade And bairro.codbairro = Cidadao.codbairro WHERE CIDADAO.CODCIDADAO=" & !CODREDUZIDO
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        If .RowCount > 0 Then
                            If Not IsNull(!NomeLogradouro) Then
                                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro) & ", " & SubNull(!NUMIMOVEL)
                            Else
                                sEnd = ""
                            End If
                            sCompl = SubNull(!Complemento)
                            sBairro = SubNull(!DescBairro)
                            sCid = SubNull(!descCidade)
                            If Not IsNull(!CodLogradouro) Then
                                sCep = RetornaCEP(!CodLogradouro, Val(SubNull(!NUMIMOVEL)))
                            End If
                            sUF = SubNull(!SiglaUF)
                        Else
'                            MsgBox "erro náo achei endereço cidadao " & nCodReduz
                        End If
                       .Close
                    End With
                ElseIf !Ee_TipoEnd = 2 Then
                    '***ENDEREÇO DE ENTREGA
                    Sql = "SELECT endentrega.*, cidadao.nomecidadao, bairro.descbairro, cidade.desccidade FROM cidade INNER JOIN bairro ON cidade.siglauf = bairro.siglauf AND cidade.codcidade = bairro.codcidade RIGHT OUTER JOIN "
                    Sql = Sql & "endentrega INNER JOIN proprietario ON endentrega.codreduzido = proprietario.codreduzido INNER JOIN cidadao ON proprietario.codcidadao = cidadao.codcidadao ON bairro.siglauf = endentrega.ee_uf AND "
                    Sql = Sql & "bairro.codcidade = endentrega.ee_cidade AND bairro.codbairro = endentrega.ee_bairro WHERE (proprietario.tipoprop = 'P') AND (proprietario.principal = 1)  and ENDENTREGA.CODREDUZIDO = " & RdoAux!CODREDUZIDO
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux3
                        If Not IsNull(!Ee_NomeLog) Then
                            sEnd = Trim$(SubNull(!Ee_NomeLog)) & ", " & SubNull(!Ee_NumImovel)
                        Else
                            sEnd = ""
                        End If
                        sCompl = SubNull(!Ee_Complemento)
                        sBairro = SubNull(!DescBairro)
                        sCid = SubNull(!descCidade)
                        sCep = SubNull(!Ee_Cep)
                        sUF = SubNull(!Ee_Uf)
                       .Close
                    End With
                End If
               .Close
            End With
            aLaser(1).sEndEntrega = sEnd
            aLaser(1).sComplEntrega = sCompl
            aLaser(1).sBairroEntrega = sBairro
            aLaser(1).sCidadeEntrega = sCid
            aLaser(1).sUFEntrega = sUF
            aLaser(1).sCepEntrega = sCep
            
           '******************************
            If !Ano = 2006 Then
                aLaser(1).nVVT1 = !vvt
                aLaser(1).nVVC1 = !vvc
                aLaser(1).nVVI1 = !VVI
                aLaser(1).nImpPre1 = !impostopredial
                aLaser(1).nImpTer1 = !IMPOSTOTERRITORIAL
                aLaser(1).nValorParcela1 = !valortotalparc
                aLaser(1).nValorUnica1 = !VALORTOTALUNICA
            Else
                aLaser(1).nVVT2 = !vvt
                aLaser(1).nVVC2 = !vvc
                aLaser(1).nVVI2 = !VVI
                aLaser(1).nImpPre2 = !impostopredial
                aLaser(1).nImpTer2 = !IMPOSTOTERRITORIAL
                aLaser(1).nValorParcela2 = !valortotalparc
                aLaser(1).nValorUnica2 = !VALORTOTALUNICA
            End If
        Else
            If !Ano = 2006 Then
                aLaser(1).nVVT1 = !vvt
                aLaser(1).nVVC1 = !vvc
                aLaser(1).nVVI1 = !VVI
                aLaser(1).nImpPre1 = !impostopredial
                aLaser(1).nImpTer1 = !IMPOSTOTERRITORIAL
                aLaser(1).nValorParcela1 = !valortotalparc
                aLaser(1).nValorUnica1 = !VALORTOTALUNICA
            Else
                aLaser(1).nVVT2 = !vvt
                aLaser(1).nVVC2 = !vvc
                aLaser(1).nVVI2 = !VVI
                aLaser(1).nImpPre2 = !impostopredial
                aLaser(1).nImpTer2 = !IMPOSTOTERRITORIAL
                aLaser(1).nValorParcela2 = !valortotalparc
                aLaser(1).nValorUnica2 = !VALORTOTALUNICA
            End If
        End If
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Close #1
Pb.value = 100: lblPB.Caption = "100"

MsgBox "FIM"
End Sub

Private Sub CalculoGeralEstimado()
Dim xId As Long, nNumRec As Long, nSomaFatorCat As Double
Dim nValorExpDocParc As Double, nValorExpDocUnica As Double, nLastDoc As Long, nAreaTerrenoReal As Double
Dim ax As String, sDataBase As String, nAliquota As Double
Dim nValorUnica As Double, nValorUnica2 As Double, nValorUnica3 As Double, nValorParcela As Double, nTestada1 As Double, nFracaoIdeal As Double
'Relatorio
Dim nValorTotalIptu As Double, nNumImovelCalc As Integer, nNumImovelOK As Integer, nNumImovelBloqueio As Integer

nValorTotalIptu = 0: nNumImovelBloqueio = 0: nNumImovelCalc = 0: nNumImovelOK = 0
'nAnoCalculo = Val(txtAnoCalculo.text)
cn.QueryTimeout = 0
cmdCalculo.Enabled = False
Sql = "DELETE FROM LASERIPTU WHERE ANO=" & nAnoCalculo
'cn.Execute Sql, rdExecDirect
If cGetInputState() <> 0 Then DoEvents

'nAnoCalculo = 2007
nNumParc = Val(txtNumParc.Text)

sDataBase = mskDataBase.Text

Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTO.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTO.TXT" For Output As #3
Open sPathBin & "\NUMDOCUMENTO.TXT" For Output As #4

'********************************
' TAXA DE EXPEDIÇÃO DE DOCUMENTO
'********************************
Calculo:
Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
        nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
        nValorExpDocUnica = FormatNumber(!ValorUnica, 2)
     Else
        MsgBox "Taxa de Expediente não cadastrada.", vbCritical, "Atenção"
        Exit Sub
     End If
    .Close
End With
'ULTIMO Nº DE DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nLastDoc = !ULTIMO + 1000
   .Close
End With

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,CADIMOB.INATIVO,LI_CODBAIRRO,PAVIMENTO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,"
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE WHERE  INATIVO=0  GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.INATIVO,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
Sql = Sql & " ORDER BY CADIMOB.CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    nNumRec = .RowCount
    Do Until .EOF
        'GAUGE
        If xId Mod 100 = 0 Then
           CallPb xId, nNumRec
        End If
        If !Inativo = True Then GoTo proximo
        If Not bExec Then
           MsgBox "Cálculo Interrompido pelo usuário", vbCritical, "Atenção"
           Exit Do
        End If
        'DADOS DO IMOVEL
        nCodReduz = !CODREDUZIDO
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & nAnoCalculo
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                GoTo proximo
            End If
           .Close
        End With
                
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND CODISENCAO=1"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                GoTo proximo
            End If
           .Close
        End With
                
        nCodBairro = !Li_CodBairro
        nAreaTerreno = !Dt_AreaTerreno
        nAreaTerrenoReal = nAreaTerreno
        nCodSituacao = !Dt_CodSituacao
        nCodPedologia = !Dt_CodPedol
        nCodTopografia = !Dt_CodTopog
        nCodAgrupamento = !CODAGRUPA
                
        nTestadaPrincipal = 0
        nFracaoIdeal = !Dt_FracaoIdeal
        bFracaoIdeal = IIf(nFracaoIdeal > 0, True, False)
        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
        'TESTADAS
        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumTestadas = .RowCount
            If nNumTestadas = 0 Then
                nTestadaPrincipal = 1
                nTestada1 = 1
            Else
                If nNumTestadas = 1 Then
                    nTestadaPrincipal = !AREATESTADA
                    nTestada1 = !AREATESTADA
                Else
                    nSomaTestada = 0
                    Do Until .EOF
                       If !NUMFACE = RdoAux!Seq Then
                          nTestada1 = !AREATESTADA
                       End If
                       nSomaTestada = nSomaTestada + !AREATESTADA
                      .MoveNext
                    Loop
                    If nNumTestadas > 0 Then
                       nTestadaPrincipal = nSomaTestada / nNumTestadas
                    Else
                       nTestadaPrincipal = 1
                    End If
                End If
            End If
           .Close
        End With
        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
        
        'BUSCA ÁREA PRINCIPAL
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P' AND YEAR(DATAAPROVA) < " & nAnoCalculo
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
        'TEM ÁREA?
            If .RowCount > 0 Then
                If Not IsNull(RdoAux!SOMAAREA) Then
                    If RdoAux!SOMAAREA <= 65 And !USOCONSTR = 1 And (!CATCONSTR = 4 Or !CATCONSTR = 7) And !QTDEPAV < 2 And nAreaTerreno < 600 Then
                        GoTo proximo
                    End If
                    bTemPredial = True
                    nAreaPrincipal = FormatNumber(RdoAux!SOMAAREA, 2)
                Else
                    bTemPredial = False
                    nAreaPrincipal = 0
                End If
                If bFracaoIdeal Then
                    If nAreaPrincipal > 0 Then
                       nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
                    Else
                       nTestadaPrincipal = 1
                    End If
                End If
                If bTemPredial Then
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                End If
            Else
                bTemPredial = False
                nAreaPrincipal = 0
            End If
           .Close
        End With
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
        End If
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For x = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
               nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
            Else
               nValorProfundidade = 1
            End If
            'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
            For x = 1 To UBound(aProf)
                If aProf(x).Distrito = !Distrito Then
                   If nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And nValorProfundidade <= FormatNumber(aProf(x).Max, 2) Then
                      Exit For
                   ElseIf nValorProfundidade >= FormatNumber(aProf(x).Min, 2) And FormatNumber(aProf(x).Max, 2) = 0 Then
                      Exit For
                   End If
                End If
            Next
            On Error Resume Next
            nCodProfundidade = aProf(x).Codigo
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
            nFatorProfundidade = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade = aFatorF(x).Fator
                   Exit For
                End If
            Next
        Else
            nFatorProfundidade = 1
        End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba, 2)
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
            nFatorDistrito = aFatorD(!Distrito)
            nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba * nFatorDistrito, 2)
            'CÁLCULO VALOR VENAL TERRITORIAL
            nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
            'FATOR DISTRITO 98
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            
            nValorVenalPredial = 0
            
            Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                    For x = 1 To UBound(aFatorC)
                        If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                           nFatorCategoria = aFatorC(x).Fator
                           Exit For
                        End If
                    Next
                    nValorVenalPredial = nValorVenalPredial + (!AREACONSTR * nFatorCategoria)
                   .MoveNext
                Loop
            End With
            
                       
           'FATOR CATEGORIA 98
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        Else
            nValorVenalPredial = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100) '1.125
            nValorFinal = nValorIptu
            nValorITU = 0
            nAliquota = nAliquotaPredial
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
            nValorFinal = nValorITU
            nValorIptu = 0
            nAliquota = nAliquotaTerritorial
        End If
        nValorTotalIptu = nValorTotalIptu + nValorFinal 'relatorio
        nNumImovelCalc = nNumImovelCalc + 1 'relatorio
        
        '** NOVA ROTINA DE QTDE DE PARCELAS ***
        If nValorFinal > 0 And nValorFinal <= 10 Then
            nNumParc = 1
        ElseIf nValorFinal > 10 And nValorFinal <= 20 Then nNumParc = 1
        ElseIf nValorFinal > 20 And nValorFinal <= 30 Then nNumParc = 2
        ElseIf nValorFinal > 30 And nValorFinal <= 40 Then nNumParc = 3
        ElseIf nValorFinal > 40 And nValorFinal <= 50 Then nNumParc = 4
        ElseIf nValorFinal > 50 And nValorFinal <= 60 Then nNumParc = 5
        ElseIf nValorFinal > 60 And nValorFinal <= 70 Then nNumParc = 6
        ElseIf nValorFinal > 70 And nValorFinal <= 80 Then nNumParc = 7
        ElseIf nValorFinal > 80 And nValorFinal <= 90 Then nNumParc = 8
        ElseIf nValorFinal > 90 And nValorFinal <= 100 Then nNumParc = 9
        Else
            nNumParc = Val(txtNumParc.Text)
        End If
        '**************************************
        
        nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
        nValorUnica2 = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica2.Caption) / 100)), 2)
        nValorUnica3 = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica3.Caption) / 100)), 2)
        nValorParcela = Round(nValorFinal / nNumParc, 2)
        
        'GRAVA TABELA LASERIPTUESTIMADO
        Sql = "INSERT LASERIPTUESTIMADO (ANO,CODREDUZIDO,VVT,VVC,VVI,VALORPARCELA)VALUES (" & nAnoCalculo & "," & nCodReduz & "," & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nValorParcela)) & ")"
        cn.Execute Sql, rdExecDirect
        
        'GRAVA TABELA LASERIPTU
'        Sql = "INSERT LASERIPTU (ANO,CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,"
'        Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA,AREATERRENO,FATORCAT,FATORPED,FATORSIT,"
'        Sql = Sql & "FATORPRO,FATORTOP,FATORDIS,FATORGLE,AGRUPAMENTO,FRACAOIDEAL,ALIQUOTA) VALUES("
'        Sql = Sql & nAnoCalculo & "," & nCodReduz & "," & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nValorIPTU)) & "," & Virg2Ponto(CStr(nValorITU)) & ",'"
'        Sql = Sql & IIf(bTemPredial, "Predial", "Territorial") & "'," & Virg2Ponto(CStr(nAreaPrincipal)) & "," & Virg2Ponto(CStr(nTestada1)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorParcela)) & "," & Virg2Ponto(CStr(nValorUnica)) & "," & nNumParc & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorExpDocParc) * Val(txtNumParc.text)) & "," & Virg2Ponto(CStr(nValorExpDocUnica)) & "," & Virg2Ponto(CStr(nAreaTerreno)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nFatorCategoria)) & "," & Virg2Ponto(CStr(nFatorPedologia)) & "," & Virg2Ponto(CStr(nFatorSituacao)) & "," & Virg2Ponto(CStr(nFatorProfundidade)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nFatorTopografia)) & "," & Virg2Ponto(CStr(nFatorDistrito)) & "," & Virg2Ponto(CStr(nFatorGleba)) & "," & Virg2Ponto(CStr(nValorAgrupamento)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nFracaoIdeal)) & "," & Virg2Ponto(CStr(nAliquota)) & ")"
'        cn.Execute Sql, rdExecDirect
        
'        For x = 0 To nNumParc
'            If x = 0 And lblUnica.Caption = "Não" Then GoTo PROXIMO
            'GRAVA NA TABELA DEBITOPARCELA
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & 3 & "," & Format(aParc(x), "mm/dd/yyyy") & "," & Format(sDataBase, "mm/dd/yyyy") & ","
'            ax = ax & 1 & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
'            ax = ax & Null & "," & 0
'            Print #1, ax
            'GRAVA NA TABELA DEBITO TRIBUTO
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & nCodTributo & "," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & 3 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
            'GRAVA NA TABELA NUMDOCUMENTO
'            nLastDoc = nLastDoc + 1
'            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2)))
'            Print #4, ax
            'GRAVA NA TABELA PARCELADOCUMENTO
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & ","
'            ax = ax & x & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
'            Print #3, ax
'        Next
proximo:
        xId = xId + 1
       .MoveNext
    Loop
End With

Close #4
Close #3
Close #2
Close #1
MsgBox "fim"
End Sub

Private Sub CalculoInd()

Dim qd As New rdoQuery, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nValorFinal As Double, nQtdeParc As Integer

Set qd.ActiveConnection = cn

Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where (CADIMOB.CODREDUZIDO = " & Val(txtCod.Text) & ") GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblIC.Caption = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00")
    .Close
End With


qd.Sql = "{ Call spCalculo(?,?) }"
qd(0) = Val(txtCod.Text)
qd(1) = Val(txtAnoCalculo.Text)
Set RdoAux = qd.OpenResultset(rdOpenKeyset)
With RdoAux
'    txtNumParc.Text = !qtdeparc
    lblAreaTerreno.Caption = FormatNumber(!AreaTerreno, 2)
    lblAreaPrincipal.Caption = FormatNumber(!AreaPredial, 2)
    lblPredial.Caption = IIf(lblAreaPrincipal.Caption = "0,00", "Não", "Sim")
    lblTestadaMedia.Caption = FormatNumber(!TESTADAPRINC, 2)
    lblFracao.Caption = FormatNumber(!FRACAO, 2)
    lblVVT.Caption = FormatNumber(!vvt, 2)
    lblVVP.Caption = FormatNumber(!VVP, 2)
    lblVVI.Caption = FormatNumber(!VVI, 2)
    lblPerc.Caption = FormatNumber(!percisencao, 2)
    lblIPTU.Caption = IIf(!Natureza = "predial", FormatNumber(!ValorIPTU, 2), FormatNumber(!ValorITU, 2))
    lblFatorC.Caption = FormatNumber(!fcat, 2)
    lblAgrup.Caption = FormatNumber(!valorAgrupamento, 2)
    lblFatorD.Caption = FormatNumber(!fdis, 2)
    lblFatorG.Caption = FormatNumber(!fgle, 2)
    lblFatorP.Caption = FormatNumber(!fped, 2)
    lblFatorS.Caption = FormatNumber(!fsit, 2)
    lblFatorT.Caption = FormatNumber(!ftop, 2)
    lblFatorF.Caption = FormatNumber(!fpro, 2)
    lblValorFinal.Caption = FormatNumber(!VALORPARCELA * !qtdeparc, 2)
    If Not IsNull(!valorfinal) Then
        nValorFinal = !valorfinal
    Else
        nValorFinal = 0
        MsgBox "Categoria da construção não cadastrada.", vbCritical, "ERRO"
        Exit Sub
    End If

            '** NOVA ROTINA DE QTDE DE PARCELAS ***
            If nValorFinal > 0 And nValorFinal <= 10 Then
                nNumParc = 1
            ElseIf nValorFinal > 10 And nValorFinal <= 20 Then nNumParc = 1
            ElseIf nValorFinal > 20 And nValorFinal <= 30 Then nNumParc = 2
            ElseIf nValorFinal > 30 And nValorFinal <= 40 Then nNumParc = 3
            ElseIf nValorFinal > 40 And nValorFinal <= 50 Then nNumParc = 4
            ElseIf nValorFinal > 50 And nValorFinal <= 60 Then nNumParc = 5
            ElseIf nValorFinal > 60 And nValorFinal <= 70 Then nNumParc = 6
            ElseIf nValorFinal > 70 And nValorFinal <= 80 Then nNumParc = 7
            ElseIf nValorFinal > 80 And nValorFinal <= 90 Then nNumParc = 8
            ElseIf nValorFinal > 90 And nValorFinal <= 100 Then nNumParc = 9
            Else
                nNumParc = 12
            End If
            '**************************************
            
'            nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
'            nValorUnica2 = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica2.Caption) / 100)), 2)
'            nValorUnica3 = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica3.Caption) / 100)), 2)
            nValorUnica = !ValorUnica
            If Val(txtAnoCalculo.Text) >= 2018 Then
                nValorUnica2 = !ValorUnica2
                nValorUnica3 = !ValorUnica3
            Else
                nValorUnica2 = 0
                nValorUnica3 = 0
            End If
            nValorParcela = Round(nValorFinal / nNumParc, 2)
    
    
    lblUnica.Caption = FormatNumber(nValorUnica, 2)
    lblUnica2.Caption = FormatNumber(nValorUnica2, 2)
    lblUnica3.Caption = FormatNumber(nValorUnica3, 2)
    lblParcela.Caption = FormatNumber(nValorParcela, 2)
    
    If CDbl(lblParcela.Caption) > 0 Then
        If CDbl(lblParcela.Caption) < 10 Then
            MsgBox "Valor da parcela não pode ser menor R$10,00." & vbCrLf & "Diminua a qtde de parcelas para menos de " & !qtdeparc + 1, vbCritical, "Atenção"
            lblParcela.Caption = "0,00"
            Exit Sub
        End If
    Else
'        If nValorFinal > 0 Then
            MsgBox SubNull(!descisencao), vbCritical, "Atenção"
 '       End If
    End If
   .Close
End With

End Sub
