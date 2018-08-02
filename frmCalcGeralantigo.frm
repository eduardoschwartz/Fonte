VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{B60B1875-E5CA-11D2-BC3D-78A407C10000}#1.0#0"; "ksdpanel.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCalcGeral 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Execução do Cálculo Geral de ITU/IPTU"
   ClientHeight    =   6150
   ClientLeft      =   1020
   ClientTop       =   2190
   ClientWidth     =   10680
   Icon            =   "frmCalcGeral.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   10680
   Begin KSDPanel.Panel Panel5 
      Height          =   600
      Left            =   2850
      TabIndex        =   25
      Top             =   5535
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   1058
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
      TextAlign       =   0
      ForeColor       =   192
      BackColor       =   12640511
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Resultado"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   44
         Left            =   60
         TabIndex        =   116
         Top             =   60
         Width           =   1140
      End
      Begin VB.Label lblParcela 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   6600
         TabIndex        =   94
         Top             =   300
         Width           =   885
      End
      Begin VB.Label lblUnica 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   4170
         TabIndex        =   93
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label lblValorFinal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   165
         Left            =   1695
         TabIndex        =   92
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Parcela...:"
         Height          =   225
         Index           =   35
         Left            =   5430
         TabIndex        =   64
         Top             =   285
         Width           =   1200
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Única...:"
         Height          =   225
         Index           =   34
         Left            =   3135
         TabIndex        =   63
         Top             =   300
         Width           =   1065
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do ITU/IPTU..:"
         Height          =   225
         Index           =   33
         Left            =   105
         TabIndex        =   62
         Top             =   315
         Width           =   1545
      End
   End
   Begin KSDPanel.Panel Panel4 
      Height          =   4245
      Left            =   6750
      TabIndex        =   24
      Top             =   1290
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   7488
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
      TextAlign       =   0
      ForeColor       =   192
      BackColor       =   16777215
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Corrigido em 6,2% (1,062 sobre IPTU de 2004)"
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   47
         Left            =   120
         TabIndex        =   120
         Top             =   3990
         Width           =   3645
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cálculo 2004"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   43
         Left            =   60
         TabIndex        =   115
         Top             =   30
         Width           =   1140
      End
      Begin VB.Label lblIPTUCorrigido 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   105
         Top             =   3750
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor IPTU Corrigido..........:"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   38
         Left            =   120
         TabIndex        =   104
         Top             =   3750
         Width           =   1980
      End
      Begin VB.Label lblVVI 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2265
         TabIndex        =   103
         Top             =   3315
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal do Imóvel........:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   37
         Left            =   120
         TabIndex        =   102
         Top             =   3315
         Width           =   1980
      End
      Begin VB.Label lblIPTU 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   2265
         TabIndex        =   91
         Top             =   3525
         Width           =   1515
      End
      Begin VB.Label lblVVP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2265
         TabIndex        =   90
         Top             =   3105
         Width           =   1515
      End
      Begin VB.Label lblVVT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2265
         TabIndex        =   89
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Label lblAgrup 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   82
         Top             =   2175
         Width           =   1515
      End
      Begin VB.Label lblMulF 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   81
         Top             =   1935
         Width           =   1515
      End
      Begin VB.Label lblFatorG 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   80
         Top             =   1710
         Width           =   1515
      End
      Begin VB.Label lblFatorD 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   79
         Top             =   1470
         Width           =   1515
      End
      Begin VB.Label lblFatorC 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   78
         Top             =   1230
         Width           =   1515
      End
      Begin VB.Label lblFatorF 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   77
         Top             =   1005
         Width           =   1515
      End
      Begin VB.Label lblFatorS 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   76
         Top             =   765
         Width           =   1515
      End
      Begin VB.Label lblFatorP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   75
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lblFatorT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2265
         TabIndex        =   74
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do ITU/IPTU............:"
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   30
         Left            =   120
         TabIndex        =   60
         Top             =   3525
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Gleba........................:"
         Height          =   225
         Index           =   28
         Left            =   105
         TabIndex        =   58
         Top             =   1710
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Topografia................:"
         Height          =   225
         Index           =   27
         Left            =   105
         TabIndex        =   56
         Top             =   300
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Distrito.......................:"
         Height          =   225
         Index           =   26
         Left            =   105
         TabIndex        =   55
         Top             =   1470
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Categoria..................:"
         Height          =   225
         Index           =   25
         Left            =   105
         TabIndex        =   54
         Top             =   1230
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Profundidade............:"
         Height          =   225
         Index           =   24
         Left            =   105
         TabIndex        =   53
         Top             =   1005
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Situação...................:"
         Height          =   225
         Index           =   23
         Left            =   120
         TabIndex        =   52
         Top             =   765
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Pedologia.................:"
         Height          =   225
         Index           =   22
         Left            =   105
         TabIndex        =   51
         Top             =   540
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Agrupamento............:"
         Height          =   225
         Index           =   21
         Left            =   120
         TabIndex        =   44
         Top             =   2175
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal Predial.............:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   17
         Left            =   120
         TabIndex        =   40
         Top             =   3105
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal Territorial.........:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   16
         Left            =   120
         TabIndex        =   39
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplicação de  Fatores...:"
         Height          =   225
         Index           =   15
         Left            =   105
         TabIndex        =   38
         Top             =   1935
         Width           =   1980
      End
   End
   Begin KSDPanel.Panel Panel3 
      Height          =   4245
      Left            =   2850
      TabIndex        =   23
      Top             =   1290
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   7488
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
      TextAlign       =   0
      ForeColor       =   192
      BackColor       =   16777215
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Cálculo 1999"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   42
         Left            =   90
         TabIndex        =   114
         Top             =   30
         Width           =   1140
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "(IPTU99 * 1,20 * 1,7947)"
         ForeColor       =   &H00FF8080&
         Height          =   225
         Index           =   32
         Left            =   135
         TabIndex        =   106
         Top             =   3975
         Width           =   3615
      End
      Begin VB.Label lblVVI98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2250
         TabIndex        =   101
         Top             =   3315
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal do Imóvel........:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   36
         Left            =   120
         TabIndex        =   100
         Top             =   3315
         Width           =   1980
      End
      Begin VB.Label lblRedutor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   88
         Top             =   3750
         Width           =   1515
      End
      Begin VB.Label lblIPTU98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00008000&
         Height          =   180
         Left            =   2250
         TabIndex        =   87
         Top             =   3525
         Width           =   1515
      End
      Begin VB.Label lblVVP98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2250
         TabIndex        =   86
         Top             =   3105
         Width           =   1515
      End
      Begin VB.Label lblVVT98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   2250
         TabIndex        =   85
         Top             =   2880
         Width           =   1515
      End
      Begin VB.Label lblTaxaL98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   84
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label lblTaxaC98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   83
         Top             =   2400
         Width           =   1515
      End
      Begin VB.Label lblAgrup98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   73
         Top             =   2175
         Width           =   1515
      End
      Begin VB.Label lblMulF98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   72
         Top             =   1935
         Width           =   1515
      End
      Begin VB.Label lblFatorG98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   71
         Top             =   1710
         Width           =   1515
      End
      Begin VB.Label lblFatorD98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   70
         Top             =   1470
         Width           =   1515
      End
      Begin VB.Label lblFatorC98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   69
         Top             =   1230
         Width           =   1515
      End
      Begin VB.Label lblFatorF98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   68
         Top             =   1005
         Width           =   1515
      End
      Begin VB.Label lblFatorS98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   67
         Top             =   765
         Width           =   1515
      End
      Begin VB.Label lblFatorP98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   66
         Top             =   540
         Width           =   1515
      End
      Begin VB.Label lblFatorT98 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2250
         TabIndex        =   65
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do IPTU Corrigido.....:"
         Height          =   225
         Index           =   31
         Left            =   120
         TabIndex        =   61
         Top             =   3750
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor do ITU/IPTU............:"
         ForeColor       =   &H00008000&
         Height          =   225
         Index           =   29
         Left            =   120
         TabIndex        =   59
         Top             =   3525
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Gleba........................:"
         Height          =   225
         Index           =   6
         Left            =   120
         TabIndex        =   57
         Top             =   1710
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Topografia................:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   300
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Distrito......................:"
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   49
         Top             =   1470
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Categoria..................:"
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   1230
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Profundidade............:"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   47
         Top             =   1005
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Situação...................:"
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   46
         Top             =   765
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fator Pedologia.................:"
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   45
         Top             =   540
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Agrupamento............:"
         Height          =   225
         Index           =   20
         Left            =   135
         TabIndex        =   43
         Top             =   2175
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Taxa de Limpeza...............:"
         Height          =   225
         Index           =   19
         Left            =   120
         TabIndex        =   42
         Top             =   2640
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Taxa de Conservação.......:"
         Height          =   225
         Index           =   18
         Left            =   120
         TabIndex        =   41
         Top             =   2400
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal Predial............:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   14
         Left            =   120
         TabIndex        =   37
         Top             =   3105
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Venal Territorial.........:"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   13
         Left            =   120
         TabIndex        =   36
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplicação de  Fatores...:"
         Height          =   225
         Index           =   12
         Left            =   120
         TabIndex        =   35
         Top             =   1935
         Width           =   1980
      End
   End
   Begin KSDPanel.Panel Panel2 
      Height          =   1260
      Left            =   2835
      TabIndex        =   22
      Top             =   15
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   2223
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
      TextAlign       =   0
      ForeColor       =   128
      BackColor       =   12640511
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Testada Média.....:"
         Height          =   225
         Index           =   46
         Left            =   5340
         TabIndex        =   119
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label lblTestadaMedia 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   6735
         TabIndex        =   118
         Top             =   1005
         Width           =   930
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Parâmetros do Cálculo"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   45
         Left            =   90
         TabIndex        =   117
         Top             =   30
         Width           =   2040
      End
      Begin VB.Label lblRua 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1530
         TabIndex        =   113
         Top             =   510
         Width           =   6150
      End
      Begin VB.Label lblProp 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1530
         TabIndex        =   112
         Top             =   270
         Width           =   6120
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário...........:"
         Height          =   225
         Index           =   41
         Left            =   120
         TabIndex        =   111
         Top             =   270
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço.............:"
         Height          =   225
         Index           =   40
         Left            =   120
         TabIndex        =   110
         Top             =   510
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Insc.Cadastral.....:"
         Height          =   225
         Index           =   39
         Left            =   2610
         TabIndex        =   108
         Top             =   60
         Width           =   1350
      End
      Begin VB.Label lblIC 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3975
         TabIndex        =   107
         Top             =   75
         Width           =   3735
      End
      Begin VB.Label lblTestada 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   6735
         TabIndex        =   99
         Top             =   780
         Width           =   930
      End
      Begin VB.Label lblAreaPrincipal 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3975
         TabIndex        =   98
         Top             =   1005
         Width           =   1200
      End
      Begin VB.Label lblAreaTerreno 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3975
         TabIndex        =   97
         Top             =   780
         Width           =   1200
      End
      Begin VB.Label lblFracao 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1530
         TabIndex        =   96
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label lblPredial 
         BackStyle       =   0  'Transparent
         Caption         =   "Sim"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1530
         TabIndex        =   95
         Top             =   750
         Width           =   570
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Tem Predial..........:"
         Height          =   225
         Index           =   11
         Left            =   120
         TabIndex        =   34
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Área Construida...:"
         Height          =   225
         Index           =   10
         Left            =   2610
         TabIndex        =   33
         Top             =   990
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Testada Principal.:"
         Height          =   225
         Index           =   9
         Left            =   5340
         TabIndex        =   32
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Área do Terreno..:"
         Height          =   225
         Index           =   8
         Left            =   2610
         TabIndex        =   27
         Top             =   765
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fração Ideal.........:"
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   990
         Width           =   1470
      End
   End
   Begin KSDPanel.Panel Panel1 
      Height          =   6120
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2820
      _ExtentX        =   4974
      _ExtentY        =   10795
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
      BackColor       =   15658734
      Begin VB.TextBox txtAnoCalculo 
         Appearance      =   0  'Flat
         Height          =   280
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   121
         Top             =   60
         Width           =   1035
      End
      Begin VB.TextBox txtNumParc 
         Appearance      =   0  'Flat
         Height          =   280
         Left            =   1635
         TabIndex        =   6
         Top             =   810
         Width           =   675
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cálculo Individual:"
         Height          =   210
         Index           =   1
         Left            =   225
         TabIndex        =   9
         Top             =   3255
         Value           =   -1  'True
         Width           =   1590
      End
      Begin VB.OptionButton opt1 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cálculo Geral"
         Height          =   210
         Index           =   0
         Left            =   225
         TabIndex        =   8
         Top             =   2940
         Width           =   1590
      End
      Begin esMaskEdit.esMaskedEdit mskDataBase 
         Height          =   285
         Left            =   1635
         TabIndex        =   5
         Top             =   405
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
         Left            =   150
         TabIndex        =   11
         ToolTipText     =   "Sair da Tela"
         Top             =   5355
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
      Begin prjChameleon.chameleonButton cmdHelp 
         Height          =   315
         Left            =   150
         TabIndex        =   12
         ToolTipText     =   "Ajuda desta Tela"
         Top             =   4995
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Ajuda"
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
      Begin prjChameleon.chameleonButton cmdPrint 
         Height          =   315
         Left            =   150
         TabIndex        =   13
         ToolTipText     =   "Cancelar Edição"
         Top             =   4635
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
      Begin prjChameleon.chameleonButton cmdCalculo 
         Height          =   315
         Left            =   150
         TabIndex        =   14
         ToolTipText     =   "Cancelar Edição"
         Top             =   3915
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
         MICON           =   "frmCalcGeral.frx":069C
         PICN            =   "frmCalcGeral.frx":06B8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   2355
         Left            =   2355
         TabIndex        =   15
         Top             =   3375
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   4154
         _Version        =   393216
         Appearance      =   0
         Orientation     =   1
         Scrolling       =   1
      End
      Begin VB.TextBox txtCod 
         Height          =   285
         Left            =   510
         TabIndex        =   10
         Top             =   3510
         Width           =   1275
      End
      Begin prjChameleon.chameleonButton cmdGravar 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   150
         TabIndex        =   109
         ToolTipText     =   "Cancelar Edição"
         Top             =   4275
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
         MICON           =   "frmCalcGeral.frx":0757
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "3 %"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1665
         TabIndex        =   31
         Top             =   2460
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Aliquota Territorial..:"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   30
         Top             =   2460
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "1,5 %"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1665
         TabIndex        =   29
         Top             =   2175
         Width           =   810
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Aliquota Predial......:"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   28
         Top             =   2190
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0 %"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   2340
         TabIndex        =   21
         Top             =   5760
         Width           =   270
      End
      Begin VB.Label lblPB 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2235
         TabIndex        =   20
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Desconto Única %.:"
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   1425
         Width           =   1455
      End
      Begin VB.Label lblPercUnica 
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1665
         TabIndex        =   18
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcela Única........:"
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label lblTemUnica 
         BackStyle       =   0  'Transparent
         Caption         =   "Sim"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1665
         TabIndex        =   16
         Top             =   1140
         Width           =   570
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Base.............:"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   75
         TabIndex        =   7
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Parcelas......:"
         Height          =   225
         Index           =   6
         Left            =   105
         TabIndex        =   4
         Top             =   855
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Imóveis Estimados.:"
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label lblEstimado 
         BackStyle       =   0  'Transparent
         Caption         =   "14350"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1665
         TabIndex        =   2
         Top             =   1920
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano de Cálculo.....:"
         Height          =   225
         Left            =   105
         TabIndex        =   1
         Top             =   135
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmCalcGeral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim nValorAgrupamento98 As Double
Dim nNumTestadas As Integer
Dim nTestadaPrincipal As Double
Dim nCodGleba As Integer
Dim nFatorGleba As Double
Dim nFatorGleba98 As Double
Dim nCodProfundidade As Integer
Dim nValorProfundidade As Double
Dim nFatorProfundidade As Double
Dim nFatorProfundidade98 As Double
Dim nCodSituacao As Integer
Dim nFatorSituacao As Double
Dim nFatorSituacao98 As Double
Dim nCodPedologia As Integer
Dim nFatorPedologia As Double
Dim nFatorPedologia98 As Double
Dim nCodTopografia As Integer
Dim nFatorTopografia As Double
Dim nFatorTopografia98 As Double
Dim nFatorDistrito As Double
Dim nFatorDistrito98 As Double
Dim nValorFatores As Double
Dim nValorFatores98 As Double
Dim nFatorCategoria As Double
Dim nFatorCategoria98 As Double
Dim nValorVenalTerritorial As Double
Dim nValorVenalTerritorial98 As Double
Dim nValorVenalPredial As Double
Dim nValorVenalPredial98 As Double
Dim nCodTributo As Integer
Dim nValorVenalImovel As Double
Dim nValorVenalImovel98 As Double
Dim nValorIPTU As Double, nValorITU As Double
Dim nValorIPTU98 As Double, nValorITU98 As Double
Dim nTaxaLimpeza As Double, nTaxaConservacao As Double
Dim nValorFinal As Double
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

Private Sub cmdCalculo_Click()

Limpa

CarregaImovel

If Val(txtNumParc.text) = 0 Then
    MsgBox "Digite a qtde de parcelas.", vbExclamation, "Atenção"
    Exit Sub
End If

If opt1(0).Value = True Then
'    MsgBox "Você não tem permissão para realizar o cálculo geral de IPTU.", vbExclamation, "ALERTA DE SEGURANÇA !!"
    
'    Exit Sub
End If

'CARREGA PARAMETROS
nUfir1999 = RetornaUFIR(1999)
nUfirCalc = RetornaUFIR(nAnoCalculo)
nAliquotaPredial = 1.5
nAliquotaTerritorial = 3
bExec = True
If opt1(1).Value = True Then
    If Val(txtCod.text) = 0 Then
       MsgBox "Digite o código do imóvel.", vbExclamation, "Atenção"
    Else
       Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCod.text)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
            If .RowCount = 0 Then
                MsgBox "Imóvel não cadastrado.", vbExclamation, "Atenção"
            Else
                CalculoIndividual (Val(txtCod.text))
            End If
       End With
    End If
Else
    If frmMdi.frTeste.Visible = False Then
        MsgBox "Calculo geral apenas para base de testes."
        Exit Sub
    End If
    If Not IsDate(mskDataBase.text) Then
        MsgBox "Data Base Inválida.", vbExclamation, "atenção"
        Exit Sub
    End If
    
    If MsgBox("Executar o cálculo de IPTU para todos os imóveis?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
    Ocupado
    CalculoGeral
    Liberado
    If bExec Then
       MsgBox "Calculo efetuado", vbExclamation, "atenção"
    End If
End If

End Sub

Private Sub CalculoIndividual(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

bIsento = False

Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Val(txtAnoCalculo.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Este imóvel esta classificado como: " & !DESCTIPO, vbExclamation, "Atenção"
'        Exit Sub
        bIsento = True
    End If
   .Close
End With

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
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
                'nTestadaPrincipal = nSomaTestada / nNumTestadas
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
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) And RdoAux2!QTDEPAV < 2 And nAreaTerreno < 600 Then
                        bIsento = True
                        MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                        Limpa
                    End If
                End If
               .Close
            End With
        Else
            bIsento = False
        End If
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
             Select Case !USOCONSTR
                  Case 0
                     nTaxaLimpeza = 3.78
                  Case 1, 2, 3, 4, 5
                     nTaxaLimpeza = 10.57
                  Case Else
                     nTaxaLimpeza = 3.01
             End Select
        Else
             nTaxaLimpeza = 3.01
        End If
        nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
       '--CÁLCULO DA TAXA DE CONSERVAÇÃO
        If RdoAux!PAVIMENTO = 1 Then
           nTaxaConservacao = 1.35 * nTestadaPrincipal
        Else
           nTaxaConservacao = 0
        End If
        If nCodBairro = 81 Then
           lblTaxaL98.Caption = FormatNumber(0, 2)
           lblTaxaC98.Caption = FormatNumber(0, 2)
           nTaxaLimpeza = 1
           nTaxaConservacao = 1
        Else
           lblTaxaL98.Caption = FormatNumber(nTaxaLimpeza, 2)
        End If
        lblTaxaC98.Caption = FormatNumber(nTaxaConservacao, 2)
       .Close
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
       nValorAgrupamento98 = aFatorR98(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
       nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
    End If
    
    lblAgrup.Caption = FormatNumber(nValorAgrupamento, 2)
'    nValorAgrupamento98 = nValorAgrupamento
'    lblAgrup98.Caption = lblAgrup.Caption
    lblAgrup98.Caption = FormatNumber(nValorAgrupamento98, 2)
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
'    If !Dt_CodUsoTerreno = 6 Then
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
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
        nFatorGleba98 = aFatorG98(nCodGleba)
'    Else
'        nFatorGleba = 1
'        nFatorGleba98 = 1
'    End If
    lblFatorG98.Caption = FormatNumber(nFatorGleba98, 2)
    lblFatorG.Caption = FormatNumber(nFatorGleba, 2)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         'nValorProfundidade = FormatNumber(nAreaTerreno / nTestadaPrincipal, 2)
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= aProf(x).Min And nValorProfundidade <= aProf(x).Max Then
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
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
        nFatorProfundidade98 = 0
        For x = 1 To UBound(aFatorF98)
            If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
               nFatorProfundidade98 = aFatorF98(x).Fator
               Exit For
            End If
        Next
        lblFatorF98.Caption = FormatNumber(nFatorProfundidade98, 2)
     Else
        nFatorProfundidade = 1
        nFatorProfundidade98 = 1
        lblFatorF.Caption = "1,00"
        lblFatorF98.Caption = "1,00"
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    lblFatorS.Caption = FormatNumber(nFatorSituacao, 2)
    'FATOR SITUACAO 98
    nFatorSituacao98 = aFatorS98(nCodSituacao)
    lblFatorS98.Caption = FormatNumber(nFatorSituacao98, 2)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    lblFatorP.Caption = FormatNumber(nFatorPedologia, 2)
    'FATOR PEDOLOGIA 98
    nFatorPedologia98 = aFatorP98(nCodPedologia)
    lblFatorP98.Caption = FormatNumber(nFatorPedologia98, 2)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    lblFatorT.Caption = FormatNumber(nFatorTopografia, 2)
    'FATOR TOPOGRAFIA 98
    nFatorTopografia98 = aFatorT98(nCodTopografia)
    lblFatorT98.Caption = FormatNumber(nFatorTopografia98, 2)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
    nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
    lblMulF.Caption = FormatNumber(nValorFatores, 2)
    lblMulF98.Caption = FormatNumber(nValorFatores98, 2)
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores
    nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
    lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2)
    lblVVT98.Caption = FormatNumber(nValorVenalTerritorial98, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        lblFatorD.Caption = FormatNumber(nFatorDistrito, 2)
        'FATOR DISTRITO 98
        nFatorDistrito98 = aFatorD98(!Distrito)
        lblFatorD98.Caption = FormatNumber(nFatorDistrito98, 2)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        nValorVenalPredial98 = 0
        nFatorCategoria = 0
        For x = 1 To UBound(aFatorC)
            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
               nFatorCategoria = aFatorC(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
        lblFatorC.Caption = FormatNumber(nFatorCategoria, 2)
        
       'FATOR CATEGORIA 98
        nFatorCategoria98 = 0
        For x = 1 To UBound(aFatorC98)
            If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
               nFatorCategoria98 = aFatorC98(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
        lblFatorC98.Caption = FormatNumber(nFatorCategoria98, 2)
        
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
        lblVVP.Caption = FormatNumber(nValorVenalPredial, 2)
        lblVVP98.Caption = FormatNumber(nValorVenalPredial98, 2)
    Else
        nFatorDistrito = 0
        nFatorDistrito98 = 0
        nFatorCategoria = 0
        nFatorCategoria98 = 0
        lblFatorD98.Caption = FormatNumber(nFatorDistrito98, 2)
        lblFatorD.Caption = FormatNumber(nFatorDistrito, 2)
        lblFatorC98.Caption = FormatNumber(nFatorCategoria98, 2)
        lblFatorC.Caption = FormatNumber(nFatorCategoria, 2)
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
        nValorIPTU = nValorVenalImovel * (nAliquotaPredial / 100) * 1.062 'reajuste 2004-2005
        nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
        nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
        lblIPTU.Caption = FormatNumber(nValorVenalImovel * (nAliquotaPredial / 100), 2)
 '       nValorIPTU = nValorIPTU * 1.3916
        lblIPTUCorrigido.Caption = FormatNumber(nValorIPTU, 2)
        lblIPTU98.Caption = FormatNumber(nValorIPTU98, 2)
        nValorIPTU98 = CDbl(lblIPTU98.Caption) * 1.7947
'       nValorIPTU98 = CDbl(lblIPTU98.Caption) * 1.6916
        lblRedutor.Caption = FormatNumber(nValorIPTU98, 2)
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorVenalImovel98 = nValorVenalTerritorial98
        nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100) * 1.062 'reajuste 2004-2005
        nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
        nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
        lblIPTU.Caption = FormatNumber(nValorVenalImovel * (nAliquotaTerritorial / 100), 2)
 '       nValorITU = nValorITU * 1.3916
        lblIPTUCorrigido.Caption = FormatNumber(nValorITU, 2)
        lblIPTU98.Caption = FormatNumber(nValorITU98, 2)
        nValorITU98 = CDbl(lblIPTU98.Caption) * 1.7947
'        nValorITU98 = CDbl(lblIPTU98.Caption) * 1.6916
        lblRedutor.Caption = FormatNumber(nValorITU98, 2)
    End If
    lblVVI.Caption = FormatNumber(nValorVenalImovel, 2)
    lblVVI98.Caption = FormatNumber(nValorVenalImovel98, 2)
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    If bTemPredial Then
        If nValorIPTU98 > nValorIPTU Then
           nValorFinal = nValorIPTU
        Else
           nValorFinal = nValorIPTU98
        End If
    Else
        If nValorITU98 > nValorITU Then
           nValorFinal = nValorITU
        Else
           nValorFinal = nValorITU98
        End If
    End If
    If bIsento Then
        lblValorFinal.Caption = FormatNumber(0, 2)
        lblUnica.Caption = FormatNumber(0, 2)
        lblParcela.Caption = FormatNumber(0, 2)
    Else
        lblValorFinal.Caption = FormatNumber(nValorFinal, 2)
        lblUnica.Caption = FormatNumber(nValorFinal - (nValorFinal * CDbl(lblPercUnica.Caption) / 100), 2)
        lblParcela.Caption = FormatNumber(nValorFinal / CDbl(txtNumParc.text), 2)
    End If
End With

End Sub

Private Sub CalculoGeral()
Dim xId As Long, nNumRec As Long
Dim nValorExpDocParc As Double, nValorExpDocUnica As Double, nLastDoc As Long, nAreaTerrenoReal As Double
Dim ax As String, sDataBase As String
Dim nValorUnica As Double, nValorParcela As Double, nTestada1 As Double
'GoTo IMPORTA
'Exit Sub
'lblWait.Caption = "APAGANDO TABELAS. AGUARDE........"
'lblWait.Refresh
cn.QueryTimeout = 0
'TABELA LASERIPTU
'lblWait.Caption = "APAGANDO TABELA LASER........"
'lblWait.Refresh
Sql = "TRUNCATE TABLE LASERIPTU"
cn.Execute Sql, rdExecDirect
DoEvents

'GoTo Calculo

Sql = "SELECT COUNT(CODREDUZIDO) AS CONTADOR FROM DEBITOPARCELA WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If !CONTADOR > 0 Then
        
        'REMOVE RELACIONAMENTO
        Sql = "BEGIN TRANSACTION SET QUOTED_IDENTIFIER ON "
  '       cn.Execute Sql, rdExecDirect
        Sql = "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE COMMIT"
  '      cn.Execute Sql, rdExecDirect
        Sql = "BEGIN TRANSACTION ALTER TABLE dbo.PARCELADOCUMENTO  DROP CONSTRAINT FK_PARCELADOCUMENTO_NUMDOCUMENTO COMMIT"
  '      cn.Execute Sql, rdExecDirect
        
        'TABELA NUMDOCUMENTO
'        lblWait.Caption = "APAGANDO DOCUMENTOS AGUARDE..."
'        lblWait.Refresh
        Sql = "DELETE FROM NUMDOCUMENTO WHERE NUMDOCUMENTO in ("
        Sql = Sql & "SELECT NumDocumento From PARCELADOCUMENTO WHERE ANOEXERCICIO =" & nAnoCalculo & " AND CODLANCAMENTO = 1)"
'        cn.Execute Sql, rdExecDirect
        DoEvents
        
        'RECRIA O RELACIONAMENTO
        Sql = "BEGIN TRANSACTION SET QUOTED_IDENTIFIER ON"
'        cn.Execute Sql, rdExecDirect
        Sql = "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE COMMIT"
'        cn.Execute Sql, rdExecDirect
        Sql = "BEGIN TRANSACTION ALTER TABLE dbo.PARCELADOCUMENTO WITH NOCHECK ADD CONSTRAINT FK_PARCELADOCUMENTO_NUMDOCUMENTO FOREIGN KEY "
        Sql = Sql & " ( NUMDOCUMENTO) REFERENCES dbo.NUMDOCUMENTO ( NUMDOCUMENTO ) COMMIT"
'        cn.Execute Sql, rdExecDirect
        
        'TABELA PARCELA DOCUMENTO
'        lblWait.Caption = "APAGANDO LIGAÇÕES AGUARDE........"
'        lblWait.Refresh
'        Sql = "DELETE FROM PARCELADOCUMENTO WHERE NUMDOCUMENTO in ("
'        Sql = Sql & "SELECT NumDocumento From PARCELADOCUMENTO WHERE ANOEXERCICIO =" & nAnoCalculo & " AND CODLANCAMENTO = 1)"
        Sql = "DELETE FROM PARCELADOCUMENTO WHERE ANOEXERCICIO =" & nAnoCalculo & " AND CODLANCAMENTO = 1"
'        cn.Execute Sql, rdExecDirect
        DoEvents
        'TABELA DEBITOTRIBUTO
        'lblWait.Caption = "APAGANDO TRIBUTOS AGUARDE......."
        'lblWait.Refresh
        Sql = "DELETE FROM DEBITOTRIBUTO WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
'        cn.Execute Sql, rdExecDirect
        DoEvents
        'TABELA DEBITOPAGO
        'lblWait.Caption = "APAGANDO DEBITOPAGO AGUARDE......."
        'lblWait.Refresh
        Sql = "DELETE FROM DEBITOPAGO WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
'        cn.Execute Sql, rdExecDirect
        'TABELA DEBITOPARCELA
        'lblWait.Caption = "APAGANDO PARCELAS AGUARDE......."
        'lblWait.Refresh
        Sql = "DELETE FROM DEBITOPARCELA WHERE ANOEXERCICIO = " & nAnoCalculo & " AND CODLANCAMENTO = 1"
'        cn.Execute Sql, rdExecDirect
    End If
End With
TESTE:
DoEvents
'lblWait.Caption = "EFETUANDO CÁLCULO GERAL. AGUARDE......."
'lblWait.Refresh

sDataBase = mskDataBase.text

Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTO.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTO.TXT" For Output As #3
Open sPathBin & "\NUMDOCUMENTO.TXT" For Output As #4
Open sPathBin & "\DIFERENCA.TXT" For Output As #5

'********************************
' TAXA DE EXPEDIÇÃO DE DOCUMENTO
'********************************
Calculo:
Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
        nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
        nValorExpDocUnica = FormatNumber(!VALORUNICA, 2)
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
    nLastDoc = !ULTIMO + 100
   .Close
End With


'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,CADIMOB.INATIVO,LI_CODBAIRRO,PAVIMENTO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,"
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE WHERE CADIMOB.CODREDUZIDO=123 GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
'Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
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
        If !Inativo = True Then GoTo PROXIMO
        If Not bExec Then
           MsgBox "Cálculo Interrompido pelo usuário", vbCritical, "Atenção"
           Exit Do
        End If
        'DADOS DO IMOVEL
        nCodReduz = !CODREDUZIDO
'        If nCodReduz = 203 Then MsgBox "IMOVEL 203"
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Val(txtAnoCalculo.text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                GoTo PROXIMO
'                Exit Sub
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
        bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
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
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
        'TEM ÁREA?
            If .RowCount > 0 Then
                If Not IsNull(RdoAux!SOMAAREA) Then
                    If RdoAux!SOMAAREA <= 65 And !USOCONSTR = 0 And (!CATCONSTR = 4 Or !CATCONSTR = 7) And !QTDEPAV < 2 And nAreaTerreno < 600 Then
                        GoTo PROXIMO
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
            Else
                bTemPredial = False
                nAreaPrincipal = 0
            End If
           'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
            If bTemPredial Then
                 If .RowCount > 0 Then
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                 
                 Select Case !USOCONSTR
                      Case 0
                         nTaxaLimpeza = 3.78
                      Case 1, 2, 3, 4, 5
                         nTaxaLimpeza = 10.57
                      Case Else
                         nTaxaLimpeza = 3.01
                 End Select
                 Else
                    nTaxaLimpeza = 3.01
                 End If
            Else
                 nTaxaLimpeza = 3.01
            End If
            nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
            If nCodBairro = 81 Then
               nTaxaLimpeza = 0
               nTaxaConservacao = 0
            End If
           '--CÁLCULO DA TAXA DE CONSERVAÇÃO
            If RdoAux!PAVIMENTO = 1 Then
               nTaxaConservacao = 1.35 * nTestadaPrincipal
            Else
               nTaxaConservacao = 0
            End If
            If nCodBairro = 81 Then
'               lblTaxaL98.Caption = FormatNumber(0, 2)
'               lblTaxaC98.Caption = FormatNumber(0, 2)
               nTaxaLimpeza = 1
               nTaxaConservacao = 1
'            Else
'               lblTaxaL98.Caption = FormatNumber(nTaxaLimpeza, 2)
            End If
'            nTaxaConservacao = 1.35 * nTestadaPrincipal
           .Close
        End With
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
           nValorAgrupamento98 = aFatorR98(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
           nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
        End If
'        nValorAgrupamento = aFatorR(nCodAgrupamento)
 '       nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
'        If !Dt_CodUsoTerreno = 6 Then
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
            'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
            nFatorGleba98 = aFatorG98(nCodGleba)
'        Else
'            nFatorGleba = 1
'            nFatorGleba98 = 1
'        End If
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
               nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestadaPrincipal, 2)
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
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
            nFatorProfundidade98 = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade98 = aFatorF98(x).Fator
                   Exit For
                End If
            Next
        Else
            nFatorProfundidade = 1
            nFatorProfundidade98 = 1
        End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        'FATOR SITUACAO 98
        nFatorSituacao98 = aFatorS98(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        'FATOR PEDOLOGIA 98
        nFatorPedologia98 = aFatorP98(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        'FATOR TOPOGRAFIA 98
        nFatorTopografia98 = aFatorT98(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
        nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores
        nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
            nFatorDistrito = aFatorD(!Distrito)
            'FATOR DISTRITO 98
            nFatorDistrito98 = aFatorD98(!Distrito)
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            nValorVenalPredial = 0
            nValorVenalPredial98 = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
           'FATOR CATEGORIA 98
            nFatorCategoria98 = 0
            For x = 1 To UBound(aFatorC98)
                If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
                   nFatorCategoria98 = aFatorC98(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
            nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
'            If !Distrito > 1 Then
'                nValorVenalPredial = nValorVenalPredial * 0.6
'                nValorVenalPredial98 = nValorVenalPredial98 * 0.6
'            End If
        Else
            nValorVenalPredial = 0
            nValorVenalPredial98 = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
            nValorIPTU = nValorVenalImovel * (nAliquotaPredial / 100) * 1.062
'            nValorIPTU = nValorIPTU * 1.3916
            nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
            nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
            'nValorIPTU98 = nValorIPTU98 * 1.6916
            nValorIPTU98 = nValorIPTU98 * 1.7947
            nValorITU = 0
            nValorITU98 = 0
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorVenalImovel98 = nValorVenalTerritorial98
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100) * 1.062
'            nValorITU = nValorITU * 1.3916
            nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
            nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
            'nValorITU98 = nValorITU98 * 1.6916
            nValorITU98 = nValorITU98 * 1.7947
            'nValorIPTU98 = CDbl(lblIPTU98.Caption) * 1.7947
            nValorIPTU = 0
            nValorIPTU98 = 0
        End If
        'COMPARAÇÃO ENTRE OS CÁLCULOS
        If bTemPredial Then
            If nValorIPTU98 > nValorIPTU Then
               nValorFinal = nValorIPTU
            Else
               nValorFinal = nValorIPTU98
            
               ax = nCodReduz & "," & Virg2Ponto(Format(nValorIPTU98, "#0.00")) & "," & Virg2Ponto(Format(nValorIPTU, "#0.00"))
               Print #5, ax
            
            End If
            nValorIPTU = nValorFinal
        Else
            If nValorITU98 > nValorITU Then
               nValorFinal = nValorITU
            Else
               nValorFinal = nValorITU98
               
               ax = nCodReduz & "," & Virg2Ponto(Format(nValorITU98, "#0.00")) & "," & Virg2Ponto(Format(nValorITU, "#0.00"))
               Print #5, ax
            End If
            nValorITU = nValorFinal
        End If
        nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
        nValorParcela = Round(nValorFinal / Val(txtNumParc.text), 2)
    'GoTo FIM
        'GRAVA TABELA LASERIPTU
        Sql = "INSERT LASERIPTU (CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,"
        Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA) VALUES("
        Sql = Sql & nCodReduz & "," & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nValorIPTU)) & "," & Virg2Ponto(CStr(nValorITU)) & ",'"
        Sql = Sql & IIf(bTemPredial, "Predial", "Territorial") & "'," & Virg2Ponto(CStr(nAreaPrincipal)) & "," & Virg2Ponto(CStr(nTestada1)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorParcela)) & "," & Virg2Ponto(CStr(nValorUnica)) & "," & Val(txtNumParc.text) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorExpDocParc) * Val(txtNumParc.text)) & "," & Virg2Ponto(CStr(nValorExpDocUnica)) & ")"
        cn.Execute Sql, rdExecDirect
'#APAGAR
 '       GoTo PROXIMO
        
        For x = 0 To Val(txtNumParc.text)
            If x = 0 And lblUnica.Caption = "Não" Then GoTo PROXIMO
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
            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
            ax = ax & 3 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ","
            ax = ax & 0 & "," & 0 & "," & 0
            Print #2, ax
            'GRAVA NA TABELA NUMDOCUMENTO
            nLastDoc = nLastDoc + 1
            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2)))
            Print #4, ax
            'GRAVA NA TABELA PARCELADOCUMENTO
            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & ","
            ax = ax & x & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
            Print #3, ax
        Next
PROXIMO:
        xId = xId + 1
       .MoveNext
    Loop
End With
'#APAGAR

Close #5
Close #4
Close #3
Close #2
Close #1
Exit Sub
IMPORTA:
'lblWait.Caption = "IMPORTANDO DADOS. AGUARDE......."
'lblWait.Refresh


IMPORT:
Dim Terminaram As Boolean, índice As Integer, RetProc As Long
Dim StartupInfo As TSTARTUPINFO
DoEvents
sPathBin = "D:\CADASTRO\BIN"
CreateProcess "BCP.EXE", "Nspn TRIBUTACAOTESTE..DEBITOPARCELA in " & sPathBin & "\DEBITOPARCELA.TXT /f" & sPathBin & "\DEBITOPARCELA.FMT" & " /SDBSERVER /Uschwartz /Pcvrcs04 -eD:\ERRO.TXT", 0, 0, 0, CREATE_NEW_CONSOLE + NORMAL_PRIORITY_CLASS, 0, sPathBin, StartupInfo, ProcessInformation(1)

Do
  Terminaram = True
  For índice = 1 To 1
     RetProc = WaitForSingleObject(ProcessInformation(índice).hProcess, 300)
     If RetProc = STATUS_TIMEOUT Then
        Terminaram = False
     End If
  Next
  Refresh
Loop Until Terminaram

Me.Refresh
DoEvents

CreateProcess "D:\Arquivos de programas\Microsoft SQL Server\MSSQL\Binn\BCP.EXE", "Nspn TRIBUTACAOTESTE..DEBITOTRIBUTO in " & sPathBin & "\DEBITOTRIBUTO.TXT /f" & sPathBin & "\DEBITOTRIBUTO.FMT" & " /SDBSERVER /Uschwartz /Pcvrcs04 -eD:\ERRO.TXT", 0, 0, 0, CREATE_NEW_CONSOLE + NORMAL_PRIORITY_CLASS, 0, App.Path, StartupInfo, ProcessInformation(1)

Do
  Terminaram = True
  For índice = 1 To 1
     RetProc = WaitForSingleObject(ProcessInformation(índice).hProcess, 300)
     If RetProc = STATUS_TIMEOUT Then
        Terminaram = False
     End If
  Next
  Refresh
Loop Until Terminaram
'PROC:
Me.Refresh
DoEvents

CreateProcess "D:\Arquivos de programas\Microsoft SQL Server\MSSQL\Binn\BCP.EXE", "Nspn TRIBUTACAOTESTE..PARCELADOCUMENTO in " & sPathBin & "\PARCELADOCUMENTO.TXT /f" & sPathBin & "\PARCELADOCUMENTO.FMT" & " /SDBSERVER /Uschwartz /Pcvrcs04 -eD:\ERRO.TXT", 0, 0, 0, CREATE_NEW_CONSOLE + NORMAL_PRIORITY_CLASS, 0, App.Path, StartupInfo, ProcessInformation(1)

Do
  Terminaram = True
  For índice = 1 To 1
     RetProc = WaitForSingleObject(ProcessInformation(índice).hProcess, 300)
     If RetProc = STATUS_TIMEOUT Then
        Terminaram = False
     End If
  Next
  Refresh
Loop Until Terminaram
'DOC:
Me.Refresh
DoEvents

CreateProcess "D:\Arquivos de programas\Microsoft SQL Server\MSSQL\Binn\BCP.EXE", "Nspn TRIBUTACAOTESTE..NUMDOCUMENTO in " & sPathBin & "\NUMDOCUMENTO.TXT /f" & sPathBin & "\NUMDOCUMENTO.FMT" & " /SDBSERVER /Uschwartz /Pcvrcs04 -eD:\ERRO.TXT", 0, 0, 0, CREATE_NEW_CONSOLE + NORMAL_PRIORITY_CLASS, 0, App.Path, StartupInfo, ProcessInformation(1)

Do
  Terminaram = True
  For índice = 1 To 1
     RetProc = WaitForSingleObject(ProcessInformation(índice).hProcess, 300)
     If RetProc = STATUS_TIMEOUT Then
        Terminaram = False
     End If
  Next
  Refresh
Loop Until Terminaram
fim:
'lblWait.Caption = "CALCULO EFETUADO COM SUCESSO."
'lblWait.Refresh

End Sub

Private Sub CalculoGeral2()
Dim xId As Long, nNumRec As Long
Dim nValorExpDocParc As Double, nValorExpDocUnica As Double, nLastDoc As Long, nAreaTerrenoReal As Double
Dim ax As String
Dim nValorUnica As Double, nValorParcela As Double, nTestada1 As Double, nValorIptuNovo As Double

DoEvents
Exit Sub
Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1

'********************************
' TAXA DE EXPEDIÇÃO DE DOCUMENTO
'********************************
Calculo:
Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
     nValorExpDocUnica = FormatNumber(!VALORUNICA, 2)
    .Close
End With
'ULTIMO Nº DE DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nLastDoc = !ULTIMO + 100
   .Close
End With

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO,PAVIMENTO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,"
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
'Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE WHERE CADIMOB.CODREDUZIDO=19457 GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
Sql = Sql & "ORDER BY CADIMOB.CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1
    nNumRec = .RowCount
    Do Until .EOF
        'GAUGE
        If xId Mod 100 = 0 Then
           CallPb xId, nNumRec
        End If
        If Not bExec Then
           MsgBox "Cálculo Interrompido pelo usuário", vbCritical, "Atenção"
           Exit Do
        End If
        'DADOS DO IMOVEL
        nCodReduz = !CODREDUZIDO
'        If nCodReduz = 5531 Then MsgBox "IMOVEL 5531"
        Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
        Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Val(txtAnoCalculo.text)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                GoTo PROXIMO
'                Exit Sub
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
        bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
        If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
        'TESTADAS
        Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumTestadas = .RowCount
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
        End With
        'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
        '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
        
        'BUSCA ÁREA PRINCIPAL
        Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
        'TEM ÁREA?
            If .RowCount > 0 Then
                If Not IsNull(RdoAux!SOMAAREA) Then
                    If RdoAux!SOMAAREA <= 65 And !USOCONSTR = 0 And (!CATCONSTR = 4 Or !CATCONSTR = 7) Then
                        GoTo PROXIMO
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
            Else
                bTemPredial = False
                nAreaPrincipal = 0
            End If
           'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
            If bTemPredial Then
                 If .RowCount > 0 Then
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                 
                 Select Case !USOCONSTR
                      Case 0
                         nTaxaLimpeza = 3.78
                      Case 1, 2, 3, 4, 5
                         nTaxaLimpeza = 10.57
                      Case Else
                         nTaxaLimpeza = 3.01
                 End Select
                 Else
                    nTaxaLimpeza = 3.01
                 End If
            Else
                 nTaxaLimpeza = 3.01
            End If
            nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
            If nCodBairro = 81 Then
               nTaxaLimpeza = 0
               nTaxaConservacao = 0
            End If
           '--CÁLCULO DA TAXA DE CONSERVAÇÃO
            If RdoAux!PAVIMENTO = 1 Then
               nTaxaConservacao = 1.35 * nTestadaPrincipal
            Else
               nTaxaConservacao = 0
            End If
            If nCodBairro = 81 Then
'               lblTaxaL98.Caption = FormatNumber(0, 2)
'               lblTaxaC98.Caption = FormatNumber(0, 2)
               nTaxaLimpeza = 1
               nTaxaConservacao = 1
'            Else
'               lblTaxaL98.Caption = FormatNumber(nTaxaLimpeza, 2)
            End If
'            nTaxaConservacao = 1.35 * nTestadaPrincipal
           .Close
        End With
        'VALOR DOS AGRUPAMENTOS
        If !Dt_CodUsoTerreno = 6 Then
           nValorAgrupamento = aFatorR(7)
           nValorAgrupamento98 = aFatorR98(7)
        Else
           nValorAgrupamento = aFatorR(nCodAgrupamento)
           nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
        End If
'        nValorAgrupamento = aFatorR(nCodAgrupamento)
 '       nValorAgrupamento98 = aFatorR98(nCodAgrupamento)
        '**************************
        'CÁLCULO DOS FATORES
        '**************************
        '**************************
        '### FATOR GLEBA ###
        '**************************
        If !Dt_CodUsoTerreno = 6 Then
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
            'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
            nFatorGleba98 = aFatorG98(nCodGleba)
        Else
            nFatorGleba = 1
            nFatorGleba98 = 1
        End If
        '**************************
        '### FATOR PROFUNDIDADE ###
        '**************************
        If !Dt_CodUsoTerreno <> 6 Then
            '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
            If nTestadaPrincipal > 0 Then
               nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestadaPrincipal, 2)
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
            'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
            nFatorProfundidade98 = 0
            For x = 1 To UBound(aFatorF)
                If aFatorF98(x).Distrito = !Distrito And aFatorF98(x).Codigo = nCodProfundidade Then
                   nFatorProfundidade98 = aFatorF98(x).Fator
                   Exit For
                End If
            Next
        Else
            nFatorProfundidade = 1
            nFatorProfundidade98 = 1
        End If
        '**************************
        '### FATOR SITUAÇÃO ###
        '**************************
        nFatorSituacao = aFatorS(nCodSituacao)
        'FATOR SITUACAO 98
        nFatorSituacao98 = aFatorS98(nCodSituacao)
        '**************************
        '### FATOR PEDOLOGIA ###
        '**************************
        nFatorPedologia = aFatorP(nCodPedologia)
        'FATOR PEDOLOGIA 98
        nFatorPedologia98 = aFatorP98(nCodPedologia)
        '**************************
        '### FATOR TOPOGRAFIA ###
        '**************************
        nFatorTopografia = aFatorT(nCodTopografia)
        'FATOR TOPOGRAFIA 98
        nFatorTopografia98 = aFatorT98(nCodTopografia)
        '**************************
        'FIM DO CÁLCULO DOS FATORES
        '**************************
        'MULTIPLICA OS FATORES
        nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
        nValorFatores98 = nFatorTopografia98 * nFatorSituacao98 * nFatorPedologia98 * nFatorProfundidade98 * nFatorGleba98
        'CÁLCULO VALOR VENAL TERRITORIAL
        nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * nValorFatores
        nValorVenalTerritorial98 = nAreaTerreno * nValorAgrupamento98 * nValorFatores98
        'CÁLCULO VALOR VENAL PREDIAL
        '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
        If bTemPredial Then
            '**************************
            '### FATOR DISTRITO ###
            '**************************
            nFatorDistrito = aFatorD(!Distrito)
            'FATOR DISTRITO 98
            nFatorDistrito98 = aFatorD98(!Distrito)
            '**************************
            '### FATOR CATEGORIA ###
            '**************************
            nValorVenalPredial = 0
            nValorVenalPredial98 = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
           'FATOR CATEGORIA 98
            nFatorCategoria98 = 0
            For x = 1 To UBound(aFatorC98)
                If aFatorC98(x).Uso = nUso And aFatorC98(x).Tipo = nTipo And aFatorC98(x).Categoria = nCat Then
                   nFatorCategoria98 = aFatorC98(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial98 = nValorVenalPredial98 + (nAreaPrincipal * nFatorCategoria98)
            nValorVenalPredial = nValorVenalPredial * nFatorDistrito
            nValorVenalPredial98 = nValorVenalPredial98 * nFatorDistrito98
'            If !Distrito > 1 Then
'                nValorVenalPredial = nValorVenalPredial * 0.6
'                nValorVenalPredial98 = nValorVenalPredial98 * 0.6
'            End If
        Else
            nValorVenalPredial = 0
            nValorVenalPredial98 = 0
        End If
        'VALOR ITU/IPTU
        If bTemPredial Then
            nCodTributo = 1
            nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
            nValorVenalImovel98 = nValorVenalTerritorial98 + nValorVenalPredial98
            nValorIPTU = nValorVenalImovel * (nAliquotaPredial / 100)
'            nValorIPTU = nValorIPTU * 1.3916
            nValorIPTU98 = nValorVenalImovel98 * (nAliquotaPredial / 100)
            nValorIPTU98 = nValorIPTU98 + nTaxaConservacao + nTaxaLimpeza
            nValorIPTU98 = nValorIPTU98 * 1.6916
            nValorITU = 0
            nValorITU98 = 0
        Else
            nCodTributo = 2
            nValorVenalImovel = nValorVenalTerritorial
            nValorVenalImovel98 = nValorVenalTerritorial98
            nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)
'            nValorITU = nValorITU * 1.3916
            nValorITU98 = nValorVenalImovel98 * (nAliquotaTerritorial / 100)
            nValorITU98 = nValorITU98 + nTaxaConservacao + nTaxaLimpeza
            nValorITU98 = nValorITU98 * 1.6916
            nValorIPTU = nValorITU
            nValorIPTU98 = nValorITU98
        End If
        nValorIptuNovo = nValorIPTU
        'COMPARAÇÃO ENTRE OS CÁLCULOS
        If bTemPredial Then
            If nValorIPTU98 > nValorIPTU Then
               nValorFinal = nValorIPTU
            Else
               nValorFinal = nValorIPTU98
            End If
            nValorIPTU = nValorFinal
        Else
            If nValorITU98 > nValorITU Then
               nValorFinal = nValorITU
            Else
               nValorFinal = nValorITU98
            End If
            nValorITU = nValorFinal
        End If
        nValorUnica = Round(nValorFinal - (nValorFinal * (CDbl(lblPercUnica.Caption) / 100)), 2)
        nValorParcela = Round(nValorFinal / Val(txtNumParc.text), 2)
    'GoTo FIM
        'GRAVA TABELA LASERIPTU
'        Sql = "INSERT LASERIPTU (CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,"
'        Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,QTDEPARC,TXEXPPARC,TXEXPUNICA) VALUES("
'        Sql = Sql & nCodReduz & "," & Virg2Ponto(CStr(nValorVenalTerritorial)) & "," & Virg2Ponto(CStr(nValorVenalPredial)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorVenalImovel)) & "," & Virg2Ponto(CStr(nValorIPTU)) & "," & Virg2Ponto(CStr(nValorITU)) & ",'"
'        Sql = Sql & IIf(bTemPredial, "Predial", "Territorial") & "'," & Virg2Ponto(CStr(nAreaPrincipal)) & "," & Virg2Ponto(CStr(nTestada1)) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorParcela)) & "," & Virg2Ponto(CStr(nValorUnica)) & "," & Val(txtNumParc.text) & ","
'        Sql = Sql & Virg2Ponto(CStr(nValorExpDocParc) * Val(txtNumParc.text)) & "," & Virg2Ponto(CStr(nValorExpDocUnica)) & ")"
'        cn.Execute Sql, rdExecDirect
'#APAGAR
'        GoTo PROXIMO
        
'        For x = 0 To Val(txtNumParc.text)
'            If x = 0 And lblUnica.Caption = "Não" Then GoTo PROXIMO
            'GRAVA NA TABELA DEBITOPARCELA
            ax = nCodReduz & " " & nAnoCalculo & " " & FormatNumber(nValorIPTU98, 2) & " " & FormatNumber(nValorIptuNovo, 2) & " " & FormatNumber(nValorFinal, 2)
            Print #1, ax
            'GRAVA NA TABELA DEBITO TRIBUTO
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & nCodTributo & "," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
'            ax = nCodRedu 'z & "," & nAnoCalculo & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
'            ax = ax & 3 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ","
'            ax = ax & 0 & "," & 0 & "," & 0
'            Print #2, ax
'            'GRAVA NA TABELA NUMDOCUMENTO'
'            nLastDoc = nLastDoc + 1
'            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2)))
'            Print #4, ax
'            'GRAVA NA TABELA PARCELADOCUMENTO
'            ax = nCodReduz & "," & nAnoCalculo & "," & 1 & "," & 0 & ","
'            ax = ax & x & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
'            Print #3, ax
'        Next
PROXIMO:
        xId = xId + 1
       .MoveNext
    Loop
End With
'#APAGAR
'Exit Sub
Close #1

fim:
'lblWait.Caption = "CALCULO EFETUADO COM SUCESSO."
'lblWait.Refresh

End Sub

Private Sub cmdCancel_Click()

If MsgBox("Cancelar o Cálculo ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
     bExec = False
End If

End Sub

Private Sub cmdGravar_Click()
Dim nValorParcela As Double, nValorUnica As Double

If Not IsDate(mskDataBase.text) Then
    MsgBox "Data base inválida", vbCritical, "atenção"
    Exit Sub
End If


Sql = "SELECT CODLANCAMENTO,VALORPARCELA,VALORUNICA FROM EXPEDIENTE WHERE ANOEXPED=" & nAnoCalculo & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     nValorExpDocParc = FormatNumber(!VALORPARCELA, 2)
     nValorExpDocUnica = FormatNumber(!VALORUNICA, 2)
    .Close
End With

'ULTIMO Nº DE DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS ULTIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nLastDoc = !ULTIMO + 20
   .Close
End With

nCodReduz = Val(txtCod.text)
nAnoCalculo = Val(txtAnoCalculo.text)
sDataBase = mskDataBase.text
nValorParcela = CDbl(lblParcela.Caption)
nValorUnica = CDbl(lblUnica.Caption)

'APAGA

'TABELA PARCELA DOCUMENTO
Sql = "DELETE FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod) & " AND ANOEXERCICIO =" & nAnoCalculo & " AND (CODLANCAMENTO = 1 OR CODLANCAMENTO = 29)"
cn.Execute Sql, rdExecDirect

'TABELA NUMDOCUMENTO
Sql = "DELETE FROM NUMDOCUMENTO WHERE NUMDOCUMENTO in ("
Sql = Sql & "SELECT NumDocumento From PARCELADOCUMENTO WHERE CODREDUZIDO=" & Val(txtCod) & " AND  ANOEXERCICIO =" & nAnoCalculo & " AND (CODLANCAMENTO = 1 OR CODLANCAMENTO = 29))"
cn.Execute Sql, rdExecDirect


'TABELA DEBITOTRIBUTO
Sql = "DELETE FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & Val(txtCod) & " AND ANOEXERCICIO = " & nAnoCalculo & " AND (CODLANCAMENTO = 1 OR CODLANCAMENTO = 29)"
cn.Execute Sql, rdExecDirect

'TABELA DEBITOPAGO
Sql = "DELETE FROM DEBITOPAGO WHERE CODREDUZIDO=" & Val(txtCod) & " AND ANOEXERCICIO = " & nAnoCalculo & " AND (CODLANCAMENTO = 1 OR CODLANCAMENTO = 29)"
cn.Execute Sql, rdExecDirect
'TABELA DEBITOPARCELA
Sql = "DELETE FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod) & " AND ANOEXERCICIO = " & nAnoCalculo & " AND (CODLANCAMENTO = 1 OR CODLANCAMENTO = 29)"
cn.Execute Sql, rdExecDirect


For x = 0 To Val(txtNumParc.text)
    If lblTemUnica = "Não" And x = 0 Then x = 1
    'GRAVA NA TABELA DEBITOPARCELA
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA) VALUES(" & Val(txtCod.text) & "," & nAnoCalculo & ",1,0," & x & ",0,3,'"
    Sql = Sql & Format(aParc(x), "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "',1)"
    cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA DEBITO TRIBUTO
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
    Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.text) & "," & nAnoCalculo & ",1,0," & x & ",0,1," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ")"
    cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA DEBITO TRIBUTO
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
    Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(txtCod.text) & "," & nAnoCalculo & ",1,0," & x & ",0,3," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ")"
    cn.Execute Sql, rdExecDirect
    'GRAVA NA TABELA NUMDOCUMENTO
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
    Sql = Sql & nLastDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0," & Virg2Ponto(IIf(x = 0, Round(nValorExpDocUnica, 2), Round(nValorExpDocParc, 2))) & ")"
    cn.Execute Sql, rdExecDirect
    
    
    'GRAVA NA TABELA PARCELADOCUMENTO
    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
    Sql = Sql & Val(txtCod.text) & "," & nAnoCalculo & ",1,0," & x & ",0," & nLastDoc & ")"
    cn.Execute Sql, rdExecDirect
    nLastDoc = nLastDoc + 1
Next

MsgBox "Debito recalculado."

End Sub

Private Sub cmdPrint_Click()
Me.PrintForm
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Ocupado

If Right$(frmMdi.Sbar.Panels(2).text, Len(frmMdi.Sbar.Panels(2).text) - 9) = "FACTORE" Or Right$(frmMdi.Sbar.Panels(2).text, Len(frmMdi.Sbar.Panels(2).text) - 9) = "JULIANA" Then
   cmdGravar.Enabled = True
Else
   cmdGravar.Enabled = False
End If

Set xImovel = New clsImovel

Centraliza Me
Pb.Value = 0
lblPB.Caption = "0 %"
If Val(txtAnoCalculo.text) = 0 Then txtAnoCalculo.text = Year(Now)
nAnoCalculo = txtAnoCalculo.text
CarregaTela
Sql = "SELECT COUNT(CODREDUZIDO) AS TOTAL FROM CADIMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
       lblEstimado.Caption = !TOTAL
     End If
      .Close
End With
fim:
LoadMatrix

Liberado

sRet = RetEventUserForm(Me.Name)
frmMdi.AddWindow Me.Name, Me.Caption
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
      "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA= 1998 ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG= 1998 ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO= 1998 ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA= 1998 ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO= 1998 ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR= 1998  AND  CODMOEDA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorP98(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT98(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!CODSITUACAO) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS98(!CODSITUACAO) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG98(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD98(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!CODAGRUPAMENTO) = !VALORTERRENO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR98(!CODAGRUPAMENTO) = !VALORTERRENO
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
ReDim aFatorF98(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN; " & _
      "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN= 1998 ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        ReDim Preserve aFatorF98(UBound(aFatorF98) + 1)
        aFatorF98(UBound(aFatorF98)).Distrito = !CODDISTRITO
        aFatorF98(UBound(aFatorF98)).Codigo = !CODPROFUN
        aFatorF98(UBound(aFatorF98)).Fator = !FATORPROFUN
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
ReDim aFatorC98(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1; " & _
      "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=1998 AND CODMOEDA=1 "
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
    .MoreResults
     Do Until .EOF
        ReDim Preserve aFatorC98(UBound(aFatorC98) + 1)
        aFatorC98(UBound(aFatorC98)).Uso = !CODUSO
        aFatorC98(UBound(aFatorC98)).Tipo = !CodTipo
        aFatorC98(UBound(aFatorC98)).Categoria = !CODCATEG
        aFatorC98(UBound(aFatorC98)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
Set xImovel = Nothing
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.Value = (nPosF * 100) / nTotal
Else
   Pb.Value = 100
End If
lblPB.Caption = FormatNumber(Pb.Value, 2)

Me.Refresh
DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub



Private Sub mskDataBase_GotFocus()
mskDataBase.SetFocus
End Sub

Private Sub txtAnoCalculo_KeyPress(KeyAscii As Integer)
Tweak txtAnoCalculo, KeyAscii, IntegerPositive
End Sub

Private Sub txtAnoCalculo_LostFocus()
'If Val(txtAnoCalculo.text) = 2004 Or Val(txtAnoCalculo.text) = 2005 Then
   CarregaTela
'Else
'    MsgBox "Ano Inválido.", vbCritical, "atenção"
'End If
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
lblVVP98.Caption = "0,00"

End Sub

Private Sub CarregaImovel()

With xImovel
    .CarregaImovel Val(txtCod.text)
    lblProp.Caption = .NomePropPrincipal
    lblRua.Caption = .EnderecoCompleto
End With

End Sub

Private Sub CarregaTela()
nAnoCalculo = Val(txtAnoCalculo.text)
Sql = "SELECT ANO,QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,VENC01,VENC02,VENC03,VENC04,VENC05,"
Sql = Sql & "VENC06,VENC07,VENC08,VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA WHERE CODTIPO=1 AND ANO=" & nAnoCalculo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount = 0 Then GoTo fim
     txtNumParc.text = !QTDEPARCELA
     lblTemUnica.Caption = IIf(!PARCELAUNICA = "S", "Sim", "Não")
     lblPercUnica.Caption = FormatNumber(!DESCONTOUNICA, 2)
     ReDim aParc(!QTDEPARCELA)
     Do Until .EOF
        If lblTemUnica.Caption = "Sim" Then
            If Not IsNull(!VENCUNICA) Then aParc(0) = Format(!VENCUNICA, "dd/mm/yyyy")
        End If
        If Not IsNull(!VENC01) Then aParc(1) = Format(!VENC01, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC02) Then aParc(2) = Format(!VENC02, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC03) Then aParc(3) = Format(!VENC03, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC04) Then aParc(4) = Format(!VENC04, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC05) Then aParc(5) = Format(!VENC05, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC06) Then aParc(6) = Format(!VENC06, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC07) Then aParc(7) = Format(!VENC07, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC08) Then aParc(8) = Format(!VENC08, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC09) Then aParc(9) = Format(!VENC09, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC10) Then aParc(10) = Format(!VENC10, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC11) Then aParc(11) = Format(!VENC11, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!VENC12) Then aParc(12) = Format(!VENC12, "dd/mm/yyyy") Else Exit Do
        x = x + 1
       .MoveNext
     Loop
    .Close
End With
fim:
End Sub
