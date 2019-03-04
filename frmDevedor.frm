VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmDevedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Devedores"
   ClientHeight    =   6150
   ClientLeft      =   5250
   ClientTop       =   2685
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6150
   ScaleWidth      =   6570
   Begin VB.CheckBox chkCNPJ 
      Caption         =   "Só com CNPJ"
      Height          =   195
      Left            =   5130
      TabIndex        =   44
      Top             =   2220
      Width           =   1320
   End
   Begin VB.CheckBox chkSemSimples 
      Caption         =   "Sem Simples"
      Height          =   195
      Left            =   1320
      TabIndex        =   43
      Top             =   3630
      Width           =   1365
   End
   Begin VB.CheckBox chkSemMei 
      Caption         =   "Sem MEI"
      Height          =   195
      Left            =   135
      TabIndex        =   42
      Top             =   3630
      Width           =   1065
   End
   Begin VB.TextBox txtAnoAte 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2835
      MaxLength       =   4
      TabIndex        =   36
      Text            =   "10"
      Top             =   5760
      Width           =   780
   End
   Begin VB.TextBox txtAnoDe 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1665
      MaxLength       =   4
      TabIndex        =   35
      Text            =   "10"
      Top             =   5760
      Width           =   780
   End
   Begin VB.TextBox txtTop2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   34
      Text            =   "10"
      Top             =   5355
      Width           =   465
   End
   Begin VB.CheckBox chkMei 
      Caption         =   "Só MEI"
      Height          =   195
      Left            =   5130
      TabIndex        =   33
      Top             =   2745
      Width           =   1320
   End
   Begin VB.CheckBox chkSimples 
      Caption         =   "Só Simples"
      Height          =   195
      Left            =   5130
      TabIndex        =   32
      Top             =   2475
      Width           =   1320
   End
   Begin VB.CheckBox chkAtivo 
      Caption         =   "Só Ativos"
      Height          =   195
      Left            =   5130
      TabIndex        =   31
      Top             =   1965
      Width           =   1320
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Apenas totais"
      Height          =   195
      Index           =   2
      Left            =   5130
      TabIndex        =   30
      Top             =   3600
      Width           =   1320
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1470
      TabIndex        =   9
      Top             =   3060
      Width           =   885
   End
   Begin VB.TextBox txtTop 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1215
      MaxLength       =   3
      TabIndex        =   13
      Text            =   "10"
      Top             =   4680
      Width           =   465
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Resumido"
      Height          =   195
      Index           =   1
      Left            =   5130
      TabIndex        =   8
      Top             =   3330
      Width           =   1320
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Normal"
      Height          =   195
      Index           =   0
      Left            =   5130
      TabIndex        =   7
      Top             =   3060
      Value           =   -1  'True
      Width           =   1320
   End
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   3960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   344
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12632064
      Scrolling       =   1
   End
   Begin VB.ListBox lstLC1 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   420
      Width           =   4245
   End
   Begin VB.ListBox lstAno1 
      Appearance      =   0  'Flat
      Height          =   1380
      Left            =   4470
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   420
      Width           =   1965
   End
   Begin VB.CheckBox chkAno 
      Caption         =   "Verificar débitos nos demais anos além dos obrigatórios"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1890
      Width           =   4440
   End
   Begin VB.TextBox txtCod1 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1470
      MaxLength       =   6
      TabIndex        =   3
      Text            =   "000001"
      Top             =   2250
      Width           =   885
   End
   Begin VB.TextBox txtCod2 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3900
      MaxLength       =   6
      TabIndex        =   4
      Text            =   "000001"
      Top             =   2250
      Width           =   885
   End
   Begin VB.ComboBox cmbDA 
      Height          =   315
      ItemData        =   "frmDevedor.frx":0000
      Left            =   1440
      List            =   "frmDevedor.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2610
      Width           =   975
   End
   Begin VB.ComboBox cmbAj 
      Height          =   315
      ItemData        =   "frmDevedor.frx":0024
      Left            =   3900
      List            =   "frmDevedor.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2610
      Width           =   975
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   5190
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   3870
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDevedor.frx":0048
      PICN            =   "frmDevedor.frx":0064
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Cancel          =   -1  'True
      Height          =   360
      Left            =   3870
      TabIndex        =   11
      ToolTipText     =   "Emitir Relatório"
      Top             =   3870
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Executar"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDevedor.frx":00D2
      PICN            =   "frmDevedor.frx":00EE
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
      Height          =   360
      Left            =   3825
      TabIndex        =   23
      ToolTipText     =   "Imprimir maiores devedores"
      Top             =   4590
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   635
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
      MICON           =   "frmDevedor.frx":018D
      PICN            =   "frmDevedor.frx":01A9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Left            =   3900
      TabIndex        =   10
      Top             =   3060
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   503
      MouseIcon       =   "frmDevedor.frx":0303
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
   Begin prjChameleon.chameleonButton cmdCheckAll1 
      Height          =   225
      Left            =   2610
      TabIndex        =   26
      ToolTipText     =   "Marcar todos os lançamentos"
      Top             =   135
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BTYPE           =   3
      TX              =   "+"
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
      MICON           =   "frmDevedor.frx":031F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDelAll1 
      Height          =   225
      Left            =   2880
      TabIndex        =   27
      ToolTipText     =   "Desmarcar todos os lançamentos"
      Top             =   135
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BTYPE           =   3
      TX              =   "-"
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
      MICON           =   "frmDevedor.frx":033B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCheckAll2 
      Height          =   225
      Left            =   5850
      TabIndex        =   28
      ToolTipText     =   "Marcar todos os anos"
      Top             =   135
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BTYPE           =   3
      TX              =   "+"
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
      MICON           =   "frmDevedor.frx":0357
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDelAll2 
      Height          =   225
      Left            =   6120
      TabIndex        =   29
      ToolTipText     =   "Desmarcar todos os anos"
      Top             =   135
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   397
      BTYPE           =   3
      TX              =   "-"
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
      MICON           =   "frmDevedor.frx":0373
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btPrintPago 
      Height          =   360
      Left            =   3825
      TabIndex        =   37
      ToolTipText     =   "Imprimir maiores devedores"
      Top             =   5580
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   635
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
      MICON           =   "frmDevedor.frx":038F
      PICN            =   "frmDevedor.frx":03AB
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
      BackStyle       =   0  'Transparent
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   2565
      TabIndex        =   41
      Top             =   5760
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Entre os anos de:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   90
      TabIndex        =   40
      Top             =   5760
      Width           =   1620
   End
   Begin VB.Line Line 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   45
      X2              =   6405
      Y1              =   5175
      Y2              =   5190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Imprimir os "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   90
      TabIndex        =   39
      Top             =   5355
      Width           =   1350
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "maiores pagadores."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   1845
      TabIndex        =   38
      Top             =   5355
      Width           =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento até...:"
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   25
      Top             =   3105
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor até.............:"
      Height          =   255
      Index           =   7
      Left            =   135
      TabIndex        =   24
      Top             =   3105
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "maiores devedores."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   6
      Left            =   1890
      TabIndex        =   22
      Top             =   4680
      Width           =   1980
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Imprimir os "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   135
      TabIndex        =   21
      Top             =   4680
      Width           =   1350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   90
      X2              =   6450
      Y1              =   4410
      Y2              =   4425
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Selecione os lançamentos"
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
      Left            =   120
      TabIndex        =   19
      Top             =   150
      Width           =   4245
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   "Exercícios"
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
      Left            =   4470
      TabIndex        =   18
      Top             =   150
      Width           =   1965
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Inicial......:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Final........:"
      Height          =   255
      Left            =   2550
      TabIndex        =   16
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Divida Ativa........:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2670
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ajuizado..............:"
      Height          =   255
      Index           =   4
      Left            =   2550
      TabIndex        =   14
      Top             =   2670
      Width           =   1395
   End
End
Attribute VB_Name = "frmDevedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MDEV
    nCodReduz As Long
    sNome As String
    nTipo As Integer
    nValorPrincipal As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorAtual As Double
End Type

Private Type mDevPrint
    nTipo As Integer
    nCodReduz As Long
    sNome As String
    nValor As Double
End Type

Private Type Debito
    nCodReduz As Long
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
    nValorGeral As Double
    nValorHon As Double
    nValorJurApl As Double
    nSaldo As Double
    nCodBanco As Integer
    dDataPag As Date
    sNome As String
    sFullLanc As String
End Type

Private Type DebitoTotal
    nCodReduz As Long
    nValor As Double
End Type
Dim aAno() As Long, aLanc() As Long, aCodigoDebito() As Long

Private Sub btPrintPago_Click()
Dim Sql As String, RdoAux As rdoResultset, nPos As Integer, RdoAux2 As rdoResultset
Dim nCodReduz As Long, nValor As Double, sNome As String

Ocupado
Sql = "delete from maiorpagador where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

nPos = 1
Sql = "SELECT debitoparcela.codreduzido, ROUND(SUM(debitotributo.valortributo), 2) AS Soma "
Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
Sql = Sql & "WHERE DEBITOPARCELA.CODLANCAMENTO=5 AND  (debitoparcela.statuslanc = 2) AND DEBITOPARCELA.ANOEXERCICIO BETWEEN " & Val(txtAnoDe.Text) & " AND " & Val(txtAnoAte.Text)
Sql = Sql & " GROUP BY debitoparcela.codreduzido Having (debitoparcela.CODREDUZIDO > 0) ORDER BY Soma DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        nValor = !soma
        If nCodReduz < 500000 Then
            Sql = "select razaosocial as nome from mobiliario where codigomob=" & nCodReduz
        Else
            Sql = "select nomecidadao as nome from cidadao where codcidadao=" & nCodReduz
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = RdoAux2!Nome
        RdoAux2.Close
        
        Sql = "insert maiorpagador(usuario,codigo,nome,valor) values('" & NomeDeLogin & "'," & nCodReduz & ",'" & Mask(sNome) & "'," & Virg2Ponto(CStr(nValor)) & ")"
        cn.Execute Sql, rdExecDirect
        
        
        If nPos >= Val(txtTop2.Text) Then Exit Do
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
Liberado

frmReport.ShowReport2 "MAIORPAGADOR", frmMdi.HWND, Me.HWND
Sql = "delete from maiorpagador where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdCheckAll1_Click()
Dim x As Integer
For x = 0 To lstLC1.ListCount - 1
    lstLC1.Selected(x) = True
Next
lstLC1.ListIndex = 0
End Sub

Private Sub cmdCheckAll2_Click()
Dim x As Integer
For x = 0 To lstAno1.ListCount - 1
    lstAno1.Selected(x) = True
Next
lstAno1.ListIndex = 0

End Sub

Private Sub cmdDelAll1_Click()
Dim x As Integer
For x = 0 To lstLC1.ListCount - 1
    lstLC1.Selected(x) = False
Next

End Sub

Private Sub cmdDelAll2_Click()
Dim x As Integer
For x = 0 To lstAno1.ListCount - 1
    lstAno1.Selected(x) = False
Next

End Sub

Private Sub cmdExec_Click()
Dim x As Integer, bAchou As Boolean
Dim RdoAux As rdoResultset, Sql As String
ReDim aCodigoDebito(0)
bAchou = False
For x = 0 To lstLC1.ListCount - 1
    If lstLC1.Selected(x) = True Then
        bAchou = True
        Exit For
    End If
Next
If Not bAchou Then
    MsgBox "Selecione ao menos um lançamento.", vbExclamation, "Atenção"
    Exit Sub
End If
For x = 0 To lstAno1.ListCount - 1
    If lstAno1.Selected(x) = True Then
        bAchou = True
        Exit For
    End If
Next
If Not bAchou Then
    MsgBox "Selecione ao menos um Ano.", vbExclamation, "Atenção"
    Exit Sub
End If
If Val(txtCod1.Text) = 0 Then
    MsgBox "Código inicial inválido.", vbExclamation, "Atenção"
    Exit Sub
End If
If Val(txtCod2.Text) = 0 Then
    MsgBox "Código final inválido.", vbExclamation, "Atenção"
    Exit Sub
End If
If Val(txtCod1.Text) > Val(txtCod2.Text) Then
    MsgBox "Código inicial maior que o inicial.", vbExclamation, "Atenção"
    Exit Sub
End If

Main2

Open sPathBin & "\ListaDevedor.txt" For Output As #1
For x = 1 To UBound(aCodigoDebito)
    If UBound(aCodigoDebito) = x Then
        Print #1, CStr(aCodigoDebito(x))
    Else
        Print #1, CStr(aCodigoDebito(x)) & ","
    End If
Next

Close #1

End Sub

Private Sub cmdPrint_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, aI() As MDEV, aM() As MDEV, aC() As MDEV, nMax As Integer
Dim p As Integer, aPrint() As mDevPrint, qd As New rdoQuery, aDebito() As Debito, nEval As Integer, sLanc As String, sAno As String
Dim Achou As Boolean, x As Integer, aMat() As MDEV, Y As Integer, aLanc() As Long, aAno() As Long

ReDim aAno(0): ReDim aLanc(0)
sLanc = "": sAno = ""
For x = 0 To lstLC1.ListCount - 1
    If lstLC1.Selected(x) = True Then
        ReDim Preserve aLanc(UBound(aLanc) + 1)
        aLanc(UBound(aLanc)) = lstLC1.ItemData(x)
        sLanc = sLanc & lstLC1.ItemData(x) & ","
    End If
Next
sLanc = Left(sLanc, Len(sLanc) - 1)
For x = 0 To lstAno1.ListCount - 1
    If lstAno1.Selected(x) = True Then
        ReDim Preserve aAno(UBound(aAno) + 1)
        aAno(UBound(aAno)) = lstAno1.List(x)
        sAno = sAno & lstAno1.List(x) & ","
    End If
Next
sAno = Left(sAno, Len(sAno) - 1)

TriQuickSortLong aAno
TriQuickSortLong aLanc

If MsgBox("Deseja imprimir o relatório?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
    Exit Sub
End If

If Option1(2).value = True Then
    GeraApenasTotais
    Exit Sub
End If

Ocupado
DoEvents
nMax = Val(txtTop.Text)
If nMax = 0 Then nMax = 10
ReDim aI(0): ReDim aM(0): ReDim aC(0): ReDim aPrint(0)

cn.QueryTimeout = 0
Sql = "SELECT debitoparcela.codreduzido, ROUND(SUM(debitotributo.valortributo), 2) AS Soma FROM debitoparcela INNER JOIN "
Sql = Sql & "debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
Sql = Sql & "WHERE (debitoparcela.statuslanc = 3 or debitoparcela.statuslanc = 42 or debitoparcela.statuslanc = 43) AND (debitotributo.codtributo <> 3) AND (debitoparcela.datavencimento < GETDATE()) AND "
Sql = Sql & "(debitoparcela.numparcela > 0) AND debitoparcela.CODLANCAMENTO in (" & sLanc & ") AND debitoparcela.ANOEXERCICIO in (" & sAno & ") GROUP BY debitoparcela.codreduzido Having (debitoparcela.CODREDUZIDO between " & Val(txtCod1.Text) & " and " & Val(txtCod2.Text) & ") ORDER BY Soma DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
       ' If UBound(aI) = nMax And UBound(aM) = nMax And UBound(aC) = nMax Then Exit Do
        If nCodReduz < 100000 Then
            If UBound(aI) < nMax Then
                ReDim Preserve aI(UBound(aI) + 1)
                aI(UBound(aI)).nCodReduz = nCodReduz
            End If
        ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
            If UBound(aM) < nMax Then
                ReDim Preserve aM(UBound(aM) + 1)
                aM(UBound(aM)).nCodReduz = nCodReduz
            End If
        ElseIf nCodReduz >= 500000 And nCodReduz < 700000 Then
            If UBound(aC) < nMax Then
                ReDim Preserve aC(UBound(aC) + 1)
                aC(UBound(aC)).nCodReduz = nCodReduz
            End If
        End If
       .MoveNext
    Loop
   .Close
End With

For p = 1 To UBound(aI)
    Sql = "SELECT NOMECIDADAO FROM vwFULLIMOVEL WHERE CODREDUZIDO=" & aI(p).nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    ReDim Preserve aPrint(UBound(aPrint) + 1)
    aPrint(UBound(aPrint)).nTipo = 1
    aPrint(UBound(aPrint)).nCodReduz = aI(p).nCodReduz
    aPrint(UBound(aPrint)).sNome = RdoAux!nomecidadao
    RdoAux.Close
Next

For p = 1 To UBound(aM)
    Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & aM(p).nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    ReDim Preserve aPrint(UBound(aPrint) + 1)
    aPrint(UBound(aPrint)).nTipo = 2
    aPrint(UBound(aPrint)).nCodReduz = aM(p).nCodReduz
    aPrint(UBound(aPrint)).sNome = SubNull(RdoAux!razaosocial)
    RdoAux.Close
Next

For p = 1 To UBound(aC)
    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & aC(p).nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    ReDim Preserve aPrint(UBound(aPrint) + 1)
    aPrint(UBound(aPrint)).nTipo = 3
    aPrint(UBound(aPrint)).nCodReduz = aC(p).nCodReduz
    aPrint(UBound(aPrint)).sNome = SubNull(RdoAux!nomecidadao)
    RdoAux.Close
Next

Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Pb.value = 0
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
ReDim aDebito(0)
For p = 1 To UBound(aPrint)
    CallPb CLng(p), CLng(UBound(aPrint))
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
    qd(0) = aPrint(p).nCodReduz:  qd(1) = aPrint(p).nCodReduz
    qd(2) = 1950: qd(3) = 2050
    qd(4) = 0: qd(5) = 99 'lanc
    qd(6) = 0: qd(7) = 9999 'seq
    qd(8) = 1: qd(9) = 99 'parc
    qd(10) = 0: qd(11) = 99 'compl
    qd(12) = 0: qd(13) = 99 'stat
    qd(14) = Format(Now, "mm/dd/yyyy")
    qd(15) = NomeDoUsuario
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        
        If RdoAux.RowCount > 0 Then
            Do Until .EOF
                If !statuslanc <> 3 And !statuslanc <> 42 And !statuslanc <> 43 Then GoTo proximo
            
                z = BinarySearchLong(aAno(), CLng(!AnoExercicio))
                If z = -1 Then GoTo proximo
                
                z = BinarySearchLong(aLanc(), CLng(!CodLancamento))
                If z = -1 Then GoTo proximo

            
                DoEvents
                If CDate(Format(!DataVencimento, "dd/mm/yyyy")) > CDate(Format(Now, "dd/mm/yyyy")) Then
                    GoTo proximo:
                End If
                nEval = UBound(aDebito)
                Achou = False
                For x = 1 To nEval
                    If aDebito(x).nCodReduz = aPrint(p).nCodReduz And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                       aDebito(x).nSeq = !SeqLancamento And _
                       aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                       Achou = True
                       Exit For
                    End If
                Next
                
                If Not Achou Then
                    ReDim Preserve aDebito(UBound(aDebito) + 1)
                    nEval = UBound(aDebito)
                    aDebito(nEval).nCodReduz = aPrint(p).nCodReduz
                    aDebito(nEval).sNome = aPrint(p).sNome
                    aDebito(nEval).nCodBanco = aPrint(p).nTipo
                    aDebito(nEval).nAno = !AnoExercicio
                    aDebito(nEval).nLanc = !CodLancamento
                    aDebito(nEval).sLanc = !DESCLANCAMENTO
                    aDebito(nEval).nSeq = !SeqLancamento
                    aDebito(nEval).nParc = !NumParcela
                    aDebito(nEval).nCompl = !CODCOMPLEMENTO
                    aDebito(nEval).nSituacao = !statuslanc
                    aDebito(nEval).sSituacao = !Situacao
                    aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                    aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                    aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                    aDebito(nEval).nCodTributo = !CodTributo
                    aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                    aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                    aDebito(nEval).nValorJuros = FormatNumber(!ValorJuros, 2)
                    aDebito(nEval).nValorMulta = FormatNumber(!ValorMulta, 2)
                    aDebito(nEval).nValorCorrecao = FormatNumber(!ValorCorrecao, 2)
                Else
                    If aDebito(x).nCodTributo = !CodTributo Then GoTo proximo
                
                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                    aDebito(x).nValorJuros = FormatNumber(aDebito(x).nValorJuros + !ValorJuros, 2)
                    aDebito(x).nValorMulta = FormatNumber(aDebito(x).nValorMulta + !ValorMulta, 2)
                    aDebito(x).nValorCorrecao = FormatNumber(aDebito(x).nValorCorrecao + !ValorCorrecao, 2)
                    aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTotal, 2)
                End If
proximo:
                .MoveNext
            Loop
          End If
       .Close
    End With
Next

'ordena a matriz
ReDim aMat(0)
For x = 1 To UBound(aDebito)
    Achou = False
    For Y = 1 To UBound(aMat)
        If aDebito(x).nCodReduz = aMat(Y).nCodReduz Then
            Achou = True
            Exit For
        End If
    Next
    If Not Achou Then
        ReDim Preserve aMat(UBound(aMat) + 1)
        aMat(UBound(aMat)).nCodReduz = aDebito(x).nCodReduz
        aMat(UBound(aMat)).nTipo = aDebito(x).nCodBanco
        aMat(UBound(aMat)).sNome = aDebito(x).sNome
        aMat(UBound(aMat)).nValorPrincipal = aDebito(x).nValorTributo
        aMat(UBound(aMat)).nValorJuros = aDebito(x).nValorJuros
        aMat(UBound(aMat)).nValorMulta = aDebito(x).nValorMulta
        aMat(UBound(aMat)).nValorCorrecao = aDebito(x).nValorCorrecao
        aMat(UBound(aMat)).nValorAtual = aDebito(x).nValorAtual
    Else
        aMat(Y).nValorPrincipal = aMat(Y).nValorPrincipal + aDebito(x).nValorTributo
        aMat(Y).nValorMulta = aMat(Y).nValorMulta + aDebito(x).nValorMulta
        aMat(Y).nValorJuros = aMat(Y).nValorJuros + aDebito(x).nValorJuros
        aMat(Y).nValorCorrecao = aMat(Y).nValorCorrecao + aDebito(x).nValorCorrecao
        aMat(Y).nValorAtual = aMat(Y).nValorAtual + aDebito(x).nValorAtual
    End If
Next

For x = 1 To UBound(aMat)
    With aMat(x)
        Sql = "INSERT EXTRATOTMP (COMPUTER,SEQ,CODREDUZIDO,NOMEPROP,CODBANCO,VALORLANCADO,VALORJUROS,VALORMULTA,VALORCORRECAO,VALORTOTAL) VALUES('" & NomeDeLogin & "',"
        Sql = Sql & x & "," & .nCodReduz & ",'" & Mask(Left$(.sNome, 30)) & "'," & .nTipo & "," & Virg2Ponto(CStr(.nValorPrincipal)) & "," & Virg2Ponto(CStr(.nValorJuros)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorCorrecao)) & "," & Virg2Ponto(CStr(.nValorAtual)) & ")"
        cn.Execute Sql, rdExecDirect
    End With
PROXIMO2:
Next

frmReport.ShowReport2 "MAIORDEVEDOR", frmMdi.HWND, Me.HWND

Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Liberado
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
txtAnoDe.Text = Year(Now)
txtAnoAte.Text = Year(Now)
Init
End Sub

Private Sub Init()
Dim x As Integer
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT CODLANCAMENTO,DESCFULL FROM LANCAMENTO WHERE CODLANCAMENTO<>25"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstLC1.AddItem !DESCFULL
        lstLC1.ItemData(lstLC1.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With

For x = 1990 To Year(Now)
    lstAno1.AddItem x
Next

cmbDA.ListIndex = 0: cmbAj.ListIndex = 0

End Sub

Private Sub mskVenc_GotFocus()
mskVenc.SelStart = 0
mskVenc.SelLength = Len(mskVenc.Text)
End Sub

Private Sub txtCod1_GotFocus()
txtCod1.SelStart = 0
txtCod1.SelLength = 6
End Sub

Private Sub txtCod1_KeyPress(KeyAscii As Integer)
Tweak txtCod1, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod1_LostFocus()
txtCod1.Text = Format(Val(txtCod1.Text), "000000")
End Sub

Private Sub txtCod2_GotFocus()
txtCod2.SelStart = 0
txtCod2.SelLength = 6
End Sub

Private Sub txtCod2_KeyPress(KeyAscii As Integer)
Tweak txtCod2, KeyAscii, IntegerPositive
End Sub

Private Sub Main()
Dim RdoAux As rdoResultset, Sql As String
Dim nCod As Long, nLanc() As Integer, sLanc As String, nAno As Integer, sAno As String
Dim aAno() As Integer, aLanc() As Integer, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, sNome As String
ReDim aAno(0): ReDim aLanc(0)
sLanc = "": sAno = ""
Pb.value = 0
For x = 0 To lstLC1.ListCount - 1
    If lstLC1.Selected(x) = True Then
        ReDim Preserve aLanc(UBound(aLanc) + 1)
        aLanc(UBound(aLanc)) = lstLC1.ItemData(x)
        sLanc = sLanc & lstLC1.ItemData(x) & ","
    End If
Next
sLanc = Left(sLanc, Len(sLanc) - 1)
For x = 0 To lstAno1.ListCount - 1
    If lstAno1.Selected(x) = True Then
        ReDim Preserve aAno(UBound(aAno) + 1)
        aAno(UBound(aAno)) = lstAno1.List(x)
        sAno = sAno & lstAno1.List(x) & ","
    End If
Next
sAno = Left(sAno, Len(sAno) - 1)

Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT distinct CODREDUZIDO FROM DEBITOPARCELA WHERE CODREDUZIDO BETWEEN "
Sql = Sql & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & " AND CODLANCAMENTO in ("
Sql = Sql & sLanc & ")"
If cmbDA.ListIndex = 1 Then
    Sql = Sql & " AND DATAINSCRICAO IS NOT NULL"
ElseIf cmbDA.ListIndex = 2 Then
    Sql = Sql & " AND DATAINSCRICAO IS NULL"
End If
If cmbAj.ListIndex = 1 Then
    Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
ElseIf cmbAj.ListIndex = 2 Then
    Sql = Sql & " AND DATAAJUIZA IS NULL"
End If
If IsDate(mskVenc.Text) Then
    Sql = Sql & " AND DATAVENCIMENTO <='" & Format(mskVenc.Text, "mm/dd/yyyy") & "'"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
      nPos = .AbsolutePosition
      If nPos Mod 10 = 0 Then
         CallPb nPos, CLng(nTot)
      End If
        For x = 1 To UBound(aAno)
            bAchou = True
            Sql = "SELECT CODREDUZIDO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & aAno(x)
            Sql = Sql & " AND (STATUSLANC=3 or STATUSLANC=19 or STATUSLANC=42 or STATUSLANC=43) AND CODLANCAMENTO in (" & sLanc & ")"
            If cmbDA.ListIndex = 1 Then
                Sql = Sql & " AND DATAINSCRICAO IS NOT NULL"
            ElseIf cmbDA.ListIndex = 2 Then
                Sql = Sql & " AND DATAINSCRICAO IS NULL"
            End If
            If cmbAj.ListIndex = 1 Then
                Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
            ElseIf cmbAj.ListIndex = 2 Then
                Sql = Sql & " AND DATAAJUIZA IS NULL"
            End If
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If RdoAux2.RowCount = 0 And chkAno.value = 0 Then
                    GoTo proximo
                End If
               .Close
            End With
        Next

        If !CODREDUZIDO < 100000 Then
            Sql = "SELECT NOMECIDADAO,INATIVO FROM VWCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & !CODREDUZIDO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    If !Inativo = 1 Then: .Close: GoTo proximo
                    sNome = !nomecidadao
                Else
                    sNome = ""
                End If
               .Close
            End With
        ElseIf !CODREDUZIDO > 100000 And !CODREDUZIDO < 500000 Then
            Sql = "SELECT RAZAOSOCIAL,DATAENCERRAMENTO FROM MOBILIARIO WHERE CODIGOMOB=" & !CODREDUZIDO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If chkAtivo.value = vbChecked Then
                    If Not IsNull(!dataencerramento) Then: .Close: GoTo proximo
                End If
                sNome = !razaosocial
               .Close
            End With
        Else
            Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !CODREDUZIDO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    sNome = !nomecidadao
                Else
                    sNome = ""
                End If
               .Close
            End With
        End If
    
        Sql = "SELECT * FROM vwCNSLANCAMENTO WHERE CODREDUZIDO=" & !CODREDUZIDO
        Sql = Sql & " AND (STATUSLANC=3 OR STATUSLANC=19 or STATUSLANC=42 or STATUSLANC=43) and NUMPARCELA>0 AND CODTRIBUTO<>3 AND CODLANCAMENTO in (" & sLanc & ")"
        If chkAno.value = 0 Then
            Sql = Sql & " AND ANOEXERCICIO in (" & sAno & ")"
        End If
        If cmbDA.ListIndex = 1 Then
            Sql = Sql & " AND DATAINSCRICAO IS NOT NULL"
        ElseIf cmbDA.ListIndex = 2 Then
            Sql = Sql & " AND DATAINSCRICAO IS NULL"
        End If
        If cmbAj.ListIndex = 1 Then
            Sql = Sql & " AND DATAAJUIZA IS NOT NULL"
        ElseIf cmbAj.ListIndex = 2 Then
            Sql = Sql & " AND DATAAJUIZA IS NULL"
        End If
        If IsDate(mskVenc.Text) Then
            Sql = Sql & " AND DATAVENCIMENTO <='" & Format(mskVenc.Text, "mm/dd/yyyy") & "'"
        End If
        If Val(txtValor.Text) > 0 Then
            Sql = Sql & " AND TOTALLANCADO <=" & Virg2Ponto(txtValor.Text)
        End If
        
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                Sql = "INSERT DAM(COMPUTER,SEQ,CODREDUZIDO,ANOEXERC,LANC,NUMSEQ,NUMPARCELA,COMP,DATAVENCTO,FULLLANC,NOMECONTRIBUINTE,"
                Sql = Sql & "PRINCIPAL) VALUES('" & NomeDoUsuario & "'," & 1 & "," & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & ","
                Sql = Sql & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','"
                Sql = Sql & !descreduz & "','" & Right$(Mask(sNome), 40) & "'," & Virg2Ponto(!TOTALLANCADO) & ")"
                cn.Execute Sql, rdExecDirect
               .MoveNext
            Loop
           .Close
        End With

proximo:
       .MoveNext
    Loop
End With
Pb.value = 100

'EXIBE RELATORIO
If Option1(0).value = True Then
    frmReport.ShowReport "DEVEDORES", frmMdi.HWND, Me.HWND
Else
    frmReport.ShowReport "DEVEDORES2", frmMdi.HWND, Me.HWND
End If

Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Main2()
Dim x As Integer, nCodReduz1 As Long, nCodReduz2 As Long, aDebito() As Debito, nCodImovel As Long, z As Long
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, aDebitoTotal() As DebitoTotal, Y As Integer
Dim nValorDebito As Double, Achou As Boolean, sExecFiscal As String, sNome As String
Dim nSomaDebito As Double, nEval As Integer, nValorCorrecao As Double, sFullLanc As String
Dim nSomaVencer As Double, nSomaDebitoUnica As Double, nSomaVencerUnica As Double
Dim sDescReduz As String, nValorAtualizado As Double, nSomaValorTributo As Double
Dim bAjuiza As Boolean, bDA As Boolean, qd As New rdoQuery, bIsentoMJ As Boolean, sSimples As String

ReDim aAno(0): ReDim aLanc(0)
Ocupado
Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

For x = 0 To lstLC1.ListCount - 1
    If lstLC1.Selected(x) = True Then
        ReDim Preserve aLanc(UBound(aLanc) + 1)
        aLanc(UBound(aLanc)) = lstLC1.ItemData(x)
    End If
Next
For x = 0 To lstAno1.ListCount - 1
    If lstAno1.Selected(x) = True Then
        ReDim Preserve aAno(UBound(aAno) + 1)
        aAno(UBound(aAno)) = lstAno1.List(x)
    End If
Next

TriQuickSortLong aAno
TriQuickSortLong aLanc
DoEvents
nCodReduz1 = Val(txtCod1.Text)
nCodReduz2 = Val(txtCod2.Text)
nPos = 1
nTot = nCodReduz2 - nCodReduz1
For nCodImovel = nCodReduz1 To nCodReduz2
    If nPos Mod 10 = 0 Then
       CallPb nPos, CLng(nTot)
    End If
    
    If nCodImovel > 100000 And nCodImovel < 300000 Then
        If chkAtivo.value = vbChecked Then
            Sql = "SELECT CODTIPOEVENTO,DATAEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodImovel
            Sql = Sql & " ORDER BY DATAEVENTO DESC"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    If !CODTIPOEVENTO = 2 Then
                        GoTo proximo
                    End If
                End If
               .Close
            End With
        End If
        If chkMei.value = vbChecked Then
            Sql = "SELECT CODIGOMOB,MEI FROM MOBILIARIO WHERE CODIGOMOB=" & nCodImovel
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    If Val(SubNull(!Mei)) = 0 Then
                        GoTo proximo
                    End If
                End If
               .Close
            End With
        End If
        If chkSimples.value = vbChecked Then
            Sql = "SELECT Tributacao.dbo.RETORNASN(" & Format(Val(nCodImovel), "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
            If RdoAux2!RETORNO = 0 Then
                GoTo proximo
            End If
        End If
    
        If chkSemMei.value = vbChecked Then
            Sql = "SELECT CODIGOMOB,MEI FROM MOBILIARIO WHERE CODIGOMOB=" & nCodImovel
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    If Val(SubNull(!Mei)) = 1 Then
                        GoTo proximo
                    End If
                End If
               .Close
            End With
           
        End If
    
        If chkSemSimples.value = vbChecked Then
            
            Sql = "SELECT Tributacao.dbo.RETORNASN(" & Format(Val(nCodImovel), "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
            If RdoAux2!RETORNO = 1 Then
                GoTo proximo
            End If
        End If
        
    
    End If
    
    Calcula (nCodImovel)

proximo:
    nPos = nPos + 1
Next


Pb.value = 0

Liberado
Sql = "SELECT * FROM DAM WHERE COMPUTER='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "A consulta não gerou nenhum resultado.", vbExclamation, "Atenção"
    Exit Sub
End If
RdoAux.Close

'EXIBE RELATORIO
If Option1(0).value = True Then
    frmReport.ShowReport "DEVEDORES", frmMdi.HWND, Me.HWND
ElseIf Option1(1).value = True Then
    frmReport.ShowReport "DEVEDORES3", frmMdi.HWND, Me.HWND
Else
    frmReport.ShowReport "DEVEDORES2", frmMdi.HWND, Me.HWND
End If

Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Calcula(nCodImovel As Long)
Dim x As Integer, nCodReduz1 As Long, nCodReduz2 As Long, aDebito() As Debito, z As Long
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPos As Long, nTot As Long, aDebitoTotal() As DebitoTotal, Y As Integer
Dim nValorDebito As Double, Achou As Boolean, sExecFiscal As String, sNome As String, RdoAux3 As rdoResultset
Dim nSomaDebito As Double, nEval As Integer, nValorCorrecao As Double, sFullLanc As String
Dim nSomaVencer As Double, nSomaDebitoUnica As Double, nSomaVencerUnica As Double, sCNPJ As String
Dim sDescReduz As String, nValorAtualizado As Double, nSomaValorTributo As Double
Dim bAjuiza As Boolean, bDA As Boolean, qd As New rdoQuery, bIsentoMJ As Boolean, aAnoFull() As Integer, sAno As String


ReDim aDebito(0): ReDim aDebitoTotal(0)
sCNPJ = ""
If nCodImovel < 100000 Then
    Sql = "SELECT NOMECIDADAO,INATIVO FROM VWCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodImovel
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            'If !Inativo = 1 Then: .Close: GoTo PROXIMO
            If !Inativo = 1 Then: .Close: Exit Sub
            sNome = !nomecidadao
        Else
            sNome = ""
        End If
       .Close
    End With
ElseIf nCodImovel > 100000 And nCodImovel < 500000 Then
    Sql = "SELECT RAZAOSOCIAL,DATAENCERRAMENTO,CNPJ FROM MOBILIARIO WHERE CODIGOMOB=" & nCodImovel
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            'If Not IsNull(!DATAENCERRAMENTO) Then: .Close: GoTo PROXIMO
            If chkAtivo.value = vbChecked Then
                If Not IsNull(!dataencerramento) Then: .Close: Exit Sub
            End If
            sCNPJ = SubNull(!Cnpj)
            If sCNPJ <> "" Then
                sCNPJ = Format(!Cnpj, "0#\.###\.###/####-##")
            Else
                If chkCNPJ.value = vbChecked Then: .Close: Exit Sub
            End If
            sNome = !razaosocial
        End If
       .Close
    End With
Else
    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodImovel
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            sNome = !nomecidadao
        Else
            sNome = ""
        End If
       .Close
    End With
End If
DoEvents
'CARREGA O EXTRATO
Set qd.ActiveConnection = cn
qd.QueryTimeout = 0
On Error Resume Next
RdoAux.Close
On Error GoTo 0
qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = nCodImovel
qd(1) = nCodImovel
qd(2) = 1950 'ano
qd(3) = 2050
qd(4) = 0 'lanc
qd(5) = 99
qd(6) = 0 'seq
qd(7) = 9999
qd(8) = 1 'parc
qd(9) = 999
qd(10) = 0 'compl
qd(11) = 99
qd(12) = 0 'status
qd(13) = 99
qd(14) = Format(Now, "mm/dd/yyyy")
qd(15) = NomeDoUsuario
Set RdoAux = qd.OpenResultset(rdOpenKeyset)

With RdoAux
    If RdoAux.RowCount > 0 Then

        nEval = UBound(aDebito)
        Do Until .EOF
            bJuros = False: bMulta = False
            If cmbAj.ListIndex = 1 Then
                If IsNull(!dataajuiza) Then GoTo proximo
            End If
            If cmbAj.ListIndex = 2 Then
                If Not IsNull(!dataajuiza) Then GoTo proximo
            End If
            If cmbDA.ListIndex = 1 Then
                If IsNull(!datainscricao) Then GoTo proximo
            End If
            If cmbDA.ListIndex = 2 Then
                If Not IsNull(!datainscricao) Then GoTo proximo
            End If
            If !statuslanc <> 3 And !statuslanc <> 18 And !statuslanc <> 19 And !statuslanc <> 38 And !statuslanc <> 39 And !statuslanc <> 42 And !statuslanc <> 43 And !statuslanc <> 40 And !statuslanc <> 31 Then GoTo proximo
            If !CodTributo = 3 Then GoTo proximo
            
            'If !AnoExercicio = 2007 Then
            '    MsgBox "teste"
            'End If
            z = BinarySearchLong(aAno(), CLng(!AnoExercicio))
            If z = -1 Then GoTo proximo
            
            z = BinarySearchLong(aLanc(), CLng(!CodLancamento))
            If z = -1 Then GoTo proximo
            
            If IsDate(mskVenc.Text) Then
                If !DataVencimento > CDate(mskVenc.Text) Then GoTo proximo
            End If

            z = BinarySearchLong(aCodigoDebito(), CLng(!CODREDUZIDO))
            If z = -1 Then
                ReDim Preserve aCodigoDebito(UBound(aCodigoDebito) + 1)
                aCodigoDebito(UBound(aCodigoDebito)) = !CODREDUZIDO
            End If
            
            Achou = False
            For x = 1 To nEval
                If aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                   aDebito(x).nSeq = !SeqLancamento And _
                   aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                   Achou = True
                   Exit For
                End If
            Next
            
            If Not Achou Then
                ReDim Preserve aDebito(UBound(aDebito) + 1)
                nEval = UBound(aDebito)
                aDebito(nEval).nCodReduz = !CODREDUZIDO
                aDebito(nEval).nAno = !AnoExercicio
                aDebito(nEval).nLanc = !CodLancamento
                aDebito(nEval).nSeq = !SeqLancamento
                aDebito(nEval).nParc = !NumParcela
                aDebito(nEval).nCompl = !CODCOMPLEMENTO
                aDebito(nEval).nSituacao = !statuslanc
                aDebito(nEval).sSituacao = !Situacao
                aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                aDebito(nEval).sDA = IIf(IsNull(!datainscricao), "N", "S")
                aDebito(nEval).sAj = IIf(IsNull(!dataajuiza), "N", "S")
                aDebito(nEval).nCodTributo = !CodTributo
                aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                aDebito(nEval).nValorMulta = FormatNumber(!ValorMulta, 2)
                aDebito(nEval).nValorJuros = FormatNumber(!ValorJuros, 2)
                aDebito(nEval).nValorCorrecao = FormatNumber(!ValorCorrecao, 2)
                aDebito(nEval).nValorAtual = FormatNumber(!ValorTotal, 2)
                aDebito(nEval).nValorGeral = FormatNumber(!ValorTotal, 2)
                aDebito(nEval).sNome = sNome
                aDebito(nEval).sFullLanc = !DESCLANCAMENTO
            Else
                If aDebito(x).nCodTributo <> !CodTributo Then
                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                    aDebito(x).nValorAtual = FormatNumber(aDebito(x).nValorAtual + !ValorTotal, 2)
                    aDebito(x).nValorGeral = FormatNumber(aDebito(x).nValorGeral + !ValorTotal, 2)
                End If
            End If
proximo:
                        
            .MoveNext
        Loop
      End If
   .Close
End With
nPos = nPos + 1

ReDim Preserve aDebitoTotal(UBound(aDebitoTotal) + 1)
aDebitoTotal(UBound(aDebitoTotal)).nCodReduz = nCodImovel
For Y = 1 To UBound(aDebito)
    
    If aDebito(Y).nCodReduz = nCodImovel Then
        aDebitoTotal(UBound(aDebitoTotal)).nValor = aDebitoTotal(UBound(aDebitoTotal)).nValor + aDebito(Y).nValorGeral
    End If
Next


sAno = "": ReDim aAnoFull(0)
For x = 0 To UBound(aDebito)
    Achou = False
    For Y = 0 To UBound(aAnoFull)
        If aAnoFull(Y) = aDebito(x).nAno Then
            Achou = True
            Exit For
        End If
    Next
    If Not Achou And aDebito(x).nAno > 0 Then
        ReDim Preserve aAnoFull(UBound(aAnoFull) + 1)
        aAnoFull(UBound(aAnoFull)) = aDebito(x).nAno
        sAno = sAno & aDebito(x).nAno & ","
    End If
Next
If Len(sAno) + 0 Then
    sAno = Left(sAno, Len(sAno) - 1)
End If

For x = 1 To UBound(aDebito)
    With aDebito(x)
'        If .nCodReduz = 21102 Then MsgBox "teste"
        If Val(txtValor.Text) > 0 Then
            nValorAtualizado = 0
            For Y = 0 To UBound(aDebitoTotal)
                If aDebitoTotal(Y).nCodReduz = aDebito(x).nCodReduz Then
                    nValorAtualizado = aDebitoTotal(Y).nValor
                    Exit For
                End If
            Next
            If nValorAtualizado > CDbl(txtValor.Text) Then GoTo PROXIMO2
        End If
        If Option1(0).value = True Then
            sAno = .sFullLanc
        End If
        Sql = "INSERT DAM(COMPUTER,SEQ,CODREDUZIDO,ANOEXERC,LANC,NUMSEQ,NUMPARCELA,COMP,DATAVENCTO,FULLLANC,NOMECONTRIBUINTE,CPF,"
        Sql = Sql & "PRINCIPAL,CORRECAO,MULTA,JUROS,TOTAL) VALUES('" & NomeDeLogin & "'," & x & "," & .nCodReduz & "," & .nAno & "," & .nLanc & ","
        Sql = Sql & .nSeq & "," & .nParc & "," & .nCompl & ",'" & Format(.sVencto, "mm/dd/yyyy") & "','"
        Sql = Sql & sAno & "','" & Right$(Mask(.sNome), 40) & "','" & sCNPJ & "'," & Virg2Ponto(CStr(.nValorTributo)) & "," & Virg2Ponto(CStr(.nValorCorrecao)) & ","
        Sql = Sql & Virg2Ponto(CStr(.nValorMulta)) & "," & Virg2Ponto(CStr(.nValorJuros)) & "," & Virg2Ponto(CStr(.nValorAtual)) & ")"
    End With
    cn.Execute Sql, rdExecDirect
PROXIMO2:
Next


End Sub

Private Sub txtCod2_LostFocus()
txtCod2.Text = Format(Val(txtCod2.Text), "000000")
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If nPosF > 0 Then
    Pb.Color = &HC0C000
Else
    Pb.Color = vbWhite
End If

If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub txtTop_KeyPress(KeyAscii As Integer)
Tweak txtTop, KeyAscii, IntegerPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive, 2
End Sub

Private Sub GeraApenasTotais()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nCodReduz As Long, nTotal As Double, RdoAux3 As rdoResultset
Dim nTotReg As Long, nPosReg As Long, qd As New rdoQuery
Ocupado
Sql = "delete from mobiliariodevedor where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "select count(codigomob) as contador from mobiliario where dataencerramento is null"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nTotReg = RdoAux!contador
RdoAux.Close

nPosReg = 1
Sql = "select codigomob from mobiliario where dataencerramento is null order by codigomob"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       nTotal = 0
       'SUSPENÇÃO
        Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !codigomob & " ORDER BY DATAEVENTO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If !CODTIPOEVENTO = 2 Then
                    .Close
                    GoTo proximo
                End If
            End If
           .Close
        End With
    
        If nPosReg Mod 20 = 0 Then
            CallPb nPosReg, nTotReg
        End If
    
        nCodReduz = !codigomob
        
        Sql = "SELECT distinct debitoparcela.anoexercicio,debitoparcela.codlancamento,debitoparcela.seqlancamento,debitoparcela.numparcela,debitoparcela.codcomplemento FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
        Sql = Sql & "Where (debitoparcela.CODREDUZIDO = " & nCodReduz & ") And (debitoparcela.codlancamento <> 20) And (debitoparcela.datavencimento < '" & Format(Now, "mm/dd/yyyy") & "') And (debitoparcela.numparcela >0) And (debitoparcela.statuslanc = 3 or debitoparcela.statuslanc = 42 or debitoparcela.statuslanc = 43 ) And (debitotributo.ValorTributo > 0) "
        'Sql = Sql & " and debitoparcela.codlancamento in (2,3,5,14,65)"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        Do Until RdoAux2.EOF
            Set qd.ActiveConnection = cn
            qd.QueryTimeout = 0
            On Error Resume Next
            RdoAux3.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = RdoAux2!AnoExercicio
            qd(3) = RdoAux2!AnoExercicio
            qd(4) = RdoAux2!CodLancamento
            qd(5) = RdoAux2!CodLancamento
            qd(6) = RdoAux2!SeqLancamento
            qd(7) = RdoAux2!SeqLancamento
            qd(8) = RdoAux2!NumParcela
            qd(9) = RdoAux2!NumParcela
            qd(10) = RdoAux2!CODCOMPLEMENTO
            qd(11) = RdoAux2!CODCOMPLEMENTO
            qd(12) = 0
            qd(13) = 99
            qd(14) = Format(Now, "mm/dd/yyyy")
            qd(15) = NomeDeLogin
            Set RdoAux3 = qd.OpenResultset(rdOpenKeyset)
           ' If nCodReduz = 100076 Then MsgBox "teste"
            Do Until RdoAux3.EOF
                If !statuslanc = 3 Or !statuslanc = 42 Or !statuslanc = 43 Then
                    nTotal = nTotal + RdoAux3!ValorTotal
                End If
                RdoAux3.MoveNext
            Loop
            RdoAux3.Close
            RdoAux2.MoveNext
        Loop
        RdoAux2.Close
        
        If nTotal > 0 Then
            Sql = "insert mobiliariodevedor(usuario,codigo,valor) values('" & NomeDeLogin & "'," & nCodReduz & "," & Virg2Ponto(Format(nTotal, "#0.00")) & ")"
            cn.Execute Sql, rdExecDirect
        End If
proximo:
        DoEvents
        nPosReg = nPosReg + 1
       .MoveNext
    Loop
   .Close
End With

Liberado

frmReport.ShowReport2 "mobiliariodevedor", frmMdi.HWND, Me.HWND

Sql = "delete from mobiliariodevedor where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Pb.value = 0
End Sub
