VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCadastroRural 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Propriedade Rural"
   ClientHeight    =   5250
   ClientLeft      =   3090
   ClientTop       =   2160
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   8655
   Begin VB.Frame Pnl1 
      Height          =   4455
      Left            =   570
      TabIndex        =   9
      Top             =   7170
      Width           =   8625
      Begin VB.Frame Frame4 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Dados Gerais"
         ForeColor       =   &H00000080&
         Height          =   2325
         Left            =   30
         TabIndex        =   23
         Top             =   0
         Width           =   6165
         Begin VB.ComboBox cmbEstrada 
            Height          =   315
            ItemData        =   "frmCadastroRural.frx":0000
            Left            =   1305
            List            =   "frmCadastroRural.frx":0002
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   1935
            Width           =   4365
         End
         Begin VB.TextBox txtPropriedade 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   31
            Top             =   1590
            Width           =   4695
         End
         Begin VB.TextBox txtRecFed 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4380
            MaxLength       =   20
            TabIndex        =   25
            Top             =   270
            Width           =   1635
         End
         Begin VB.TextBox txtIE 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   26
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txtIncra 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   13
            TabIndex        =   24
            Top             =   270
            Width           =   1635
         End
         Begin prjChameleon.chameleonButton cmdCnsCid 
            Height          =   285
            Left            =   1320
            TabIndex        =   30
            ToolTipText     =   "Consulta Imóvel"
            Top             =   1260
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   503
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
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCadastroRural.frx":0004
            PICN            =   "frmCadastroRural.frx":0020
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdCRI 
            Height          =   255
            Left            =   4380
            TabIndex        =   27
            ToolTipText     =   "Novo Registro"
            Top             =   600
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   450
            BTYPE           =   3
            TX              =   "Visualizar ..."
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
            MICON           =   "frmCadastroRural.frx":017A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin esMaskEdit.esMaskedEdit mskCPF 
            Height          =   285
            Left            =   4380
            TabIndex        =   29
            Top             =   930
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            MouseIcon       =   "frmCadastroRural.frx":0196
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
            MaxLength       =   20
            SelText         =   ""
            HideSelection   =   -1  'True
         End
         Begin esMaskEdit.esMaskedEdit mskCNPJ 
            Height          =   285
            Left            =   1320
            TabIndex        =   28
            Top             =   930
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   503
            MouseIcon       =   "frmCadastroRural.frx":01B2
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
            MaxLength       =   20
            SelText         =   ""
            HideSelection   =   -1  'True
         End
         Begin prjChameleon.chameleonButton cmdRefresh3 
            Height          =   285
            Left            =   5715
            TabIndex        =   103
            ToolTipText     =   "Atualizar Lista"
            Top             =   1935
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   503
            BTYPE           =   3
            TX              =   "!"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            FCOL            =   192
            FCOLO           =   192
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCadastroRural.frx":01CE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Estrada.............:"
            Height          =   225
            Index           =   25
            Left            =   45
            TabIndex        =   102
            Top             =   1980
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CNPJ................:"
            Height          =   225
            Index           =   14
            Left            =   60
            TabIndex        =   43
            Top             =   990
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "CPF.................:"
            Height          =   225
            Index           =   13
            Left            =   3150
            TabIndex        =   42
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label lblProp 
            BackStyle       =   0  'Transparent
            Caption         =   "Proprietário.......:"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   1830
            TabIndex        =   41
            Top             =   1320
            Width           =   4125
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Matrícula CRI..:"
            Height          =   225
            Index           =   3
            Left            =   3150
            TabIndex        =   40
            Top             =   630
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cód.Rec.Fed...:"
            Height          =   225
            Index           =   2
            Left            =   3150
            TabIndex        =   39
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Propriedade......:"
            Height          =   225
            Index           =   5
            Left            =   60
            TabIndex        =   38
            Top             =   1650
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Proprietário.......:"
            Height          =   225
            Index           =   4
            Left            =   60
            TabIndex        =   37
            Top             =   1320
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Insc.Estadual....:"
            Height          =   225
            Index           =   1
            Left            =   60
            TabIndex        =   36
            Top             =   660
            Width           =   1305
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Inscr. no Incra...:"
            Height          =   225
            Index           =   0
            Left            =   60
            TabIndex        =   35
            Top             =   330
            Width           =   1305
         End
      End
      Begin VB.Frame pnlCRI 
         BorderStyle     =   0  'None
         Height          =   2925
         Left            =   855
         TabIndex        =   55
         Top             =   585
         Visible         =   0   'False
         Width           =   3495
         Begin VB.ListBox lstCri 
            Appearance      =   0  'Flat
            Height          =   2175
            Left            =   90
            TabIndex        =   57
            Top             =   240
            Width           =   3285
         End
         Begin VB.TextBox txtCri 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   90
            MaxLength       =   50
            TabIndex        =   56
            Top             =   2490
            Width           =   2295
         End
         Begin prjChameleon.chameleonButton cmdC2 
            Height          =   285
            Left            =   2940
            TabIndex        =   58
            ToolTipText     =   "Remove matrícula"
            Top             =   2490
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   503
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
            MICON           =   "frmCadastroRural.frx":01EA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdC1 
            Height          =   285
            Left            =   2490
            TabIndex        =   59
            ToolTipText     =   "Adiciona matrícula"
            Top             =   2490
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   503
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
            MICON           =   "frmCadastroRural.frx":0206
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Local"
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   30
         TabIndex        =   52
         Top             =   2385
         Width           =   6165
         Begin VB.TextBox txtReferencia 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   735
            Left            =   60
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   33
            Top             =   270
            Width           =   5985
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Endereço de Correspondência"
         ForeColor       =   &H00000080&
         Height          =   855
         Left            =   45
         TabIndex        =   51
         Top             =   3510
         Width           =   6165
         Begin VB.TextBox txtEndereco 
            Appearance      =   0  'Flat
            ForeColor       =   &H00000000&
            Height          =   465
            Left            =   60
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   270
            Width           =   5985
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Área da Propriedade"
         ForeColor       =   &H00000080&
         Height          =   1365
         Left            =   6210
         TabIndex        =   44
         Top             =   0
         Width           =   2295
         Begin VB.TextBox txtHa 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   900
            TabIndex        =   47
            Text            =   "0"
            Top             =   300
            Width           =   1275
         End
         Begin VB.TextBox txtAl 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   900
            TabIndex        =   46
            Text            =   "0"
            Top             =   630
            Width           =   1275
         End
         Begin VB.TextBox txtM2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   900
            TabIndex        =   45
            Text            =   "0"
            Top             =   960
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hectares.:"
            Height          =   225
            Index           =   8
            Left            =   90
            TabIndex        =   50
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Alqueires.:"
            Height          =   225
            Index           =   6
            Left            =   90
            TabIndex        =   49
            Top             =   690
            Width           =   765
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Metros(²)..:"
            Height          =   225
            Index           =   7
            Left            =   90
            TabIndex        =   48
            Top             =   1020
            Width           =   765
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Coordenadas Geográficas"
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   6210
         TabIndex        =   14
         Top             =   2070
         Width           =   2295
         Begin VB.TextBox txtCY2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   18
            Top             =   630
            Width           =   645
         End
         Begin VB.TextBox txtCX2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   17
            Top             =   300
            Width           =   645
         End
         Begin VB.TextBox txtCX1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   420
            TabIndex        =   16
            Top             =   300
            Width           =   645
         End
         Begin VB.TextBox txtCY1 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   420
            TabIndex        =   15
            Top             =   630
            Width           =   645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "(X2):"
            Height          =   225
            Index           =   11
            Left            =   1170
            TabIndex        =   22
            Top             =   360
            Width           =   405
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "(Y2):"
            Height          =   225
            Index           =   10
            Left            =   1170
            TabIndex        =   21
            Top             =   660
            Width           =   405
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "(Y1):"
            Height          =   225
            Index           =   12
            Left            =   60
            TabIndex        =   20
            Top             =   660
            Width           =   375
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "(X1):"
            Height          =   225
            Index           =   9
            Left            =   60
            TabIndex        =   19
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Área Agricultavel"
         ForeColor       =   &H00000080&
         Height          =   675
         Left            =   6210
         TabIndex        =   11
         Top             =   1380
         Width           =   2295
         Begin VB.TextBox txtAreaAgr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   900
            TabIndex        =   12
            Text            =   "0"
            Top             =   270
            Width           =   1275
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Hectares.:"
            Height          =   225
            Index           =   15
            Left            =   90
            TabIndex        =   13
            Top             =   330
            Width           =   765
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Código da Propriedade"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   24
         Left            =   6450
         TabIndex        =   54
         Top             =   3480
         Width           =   1965
      End
      Begin VB.Label lblCodReduzido 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   6450
         TabIndex        =   53
         Top             =   3810
         Width           =   1965
      End
   End
   Begin VB.Frame Pnl5 
      Height          =   4455
      Left            =   17640
      TabIndex        =   10
      Top             =   7980
      Width           =   8625
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   1
         Left            =   6450
         TabIndex        =   88
         Top             =   1140
         Width           =   1905
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   2
         Left            =   6450
         TabIndex        =   90
         Top             =   1455
         Width           =   1905
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   4
         Left            =   6450
         TabIndex        =   94
         Top             =   2100
         Width           =   1905
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   3
         Left            =   6450
         TabIndex        =   92
         Top             =   1785
         Width           =   1905
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   5
         Left            =   6450
         TabIndex        =   96
         Top             =   2850
         Width           =   1905
      End
      Begin VB.TextBox txtValor 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   6
         Left            =   6450
         TabIndex        =   98
         Top             =   3180
         Width           =   1905
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CÁLCULO DO VALOR DA TERRA NUA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   101
         Top             =   720
         Width           =   4155
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CÁLCULO DO IMPOSTO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   100
         Top             =   2490
         Width           =   2565
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "01. Valor total do Imóvel.................................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   99
         Top             =   1170
         Width           =   6255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "02. Valor das benfeitorias................................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   97
         Top             =   1485
         Width           =   6255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "03. Valor das culturas,past.cult.e melh.e florestas plant.:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   95
         Top             =   1815
         Width           =   6255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "04. Valor da terra nua....................................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   93
         Top             =   2130
         Width           =   6255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "05. Valor da terra nua tributavel.........................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   91
         Top             =   2880
         Width           =   6255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "06. Imposto devido........................................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   89
         Top             =   3195
         Width           =   6255
      End
   End
   Begin VB.Frame Pnl4 
      Height          =   4455
      Left            =   10200
      TabIndex        =   75
      Top             =   9060
      Width           =   8625
      Begin VB.TextBox txtProdA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3540
         TabIndex        =   83
         Text            =   "0"
         Top             =   210
         Width           =   1275
      End
      Begin VB.TextBox txtObsA 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   1260
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   82
         Top             =   3420
         Width           =   6945
      End
      Begin VB.ComboBox cmbCultivoA 
         Height          =   315
         Left            =   1290
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   2565
         Width           =   3915
      End
      Begin VB.TextBox txtAreaCA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         TabIndex        =   76
         Text            =   "0"
         Top             =   2940
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid grdA 
         Height          =   1755
         Left            =   480
         TabIndex        =   78
         Top             =   750
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3096
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "Cód   |<Descrição do Cultivo                                  |>Qtde. Utilizada   "
      End
      Begin prjChameleon.chameleonButton cmdAddA 
         Height          =   315
         Left            =   5730
         TabIndex        =   79
         ToolTipText     =   "Adicionar um Cultivo"
         Top             =   1800
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Adicionar"
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
         MICON           =   "frmCadastroRural.frx":0222
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelA 
         Height          =   315
         Left            =   5730
         TabIndex        =   80
         ToolTipText     =   "Remover um Cultivo"
         Top             =   2160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
         MICON           =   "frmCadastroRural.frx":023E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRefresh1 
         Height          =   285
         Left            =   5250
         TabIndex        =   81
         ToolTipText     =   "Atualizar Lista"
         Top             =   2580
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "!"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCadastroRural.frx":025A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Produção por arrendamento em hectares..:"
         Height          =   225
         Index           =   17
         Left            =   360
         TabIndex        =   87
         Top             =   240
         Width           =   3105
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
         Height          =   225
         Index           =   21
         Left            =   270
         TabIndex        =   86
         Top             =   3420
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cultivo....:"
         Height          =   225
         Index           =   23
         Left            =   510
         TabIndex        =   85
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantid...:"
         Height          =   225
         Index           =   22
         Left            =   480
         TabIndex        =   84
         Top             =   3000
         Width           =   765
      End
   End
   Begin VB.Frame Pnl2 
      Height          =   4455
      Left            =   10080
      TabIndex        =   60
      Top             =   90
      Width           =   8625
      Begin VB.TextBox txtHist 
         Appearance      =   0  'Flat
         Height          =   3945
         Left            =   240
         MaxLength       =   5000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   61
         Top             =   300
         Width           =   7695
      End
   End
   Begin vbalDTab6.vbalDTabControl dTab 
      Height          =   4785
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   8440
      AllowScroll     =   0   'False
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15658734
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   60
      TabIndex        =   7
      ToolTipText     =   "Novo Registro"
      Top             =   4860
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadastroRural.frx":0276
      PICN            =   "frmCadastroRural.frx":0292
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
      Height          =   315
      Left            =   1260
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   4860
      Width           =   1155
      _ExtentX        =   2037
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadastroRural.frx":03EC
      PICN            =   "frmCadastroRural.frx":0408
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
      Left            =   60
      TabIndex        =   1
      ToolTipText     =   "Gravar o Registro"
      Top             =   4860
      Width           =   1155
      _ExtentX        =   2037
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadastroRural.frx":0562
      PICN            =   "frmCadastroRural.frx":057E
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
      Left            =   1260
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   4860
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadastroRural.frx":0923
      PICN            =   "frmCadastroRural.frx":093F
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
      Left            =   2460
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   4860
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadastroRural.frx":0A99
      PICN            =   "frmCadastroRural.frx":0AB5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   6270
      TabIndex        =   4
      ToolTipText     =   "Consulta Imóveis Cadastrados"
      Top             =   4860
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consultar"
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
      MICON           =   "frmCadastroRural.frx":0B57
      PICN            =   "frmCadastroRural.frx":0B73
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
      Left            =   7380
      TabIndex        =   5
      ToolTipText     =   "Sair da Tela"
      Top             =   4860
      Width           =   1065
      _ExtentX        =   1879
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
      MICON           =   "frmCadastroRural.frx":0CCD
      PICN            =   "frmCadastroRural.frx":0CE9
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
      Left            =   3660
      TabIndex        =   6
      ToolTipText     =   "Imprime os Dados do Cadastro"
      Top             =   4860
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCadastroRural.frx":0D57
      PICN            =   "frmCadastroRural.frx":0D73
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Pnl3 
      Height          =   4455
      Left            =   10800
      TabIndex        =   62
      Top             =   900
      Width           =   8625
      Begin VB.ComboBox cmbCultivoP 
         Height          =   315
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   2580
         Width           =   3915
      End
      Begin VB.TextBox txtAreaCP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   65
         Text            =   "0"
         Top             =   2940
         Width           =   1275
      End
      Begin VB.TextBox txtObsP 
         Appearance      =   0  'Flat
         Height          =   825
         Left            =   1140
         MaxLength       =   500
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   64
         Top             =   3450
         Width           =   6945
      End
      Begin VB.TextBox txtProdP 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   63
         Text            =   "0"
         Top             =   240
         Width           =   1275
      End
      Begin MSFlexGridLib.MSFlexGrid grdP 
         Height          =   1755
         Left            =   270
         TabIndex        =   67
         Top             =   750
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   3096
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   15658734
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         FormatString    =   "Cód   |<Descrição do Cultivo                                  |>Qtde. Utilizada   "
      End
      Begin prjChameleon.chameleonButton cmdAddP 
         Height          =   315
         Left            =   5520
         TabIndex        =   68
         ToolTipText     =   "Adicionar um Cultivo"
         Top             =   1800
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Adicionar"
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
         MICON           =   "frmCadastroRural.frx":0ECD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelP 
         Height          =   315
         Left            =   5520
         TabIndex        =   69
         ToolTipText     =   "Remover um Cultivo"
         Top             =   2160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
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
         MICON           =   "frmCadastroRural.frx":0EE9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRefresh2 
         Height          =   285
         Left            =   5040
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Atualizar Lista"
         Top             =   2580
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         BTYPE           =   3
         TX              =   "!"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCadastroRural.frx":0F05
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cultivo....:"
         Height          =   225
         Index           =   19
         Left            =   300
         TabIndex        =   74
         Top             =   2640
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantid...:"
         Height          =   225
         Index           =   18
         Left            =   270
         TabIndex        =   73
         Top             =   3000
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Observação:"
         Height          =   225
         Index           =   20
         Left            =   150
         TabIndex        =   72
         Top             =   3450
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Produção própria em hectares..:"
         Height          =   225
         Index           =   16
         Left            =   240
         TabIndex        =   71
         Top             =   270
         Width           =   2325
      End
   End
End
Attribute VB_Name = "frmCadastroRural"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tCampo
    nCodigo As Integer
    sNome As String
End Type

Private Type tEdit
    sAreaProp As String
    sAreaAgric As String
    sValor1 As String
    sValor2 As String
    sValor3 As String
    sValor4 As String
    sValor5 As String
    sValor6 As String
    sOldDesc As String
    sNewDesc As String
End Type

Dim RdoAux As rdoResultset, aEdit() As tEdit, aCampo() As tCampo
Dim Sql As String, sEvento As String

Private Sub cmdAddA_Click()
Dim x As Integer, bAchou As Boolean

If cmbCultivoA.ListIndex = -1 Then
    MsgBox "Selecione um tipo de cultivo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAreaCA.Text) = 0 Then
    MsgBox "Digite a área do cultivo.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
With grdA
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 0) = cmbCultivoA.ItemData(cmbCultivoA.ListIndex) Then
            bAchou = True
            Exit For
        End If
    Next
End With

If bAchou Then
    MsgBox "Este cultivo já foi cadastrado.", vbExclamation, "Atenção"
    Exit Sub
End If

grdA.AddItem cmbCultivoA.ItemData(cmbCultivoA.ListIndex) & Chr(9) & cmbCultivoA.Text & Chr(9) & txtAreaCA.Text

End Sub

Private Sub cmdAddP_Click()

Dim x As Integer, bAchou As Boolean

If cmbCultivoP.ListIndex = -1 Then
    MsgBox "Selecione um tipo de cultivo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAreaCP.Text) = 0 Then
    MsgBox "Digite a área do cultivo.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
With grdP
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 0) = cmbCultivoP.ItemData(cmbCultivoP.ListIndex) Then
            bAchou = True
            Exit For
        End If
    Next
End With

If bAchou Then
    MsgBox "Este cultivo já foi cadastrado.", vbExclamation, "Atenção"
    Exit Sub
End If

grdP.AddItem cmbCultivoP.ItemData(cmbCultivoP.ListIndex) & Chr(9) & cmbCultivoP.Text & Chr(9) & txtAreaCP.Text

End Sub

Private Sub cmdAlterar_Click()
If Val(txtIncra.Text) = 0 Then
    MsgBox "Selecione uma propriedade.", vbExclamation, "Atenção"
    Exit Sub
End If

sEvento = "Alterar"
Eventos "INCLUIR"
End Sub

Private Sub cmdC1_Click()
Dim p As Integer, bAchou As Boolean

If Trim(txtCri.Text) = "" Then
    MsgBox "Digite um nº de matrícula.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
For p = 0 To lstCri.ListCount - 1
    If lstCri.List(p) = txtCri.Text Then
        bAchou = True
    End If
Next

If bAchou Then
    MsgBox "Matrícula já cadastrada.", vbExclamation, "Atenção"
    Exit Sub
End If

lstCri.AddItem txtCri.Text

End Sub

Private Sub cmdC2_Click()
If lstCri.ListIndex = -1 Then Exit Sub

If MsgBox("Remover a matrícula " & lstCri.Text & " ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
    lstCri.RemoveItem (lstCri.ListIndex)
End If

End Sub

Private Sub cmdCancel_Click()
sEvento = ""
Eventos "INICIAR"
End Sub

Private Sub cmdCnsCid_Click()
Set frm = frmCnsCidadao
frm.sForm = "frmCadastroRural"
frm.show
frm.ZOrder 0

End Sub

Private Sub cmdConsultar_Click()
frmConsultaRural.show
frmConsultaRural.ZOrder 0
End Sub

Private Sub cmdCRI_Click()
pnlCRI.Visible = Not pnlCRI.Visible
End Sub

Private Sub cmdDelA_Click()

With grdA
    If .Rows = 1 Then
        MsgBox "Nenhum cultivo cadastrado.", vbExclamation, "Atenção"
    Else
        If .Row < 1 Then
            MsgBox "Selecione um cultivo.", vbExclamation, "Atenção"
        Else
            If .Rows <= 2 Then
                .Rows = 1
            Else
                .RemoveItem (.Row)
            End If
        End If
    End If
End With

End Sub

Private Sub cmdDelP_Click()

With grdP
    If .Rows = 1 Then
        MsgBox "Nenhum cultivo cadastrado.", vbExclamation, "Atenção"
    Else
        If .Row < 1 Then
            MsgBox "Selecione um cultivo.", vbExclamation, "Atenção"
        Else
            If .Rows <= 2 Then
                .Rows = 1
            Else
                .RemoveItem (.Row)
            End If
        End If
    End If
End With

End Sub

Private Sub cmdExcluir_Click()
If Val(txtIncra.Text) = 0 Then
    MsgBox "Selecione uma propriedade.", vbExclamation, "Atenção"
    Exit Sub
End If
If MsgBox("Excluir esta propridade do sistema ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "DELETE FROM CADASTRORURALPRODUTO WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM CADASTRORURALMAT WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM CADASTRORURALHIST WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM CADASTRORURAL WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
    cn.Execute Sql, rdExecDirect
    Log Form, Me.Caption, Exclusão, "Excluído registro " & Val(lblCodReduzido.Caption)
    Limpa
End If

End Sub

Private Sub cmdGravar_Click()

If Val(txtIncra.Text) = 0 Then
    MsgBox "Digite a Incrição do Incra.", vbExclamation, "Atenção"
    Exit Sub
End If

If Len(txtIncra.Text) <> 13 Then
    MsgBox "Incrição do Incra deve ter 13 digitos.", vbExclamation, "Atenção"
    Exit Sub
End If

'If Val(txtRecFed.text) = 0 Then
'    txtRecFed.text = 0
'End If

If Val(txtM2.Text) = 0 Then
    MsgBox "Digite a área da Propriedade.", vbExclamation, "Atenção"
    Exit Sub
End If

If lblProp.Caption = "" Then
    MsgBox "Selecione o proprietário.", vbExclamation, "Atenção"
    Exit Sub
End If

If sEvento = "Novo" Then
    Sql = "SELECT INCRA FROM CADASTRORURAL WHERE INCRA=" & Val(txtIncra.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Inscrição no Incra já cadastrada.", vbExclamation, "Atenção"
            Exit Sub
        End If
       .Close
    End With
End If

Grava
sEvento = ""
Eventos "INICIAR"

End Sub

Private Sub cmdNovo_Click()
Limpa
sEvento = "Novo"
Eventos "INCLUIR"
End Sub


Private Sub cmdPrint_Click()
If Val(txtIncra.Text) = 0 Then
    MsgBox "Selecione uma propriedade.", vbExclamation, "Atenção"
    Exit Sub
End If
frmReport.ShowReport "CADASTRORURAL", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub cmdRefresh3_Click()
CarregaEstrada
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Dim nOldCod As Double
If CodRural > 0 And sEvento = "" Then
    nOldCod = Val(lblCodReduzido.Caption)
    Limpa
    lblCodReduzido.Caption = Format(nOldCod, "000000")
    Le
End If

End Sub

Private Sub Form_Load()
Dim c As cTab, Sql As String, RdoAux As rdoResultset

ReDim aCampo(0)
Sql = "SELECT * FROM EVRURALCAMPO ORDER BY CODIGO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCampo(UBound(aCampo) + 1)
        aCampo(UBound(aCampo)).nCodigo = !Codigo
        aCampo(UBound(aCampo)).sNome = !Nome
       .MoveNext
    Loop
   .Close
End With

With dTab
    .ShowCloseButton = False
    Set c = .Tabs.Add("Tab1", , "Dados Gerais")
    c.Panel = Pnl1
    Set c = .Tabs.Add("Tab2", , "Histórico")
    c.Panel = Pnl2
    Set c = .Tabs.Add("Tab3", , "Produção Própria")
    c.Panel = Pnl3
    Set c = .Tabs.Add("Tab4", , "Produção por Arrendamento")
    c.Panel = Pnl4
    Set c = .Tabs.Add("Tab5", , "Valor da Terra")
    c.Panel = Pnl5
End With
CarregaLista
CarregaEstrada
Centraliza Me
Eventos "INICIAR"

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   cmdConsultar.Visible = True
   cmdCnsCid.Enabled = False
   For Each Ct In frmCadastroRural
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = Kde
          Ct.Locked = True
       End If
   Next
   lstCri.BackColor = Kde
   cmdC1.Enabled = False
   cmdC2.Enabled = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   cmdConsultar.Visible = False
   cmdCnsCid.Enabled = True
   For Each Ct In frmCadastroRural
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = vbWhite
          Ct.Locked = False
       End If
   Next
   lstCri.BackColor = Branco
   cmdC1.Enabled = True
   cmdC2.Enabled = True
   pnlCRI.Visible = False
End If

End Sub

Private Sub Limpa()

lblProp.Caption = ""
For Each Ct In frmCadastroRural
    If TypeOf Ct Is TextBox Then
       Ct.Text = ""
    End If
Next
cmbEstrada.ListIndex = -1
LimpaMascara mskCPF
LimpaMascara mskCNPJ
cmbCultivoA.ListIndex = -1
cmbCultivoP.ListIndex = -1
lblCodReduzido.Caption = "000000"
grdA.Rows = 1: grdP.Rows = 1
txtHa.Text = 0: txtAl.Text = 0: txtM2.Text = 0: txtAreaAgr.Text = 0
lstCri.Clear
End Sub

Private Sub lstCri_Click()
If lstCri.ListIndex > -1 Then
    txtCri.Text = lstCri.Text
Else
    txtCri.Text = ""
End If
End Sub

Private Sub txtAl_KeyPress(KeyAscii As Integer)
Tweak txtAl, KeyAscii, DecimalPositive, 4
End Sub

Private Sub txtAl_LostFocus()

If Trim$(txtAl.Text) = "" Then txtAl.Text = 0
   
txtHa.Text = FormatNumber(txtAl.Text * 2.42, 4)
txtM2.Text = FormatNumber(txtAl.Text * 24200, 4)

End Sub

Private Sub txtAreaAgr_KeyPress(KeyAscii As Integer)
Tweak txtAreaAgr, KeyAscii, DecimalPositive, 4
End Sub

Private Sub txtAreaCA_KeyPress(KeyAscii As Integer)
Tweak txtAreaCA, KeyAscii, DecimalPositive, 4
End Sub

Private Sub txtAreaCP_KeyPress(KeyAscii As Integer)
Tweak txtAreaCP, KeyAscii, DecimalPositive, 4
End Sub

'Private Sub txtCri_KeyPress(KeyAscii As Integer)
'Tweak txtCri, KeyAscii, IntegerPositive
'End Sub

Private Sub txtCX1_KeyPress(KeyAscii As Integer)
Tweak txtCX1, KeyAscii, DecimalPositive
End Sub

Private Sub txtCY1_KeyPress(KeyAscii As Integer)
Tweak txtCY1, KeyAscii, DecimalPositive
End Sub

Private Sub txtCX2_KeyPress(KeyAscii As Integer)
Tweak txtCX2, KeyAscii, DecimalPositive
End Sub

Private Sub txtCY2_KeyPress(KeyAscii As Integer)
Tweak txtCY2, KeyAscii, DecimalPositive
End Sub

Private Sub txtHa_KeyPress(KeyAscii As Integer)
Tweak txtHa, KeyAscii, DecimalPositive, 4
End Sub

Private Sub txtHa_LostFocus()
If Trim$(txtHa.Text) = "" Then txtHa.Text = 0
   
txtAl.Text = FormatNumber(txtHa.Text / 2.42, 4)
txtM2.Text = FormatNumber(txtHa.Text * 10000, 4)

End Sub

Private Sub txtM2_KeyPress(KeyAscii As Integer)
Tweak txtM2, KeyAscii, DecimalPositive, 4
End Sub

Private Sub txtM2_LostFocus()
If Trim$(txtHa.Text) = "" Then txtHa.Text = 0
   
txtAl.Text = FormatNumber(txtM2.Text / 24200, 4)
txtHa.Text = FormatNumber(txtM2.Text / 10000, 4)

End Sub

Private Sub Grava()
Dim p As Integer, MinCod As Long, MaxCod As Long, nCodEstrada As Integer

If sEvento = "Novo" Then
    
    'retorna os buracos
    Sql = "SELECT CODREDUZIDO FROM CADASTRORURAL ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            Do Until .EOF
               If MinCod = 0 Then
                  MinCod = !CODREDUZIDO
               Else
                  MaxCod = !CODREDUZIDO
                  If MaxCod - MinCod > 1 Then
                      MaxCod = MinCod + 1
                      Exit Do
                  Else
                      MinCod = MaxCod
                  End If
               End If
              .MoveNext
            Loop
        Else
            MaxCod = 800000
        End If
       .Close
    End With
    Sql = "SELECT CODREDUZIDO FROM CADASTRORURAL WHERE CODREDUZIDO=" & MaxCod
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        MaxCod = MaxCod + 1
        RdoAux.Close
    End If
    
Else
    MaxCod = lblCodReduzido.Caption
End If

If txtCX1.Text = "" Then txtCX1.Text = "0"
If txtCY1.Text = "" Then txtCY1.Text = "0"
If txtCX2.Text = "" Then txtCX2.Text = "0"
If txtCY2.Text = "" Then txtCY2.Text = "0"
If txtProdP.Text = "" Then txtProdP.Text = "0"
If txtProdA.Text = "" Then txtProdA.Text = "0"
For x = 1 To 6
    If Trim$(txtValor(x).Text) = "" Then
        txtValor(x).Text = "0"
    End If
Next

If cmbEstrada.ListIndex = -1 Then
    nCodEstrada = -1
Else
    nCodEstrada = cmbEstrada.ItemData(cmbEstrada.ListIndex)
End If

If sEvento = "Novo" Then
    'Sql = "INSERT CADASTRORURAL(CODREDUZIDO,INCRA,RECFED,DECAP,INSCPRODUTOR,CPF,PROPRIETARIO,PROPRIEDADE,REFERENCIA,"
    Sql = "INSERT CADASTRORURAL(CODREDUZIDO,INCRA,RECFED,CPF,PROPRIETARIO,PROPRIEDADE,REFERENCIA,"
    Sql = Sql & "ENDERECO,HECTARE,ALQUEIRE,METRO,AREAAGRICULTAVEL,AREAPROPRIA,AREAARRENDADA,OBSPROPRIA,OBSARRENDADA,"
    Sql = Sql & "COORDX1,COORDY1,COORDX2,COORDY2,VALOR1,VALOR2,VALOR3,VALOR4,VALOR5,VALOR6,IE,CNPJ,DATAINCLUSAO,CODESTRADA) VALUES(" & MaxCod & "," & Val(txtIncra.Text) & ",'" & Mask(txtRecFed.Text) & "',"
    Sql = Sql & IIf(mskCPF.ClipText = "", "Null", "'" & mskCPF.ClipText & "'") & "," & Val(Left$(lblProp.Caption, 6)) & "," & IIf(txtPropriedade.Text = "", "Null", "'" & Mask(txtPropriedade.Text) & "'") & "," & IIf(txtReferencia.Text = "", "Null", "'" & Mask(txtReferencia.Text) & "'") & ","
    Sql = Sql & IIf(txtEndereco.Text = "", "Null", "'" & Mask(txtEndereco.Text) & "'") & "," & Virg2Ponto(RemovePonto(txtHa.Text)) & "," & Virg2Ponto(RemovePonto(txtAl.Text)) & "," & Virg2Ponto(RemovePonto(txtM2.Text)) & "," & Virg2Ponto(txtAreaAgr.Text) & "," & Virg2Ponto(txtProdP.Text) & ","
    Sql = Sql & Virg2Ponto(txtProdA.Text) & "," & IIf(txtObsP.Text = "", "Null", "'" & Mask(txtObsP.Text) & "'") & "," & IIf(txtObsA.Text = "", "Null", "'" & Mask(txtObsA.Text) & "'") & "," & Virg2Ponto(RemovePonto(txtCX1.Text)) & "," & Virg2Ponto(RemovePonto(txtCY1.Text)) & "," & Virg2Ponto(RemovePonto(txtCX2.Text)) & "," & Virg2Ponto(RemovePonto(txtCY2.Text)) & ","
    Sql = Sql & Virg2Ponto(txtValor(1).Text) & "," & Virg2Ponto(txtValor(2).Text) & "," & Virg2Ponto(txtValor(3).Text) & "," & Virg2Ponto(txtValor(4).Text) & "," & Virg2Ponto(txtValor(5).Text) & "," & Virg2Ponto(txtValor(6).Text) & ",'"
    Sql = Sql & txtIE.Text & "'," & IIf(mskCNPJ.ClipText = "", "Null", "'" & mskCNPJ.ClipText & "'") & ",'" & Format(Now, "mm/dd/yyyy") & "'," & nCodEstrada & ")"
Else
    Sql = "UPDATE CADASTRORURAL SET INCRA=" & Val(txtIncra.Text) & ",RECFED='" & Mask(txtRecFed.Text) & "',"
    Sql = Sql & "IE='" & Mask(txtIE.Text) & "',CNPJ=" & IIf(mskCNPJ.ClipText = "", "Null", "'" & mskCNPJ.ClipText & "'") & ",CPF=" & IIf(mskCPF.ClipText = "", "Null", "'" & mskCPF.ClipText & "'") & ","
    Sql = Sql & "PROPRIETARIO=" & Val(Left$(lblProp.Caption, 6)) & ",PROPRIEDADE=" & IIf(txtPropriedade.Text = "", "Null", "'" & Mask(txtPropriedade.Text) & "'") & ","
    Sql = Sql & "REFERENCIA=" & IIf(txtReferencia.Text = "", "Null", "'" & Mask(txtReferencia.Text) & "'") & ",ENDERECO=" & IIf(txtEndereco.Text = "", "Null", "'" & Mask(txtEndereco.Text) & "'") & ","
    Sql = Sql & "HECTARE=" & Virg2Ponto(RemovePonto(txtHa.Text)) & ",ALQUEIRE=" & Virg2Ponto(RemovePonto(txtAl.Text)) & ","
    Sql = Sql & "METRO=" & Virg2Ponto(RemovePonto(txtM2.Text)) & ",AREAAGRICULTAVEL=" & Virg2Ponto(RemovePonto(txtAreaAgr.Text)) & ",AREAPROPRIA=" & Virg2Ponto(RemovePonto(txtProdP.Text)) & ",AREAARRENDADA=" & Virg2Ponto(RemovePonto(txtProdA.Text)) & ",OBSPROPRIA=" & IIf(txtObsP.Text = "", "Null", "'" & Mask(txtObsP.Text) & "'") & ","
    Sql = Sql & "OBSARRENDADA=" & IIf(txtObsA.Text = "", "Null", "'" & Mask(txtObsA.Text) & "'") & ",COORDX1=" & Virg2Ponto(RemovePonto(txtCX1.Text)) & ","
    Sql = Sql & "COORDY1=" & Virg2Ponto(RemovePonto(txtCY1.Text)) & ",COORDX2=" & Virg2Ponto(RemovePonto(txtCX2.Text)) & ",COORDY2=" & Virg2Ponto(RemovePonto(txtCY2.Text)) & ","
    Sql = Sql & "VALOR1=" & Virg2Ponto(RemovePonto(txtValor(1).Text)) & ",VALOR2=" & Virg2Ponto(RemovePonto(txtValor(2).Text)) & ","
    Sql = Sql & "VALOR3=" & Virg2Ponto(RemovePonto(txtValor(3).Text)) & ",VALOR4=" & Virg2Ponto(RemovePonto(txtValor(4).Text)) & ","
    Sql = Sql & "VALOR5=" & Virg2Ponto(RemovePonto(txtValor(5).Text)) & ",VALOR6=" & Virg2Ponto(RemovePonto(txtValor(6).Text)) & ",CODESTRADA=" & nCodEstrada
    Sql = Sql & " WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
End If

cn.Execute Sql, rdExecDirect
lblCodReduzido.Caption = MaxCod

Sql = "DELETE FROM CADASTRORURALMAT WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
cn.Execute Sql, rdExecDirect

For p = 0 To lstCri.ListCount - 1
    Sql = "INSERT CADASTRORURALMAT(CODREDUZIDO,MATRICULA) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Mask(lstCri.List(p)) & "')"
    cn.Execute Sql, rdExecDirect
Next

Sql = "DELETE FROM CADASTRORURALHIST WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
cn.Execute Sql, rdExecDirect

If Trim$(txtHist.Text) <> "" Then
    Sql = "INSERT CADASTRORURALHIST(CODREDUZIDO,HISTORICO) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Mask(txtHist.Text) & "')"
    cn.Execute Sql, rdExecDirect
End If

Sql = "DELETE FROM CADASTRORURALPRODUTO WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
cn.Execute Sql, rdExecDirect

With grdP
    For x = 1 To .Rows - 1
        Sql = "INSERT CADASTRORURALPRODUTO(CODREDUZIDO,TIPO,CODPRODUTO,AREAPRODUTO) VALUES("
        Sql = Sql & lblCodReduzido.Caption & ",'P'," & .TextMatrix(x, 0) & "," & Virg2Ponto(.TextMatrix(x, 2)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

With grdA
    For x = 1 To .Rows - 1
        Sql = "INSERT CADASTRORURALPRODUTO(CODREDUZIDO,TIPO,CODPRODUTO,AREAPRODUTO) VALUES("
        Sql = Sql & lblCodReduzido.Caption & ",'A'," & .TextMatrix(x, 0) & "," & Virg2Ponto(.TextMatrix(x, 2)) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

'*** LOG ***
If sEvento <> "Novo" Then
    If aEdit(0).sAreaProp <> txtHa.Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',1,'" & aEdit(0).sAreaProp & "','" & txtHa.Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sAreaAgric <> txtAreaAgr.Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',2,'" & aEdit(0).sAreaAgric & "','" & txtAreaAgr.Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sValor1 <> txtValor(1).Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',3,'" & aEdit(0).sValor1 & "','" & txtValor(1).Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sValor2 <> txtValor(2).Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',4,'" & aEdit(0).sValor2 & "','" & txtValor(2).Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sValor3 <> txtValor(3).Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',5,'" & aEdit(0).sValor3 & "','" & txtValor(3).Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sValor4 <> txtValor(4).Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',6,'" & aEdit(0).sValor4 & "','" & txtValor(4).Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sValor5 <> txtValor(5).Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',7,'" & aEdit(0).sValor5 & "','" & txtValor(5).Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
    If aEdit(0).sValor6 <> txtValor(6).Text Then
        Sql = "INSERT EVRURALEDIT(CODREDUZIDO,DATAEDIT,USUARIO,TIPO,OLDDESC,NEWDESC) VALUES(" & Val(lblCodReduzido.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & NomeDeLogin & "',8,'" & aEdit(0).sValor6 & "','" & txtValor(6).Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
End If


End Sub

Private Sub Le()
Dim RdoAux2 As rdoResultset

ReDim aEdit(0)

If Val(lblCodReduzido.Caption) = 0 Then Exit Sub
Sql = "SELECT * FROM CADASTRORURAL WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    aEdit(0).sAreaProp = FormatNumber(!HECTARE, 4)
    aEdit(0).sAreaAgric = FormatNumber(!AREAAGRICULTAVEL, 4)
    aEdit(0).sValor1 = FormatNumber(!VALOR1, 2)
    aEdit(0).sValor2 = FormatNumber(!VALOR2, 2)
    aEdit(0).sValor3 = FormatNumber(!VALOR3, 2)
    aEdit(0).sValor4 = FormatNumber(!VALOR4, 2)
    aEdit(0).sValor5 = FormatNumber(!VALOR5, 2)
    aEdit(0).sValor6 = FormatNumber(!VALOR6, 2)

    txtIncra.Text = !INCRA
    txtRecFed.Text = SubNull(!RECFED)
    txtIE.Text = SubNull(!IE)
    mskCNPJ.Text = SubNull(!Cnpj)
    mskCPF.Text = SubNull(!CPF)
    txtPropriedade.Text = SubNull(!PROPRIEDADE)
    txtReferencia.Text = SubNull(!REFERENCIA)
    txtEndereco.Text = SubNull(!Endereco)
    txtHa.Text = FormatNumber(!HECTARE, 4)
    txtAl.Text = FormatNumber(!ALQUEIRE, 4)
    txtM2.Text = FormatNumber(!METRO, 4)
    txtAreaAgr.Text = FormatNumber(!AREAAGRICULTAVEL, 4)
    txtProdP.Text = FormatNumber(!AREAPROPRIA, 4)
    txtProdA.Text = FormatNumber(!AREAARRENDADA, 4)
    txtCX1.Text = Val(SubNull(!COORDX1))
    txtCY1.Text = Val(SubNull(!COORDY1))
    txtCX2.Text = Val(SubNull(!COORDX2))
    txtCY2.Text = Val(SubNull(!COORDY2))
    txtValor(1).Text = FormatNumber(!VALOR1, 2)
    txtValor(2).Text = FormatNumber(!VALOR2, 2)
    txtValor(3).Text = FormatNumber(!VALOR3, 2)
    txtValor(4).Text = FormatNumber(!VALOR4, 2)
    txtValor(5).Text = FormatNumber(!VALOR5, 2)
    txtValor(6).Text = FormatNumber(!VALOR6, 2)
    txtObsA.Text = SubNull(!OBSARRENDADA)
    txtObsP.Text = SubNull(!OBSPROPRIA)
    For x = 0 To cmbEstrada.ListCount - 1
        If cmbEstrada.ItemData(x) = !codestrada Then
           cmbEstrada.ListIndex = x
           Exit For
        End If
    Next

    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !Proprietario
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            lblProp.Caption = Format(RdoAux!Proprietario, "000000") & " - " & !nomecidadao
        Else
            lblProp.Caption = ""
        End If
       .Close
    End With
   .Close
End With

Sql = "SELECT MATRICULA FROM CADASTRORURALMAT WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstCri.AddItem !MATRICULA
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT HISTORICO FROM CADASTRORURALHIST WHERE CODREDUZIDO=" & Val(lblCodReduzido.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtHist.Text = !HISTORICO
    End If
   .Close
End With

Sql = "SELECT  cadastroruralproduto.codproduto, produtorural.nomeproduto, cadastroruralproduto.areaproduto "
Sql = Sql & "FROM  cadastroruralproduto INNER JOIN  produtorural ON cadastroruralproduto.codproduto = produtorural.codproduto "
Sql = Sql & "WHERE cadastroruralproduto.tipo = 'P' AND cadastroruralproduto.codreduzido = " & Val(lblCodReduzido.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdP.AddItem !CODPRODUTO & Chr(9) & !NOMEPRODUTO & Chr(9) & !AREAPRODUTO
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT  cadastroruralproduto.codproduto, produtorural.nomeproduto, cadastroruralproduto.areaproduto "
Sql = Sql & "FROM  cadastroruralproduto INNER JOIN  produtorural ON cadastroruralproduto.codproduto = produtorural.codproduto "
Sql = Sql & "WHERE cadastroruralproduto.tipo = 'A' AND cadastroruralproduto.codreduzido = " & Val(lblCodReduzido.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdA.AddItem !CODPRODUTO & Chr(9) & !NOMEPRODUTO & Chr(9) & !AREAPRODUTO
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub CarregaLista()

Sql = "SELECT CODPRODUTO,NOMEPRODUTO FROM PRODUTORURAL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbCultivoA.AddItem !NOMEPRODUTO
        cmbCultivoP.AddItem !NOMEPRODUTO
        cmbCultivoA.ItemData(cmbCultivoA.NewIndex) = !CODPRODUTO
        cmbCultivoP.ItemData(cmbCultivoP.NewIndex) = !CODPRODUTO
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaEstrada()

Sql = "SELECT CODIGO,NOME FROM ESTRADARURAL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbEstrada.AddItem !Nome
        cmbEstrada.ItemData(cmbEstrada.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub



Private Sub txtProdA_KeyPress(KeyAscii As Integer)
Tweak txtProdA, KeyAscii, DecimalPositive
End Sub

Private Sub txtProdP_KeyPress(KeyAscii As Integer)
Tweak txtProdP, KeyAscii, DecimalPositive
End Sub

Private Sub txtValor_KeyPress(Index As Integer, KeyAscii As Integer)
Tweak txtValor(Index), KeyAscii, DecimalPositive
End Sub
