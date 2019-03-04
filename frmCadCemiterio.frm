VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCadCemiterio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro do cemitério"
   ClientHeight    =   5760
   ClientLeft      =   12150
   ClientTop       =   8460
   ClientWidth     =   9225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5760
   ScaleWidth      =   9225
   Begin VB.Frame Frame4 
      Caption         =   "D) Procedimentos Autorizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   45
      TabIndex        =   49
      Top             =   4590
      Width           =   7755
      Begin VB.CheckBox chkProc 
         Caption         =   "Exumação de Carneiro"
         Height          =   195
         Index           =   7
         Left            =   4950
         TabIndex        =   57
         Top             =   855
         Width           =   2040
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Abert. e Fech. de Carneiro"
         Height          =   195
         Index           =   6
         Left            =   4950
         TabIndex        =   56
         Top             =   585
         Width           =   2625
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Concessão de Carneiro"
         Height          =   195
         Index           =   5
         Left            =   4950
         TabIndex        =   55
         Top             =   315
         Width           =   2040
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Jazigo Solidário (menor)"
         Height          =   195
         Index           =   4
         Left            =   2565
         TabIndex        =   54
         Top             =   585
         Width           =   2040
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Jazigo Solidário (adulto)"
         Height          =   195
         Index           =   3
         Left            =   2565
         TabIndex        =   53
         Top             =   315
         Width           =   2040
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Exumação de Jazigo"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   52
         Top             =   855
         Width           =   2040
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Abert. e fech. de Jazigo"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   51
         Top             =   585
         Width           =   2040
      End
      Begin VB.CheckBox chkProc 
         Caption         =   "Concessão de Jazigo"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   50
         Top             =   315
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "C) Outros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   45
      TabIndex        =   27
      Top             =   2700
      Width           =   7755
      Begin VB.CheckBox chkSep 
         Caption         =   "Jazigo solidário criança"
         Height          =   195
         Index           =   5
         Left            =   3735
         TabIndex        =   48
         Top             =   1260
         Width           =   2040
      End
      Begin VB.CheckBox chkSep 
         Caption         =   "Jazigo solidário adulto"
         Height          =   195
         Index           =   4
         Left            =   1665
         TabIndex        =   47
         Top             =   1260
         Width           =   2040
      End
      Begin VB.CheckBox chkSep 
         Caption         =   "Jazigo 6 gavetas"
         Height          =   195
         Index           =   3
         Left            =   5940
         TabIndex        =   46
         Top             =   990
         Width           =   1590
      End
      Begin VB.CheckBox chkSep 
         Caption         =   "Jazigo 4 gavetas"
         Height          =   195
         Index           =   2
         Left            =   4275
         TabIndex        =   45
         Top             =   990
         Width           =   1590
      End
      Begin VB.CheckBox chkSep 
         Caption         =   "Jazigo 3 gavetas"
         Height          =   195
         Index           =   1
         Left            =   2655
         TabIndex        =   44
         Top             =   990
         Width           =   1590
      End
      Begin VB.CheckBox chkSep 
         Caption         =   "Carneiro"
         Height          =   195
         Index           =   0
         Left            =   1665
         TabIndex        =   43
         Top             =   990
         Width           =   960
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   6210
         MaxLength       =   6
         TabIndex        =   41
         Top             =   1530
         Width           =   675
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   4095
         MaxLength       =   6
         TabIndex        =   39
         Top             =   1530
         Width           =   675
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   37
         Top             =   1530
         Width           =   675
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   5040
         MaxLength       =   6
         TabIndex        =   33
         Top             =   585
         Width           =   675
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1665
         TabIndex        =   31
         Text            =   "Combo1"
         Top             =   585
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1665
         TabIndex        =   29
         Top             =   225
         Width           =   5920
      End
      Begin esMaskEdit.esMaskedEdit mskDataAP 
         Height          =   285
         Left            =   6660
         TabIndex        =   35
         Top             =   585
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   503
         MouseIcon       =   "frmCadCemiterio.frx":0000
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
      Begin VB.Label Label5 
         Caption         =   "Sepulturas..............:"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   42
         Top             =   990
         Width           =   1500
      End
      Begin VB.Label Label5 
         Caption         =   "Chapa..:"
         Height          =   195
         Index           =   5
         Left            =   5535
         TabIndex        =   40
         Top             =   1575
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Túmulo nº..:"
         Height          =   195
         Index           =   4
         Left            =   3150
         TabIndex        =   38
         Top             =   1575
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Quadra...................:"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   36
         Top             =   1575
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Data.:"
         Height          =   195
         Index           =   2
         Left            =   6120
         TabIndex        =   34
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label5 
         Caption         =   "Horário do sepultamento:"
         Height          =   195
         Index           =   1
         Left            =   3195
         TabIndex        =   32
         Top             =   630
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Velório....................:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   30
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do exumado.:"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   28
         Top             =   270
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "B) Dados do Falecido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   45
      TabIndex        =   18
      Top             =   1710
      Width           =   7755
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome.....................:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   26
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   1665
         TabIndex        =   25
         Top             =   270
         Width           =   5925
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6210
         TabIndex        =   24
         Top             =   570
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Idade.:"
         Height          =   225
         Index           =   4
         Left            =   5625
         TabIndex        =   23
         Top             =   585
         Width           =   510
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1665
         TabIndex        =   22
         Top             =   570
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RG.........................:"
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   21
         Top             =   585
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF..:"
         Height          =   225
         Index           =   2
         Left            =   3330
         TabIndex        =   20
         Top             =   570
         Width           =   570
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3870
         TabIndex        =   19
         Top             =   570
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "A) Dados do Permissionário"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7755
      Begin VB.OptionButton optTipo 
         Caption         =   "Parente em 1º grau"
         Height          =   195
         Index           =   1
         Left            =   5130
         TabIndex        =   17
         Top             =   360
         Width           =   2085
      End
      Begin VB.OptionButton optTipo 
         Caption         =   "Concessionário"
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1665
         MaxLength       =   6
         TabIndex        =   1
         Top             =   300
         Width           =   945
      End
      Begin prjChameleon.chameleonButton cmdCnsImovel 
         Height          =   315
         Left            =   2670
         TabIndex        =   2
         ToolTipText     =   "Consulta Imóvel"
         Top             =   270
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
         MICON           =   "frmCadCemiterio.frx":001C
         PICN            =   "frmCadCemiterio.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3870
         TabIndex        =   15
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF..:"
         Height          =   225
         Index           =   11
         Left            =   3330
         TabIndex        =   14
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RG.........................:"
         Height          =   225
         Index           =   10
         Left            =   165
         TabIndex        =   13
         Top             =   1215
         Width           =   1515
      End
      Begin VB.Label lblCompl 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1680
         TabIndex        =   12
         Top             =   1200
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fone.:"
         Height          =   225
         Index           =   9
         Left            =   5535
         TabIndex        =   11
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label lblCep 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6075
         TabIndex        =   10
         Top             =   1200
         Width           =   1530
      End
      Begin VB.Label lblNumImovel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6090
         TabIndex        =   9
         Top             =   915
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº.....:"
         Height          =   225
         Index           =   1
         Left            =   5535
         TabIndex        =   8
         Top             =   930
         Width           =   405
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido...:"
         Height          =   225
         Index           =   0
         Left            =   165
         TabIndex        =   7
         Top             =   315
         Width           =   1560
      End
      Begin VB.Label lblRua 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1680
         TabIndex        =   6
         Top             =   915
         Width           =   3735
      End
      Begin VB.Label lblProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1680
         TabIndex        =   5
         Top             =   630
         Width           =   5925
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...............:"
         Height          =   225
         Index           =   6
         Left            =   165
         TabIndex        =   4
         Top             =   930
         Width           =   1695
      End
      Begin VB.Label lblRS 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome.....................:"
         Height          =   225
         Left            =   165
         TabIndex        =   3
         Top             =   630
         Width           =   1695
      End
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   7950
      TabIndex        =   58
      ToolTipText     =   "Consulta Imóveis Cadastrados"
      Top             =   1215
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmCadCemiterio.frx":0192
      PICN            =   "frmCadCemiterio.frx":01AE
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
      Left            =   7950
      TabIndex        =   59
      ToolTipText     =   "Gravar o Registro"
      Top             =   135
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   14
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
      MICON           =   "frmCadCemiterio.frx":0308
      PICN            =   "frmCadCemiterio.frx":0324
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7950
      TabIndex        =   60
      ToolTipText     =   "Cancelar Edição"
      Top             =   495
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   14
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
      MICON           =   "frmCadCemiterio.frx":06C9
      PICN            =   "frmCadCemiterio.frx":06E5
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
      Height          =   315
      Left            =   7950
      TabIndex        =   61
      ToolTipText     =   "Novo Registro"
      Top             =   135
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
      MICON           =   "frmCadCemiterio.frx":083F
      PICN            =   "frmCadCemiterio.frx":085B
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
      Left            =   7950
      TabIndex        =   62
      ToolTipText     =   "Editar Registro"
      Top             =   495
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
      MICON           =   "frmCadCemiterio.frx":09B5
      PICN            =   "frmCadCemiterio.frx":09D1
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
      Left            =   7950
      TabIndex        =   63
      ToolTipText     =   "Desativar este imóvel"
      Top             =   855
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
      MICON           =   "frmCadCemiterio.frx":0B2B
      PICN            =   "frmCadCemiterio.frx":0B47
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
      Left            =   7965
      TabIndex        =   64
      ToolTipText     =   "Sair da Tela"
      Top             =   1575
      Width           =   1155
      _ExtentX        =   2037
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCadCemiterio.frx":0BE9
      PICN            =   "frmCadCemiterio.frx":0C05
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
Attribute VB_Name = "frmCadCemiterio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Centraliza Me
End Sub

