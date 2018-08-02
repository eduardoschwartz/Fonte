VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmParamParcela 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parâmetros das Parcelas do Cálculo"
   ClientHeight    =   3360
   ClientLeft      =   2400
   ClientTop       =   3000
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3360
   ScaleWidth      =   8250
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   7200
      TabIndex        =   41
      ToolTipText     =   "Gravar os Dados"
      Top             =   2880
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParamParcela.frx":0000
      PICN            =   "frmParamParcela.frx":001C
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
      Left            =   6120
      TabIndex        =   36
      ToolTipText     =   "Cancelar Edição"
      Top             =   2880
      Width           =   1035
      _ExtentX        =   1826
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParamParcela.frx":03C1
      PICN            =   "frmParamParcela.frx":03DD
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
      Left            =   120
      TabIndex        =   37
      ToolTipText     =   "Novo Registro"
      Top             =   2880
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParamParcela.frx":0537
      PICN            =   "frmParamParcela.frx":0553
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
      Left            =   1170
      TabIndex        =   38
      ToolTipText     =   "Editar Registro"
      Top             =   2880
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParamParcela.frx":06AD
      PICN            =   "frmParamParcela.frx":06C9
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
      Left            =   2220
      TabIndex        =   39
      ToolTipText     =   "Excluir Registro"
      Top             =   2880
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParamParcela.frx":0823
      PICN            =   "frmParamParcela.frx":083F
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
      Left            =   7170
      TabIndex        =   40
      ToolTipText     =   "Sair da Tela"
      Top             =   2880
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmParamParcela.frx":08E1
      PICN            =   "frmParamParcela.frx":08FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmParamParcela.frx":096B
      Left            =   1770
      List            =   "frmParamParcela.frx":096D
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   780
      Width           =   1065
   End
   Begin VB.TextBox txtQtde 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1770
      TabIndex        =   4
      Top             =   1170
      Width           =   1035
   End
   Begin VB.ComboBox cmbUnica 
      Height          =   315
      ItemData        =   "frmParamParcela.frx":096F
      Left            =   1770
      List            =   "frmParamParcela.frx":0979
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   1065
   End
   Begin VB.TextBox txtPerc 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1770
      TabIndex        =   2
      Top             =   1950
      Width           =   1035
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmParamParcela.frx":0987
      Left            =   1740
      List            =   "frmParamParcela.frx":0989
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   5175
   End
   Begin esMaskEdit.esMaskedEdit mskVencUnica 
      Height          =   285
      Left            =   1770
      TabIndex        =   6
      Top             =   2310
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":098B
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   1
      Left            =   4440
      TabIndex        =   12
      Top             =   780
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":09A7
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   1110
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":09C3
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   1440
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":09DF
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   4
      Left            =   4440
      TabIndex        =   15
      Top             =   1770
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":09FB
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   5
      Left            =   4440
      TabIndex        =   16
      Top             =   2100
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0A17
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   6
      Left            =   4440
      TabIndex        =   17
      Top             =   2430
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0A33
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   7
      Left            =   6945
      TabIndex        =   18
      Top             =   780
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0A4F
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   8
      Left            =   6945
      TabIndex        =   19
      Top             =   1110
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0A6B
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   9
      Left            =   6945
      TabIndex        =   20
      Top             =   1425
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0A87
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   10
      Left            =   6945
      TabIndex        =   21
      Top             =   1755
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0AA3
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   11
      Left            =   6945
      TabIndex        =   22
      Top             =   2070
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0ABF
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
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   12
      Left            =   6945
      TabIndex        =   23
      Top             =   2400
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   503
      MouseIcon       =   "frmParamParcela.frx":0ADB
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
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 01:"
      Height          =   195
      Index           =   1
      Left            =   3240
      TabIndex        =   35
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 02:"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   34
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 03:"
      Height          =   195
      Index           =   3
      Left            =   3240
      TabIndex        =   33
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 04:"
      Height          =   195
      Index           =   4
      Left            =   3240
      TabIndex        =   32
      Top             =   1830
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 05:"
      Height          =   195
      Index           =   5
      Left            =   3240
      TabIndex        =   31
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 06:"
      Height          =   195
      Index           =   6
      Left            =   3240
      TabIndex        =   30
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 07:"
      Height          =   195
      Index           =   7
      Left            =   5760
      TabIndex        =   29
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 08:"
      Height          =   195
      Index           =   8
      Left            =   5760
      TabIndex        =   28
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 09:"
      Height          =   195
      Index           =   9
      Left            =   5760
      TabIndex        =   27
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 10:"
      Height          =   195
      Index           =   10
      Left            =   5760
      TabIndex        =   26
      Top             =   1830
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 11:"
      Height          =   195
      Index           =   11
      Left            =   5760
      TabIndex        =   25
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 12:"
      Height          =   195
      Index           =   12
      Left            =   5760
      TabIndex        =   24
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de Cálculo......:"
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde de Parcelas...:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   1230
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tem Parcela Única:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1620
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "% Parcela Única....:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   2010
      Width           =   1455
   End
   Begin VB.Label lblVencUnica 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencto.Única........:"
      Height          =   195
      Left            =   270
      TabIndex        =   7
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Lançamento:"
      Height          =   240
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   1695
   End
End
Attribute VB_Name = "frmParamParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim sRet As String

Private Sub cmbAno_Click()
Limpa
Le
End Sub

Private Sub cmbTipo_Click()
cmbAno.Text = Year(Now)
Le
End Sub

Private Sub cmdAlterar_Click()
    Eventos "INCLUIR"
    Evento = "Alterar"

End Sub

Private Sub cmdCancel_Click()
    Eventos "INICIAR"
    Evento = ""

End Sub

Private Sub cmdExcluir_Click()

If MsgBox("Excluir os parâmetros para o ano de " & cmbAno.Text & " ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
    Sql = "DELETE FROM PARAMPARCELA WHERE ANO=" & cmbAno.Text
    cn.Execute Sql, rdExecDirect
    cmbAno.ListIndex = 0
End If

End Sub

Private Sub cmdGravar_Click()
    
If Val(txtQtde.Text) = 0 Then
      MsgBox "Digite a Qtde de Parcelas.", vbExclamation, "Atenção"
      txtQtde.SetFocus
      Exit Sub
End If
    
If Val(txtPerc.Text) = 0 Then txtPerc.Text = 0

If Not IsDate(mskVencUnica.Text) Then
     MsgBox "Digite a Data de Vencimento da Parcela única", vbExclamation, "Atenção"
     mskVencUnica.SetFocus
     Exit Sub
End If

If Right$(mskVencUnica.Text, 4) <> cmbAno.Text Then
     MsgBox "Ano de vencimento da parcela única tem que ser igual ao ano de cálculo.", vbExclamation, "Atenção"
     mskVencUnica.SetFocus
     Exit Sub
End If

For x = 1 To Val(txtQtde.Text)
      If Not IsDate(mskVenc(x).Text) Then
           MsgBox "Digite a Data de Vencimento para a Parcela nº " & x, vbExclamation, "Atenção"
           mskVenc(x).SetFocus
           Exit Sub
      End If
      If Right$(mskVenc(x).Text, 4) <> cmbAno.Text And Right$(mskVenc(x).Text, 4) <> Val(cmbAno.Text) + 1 Then
           MsgBox "O Vencimento da parcela nº " & x & " deve ser do mesmo ano.", vbExclamation, "Atenção"
           mskVenc(x).SetFocus
           Exit Sub
      End If
Next
    
Grava
    
Eventos "INICIAR"
Evento = ""

End Sub

Private Sub Grava()

Dim qd As New rdoQuery

On Error Resume Next
RdoAux.Close
On Error GoTo 0
Set qd.ActiveConnection = cn

qd.Sql = "{ Call spGRAVAPARAMPARCELA(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
qd(0) = cmbTipo.ItemData(cmbTipo.ListIndex)
qd(1) = Val(cmbAno.Text)
qd(2) = Val(txtQtde.Text)
qd(3) = IIf(cmbUnica.ListIndex = 0, "S", "N")
qd(4) = Virg2Ponto(txtPerc.Text)
qd(5) = Format(mskVencUnica.Text, "mm/dd/yyyy")
If IsDate(mskVenc(1).Text) Then
     qd(6) = Format(mskVenc(1).Text, "mm/dd/yyyy")
Else
     qd(6) = Null
End If
If IsDate(mskVenc(2).Text) And mskVenc(2).Enabled = True Then
     qd(7) = Format(mskVenc(2).Text, "mm/dd/yyyy")
Else
     qd(7) = Null
End If
If IsDate(mskVenc(3).Text) And mskVenc(3).Enabled = True Then
     qd(8) = Format(mskVenc(3).Text, "mm/dd/yyyy")
Else
     qd(8) = Null
End If
If IsDate(mskVenc(4).Text) And mskVenc(4).Enabled = True Then
     qd(9) = Format(mskVenc(4).Text, "mm/dd/yyyy")
Else
     qd(9) = Null
End If
If IsDate(mskVenc(5).Text) And mskVenc(5).Enabled = True Then
     qd(10) = Format(mskVenc(5).Text, "mm/dd/yyyy")
Else
     qd(10) = Null
End If
If IsDate(mskVenc(6).Text) And mskVenc(6).Enabled = True Then
     qd(11) = Format(mskVenc(6).Text, "mm/dd/yyyy")
Else
     qd(11) = Null
End If
If IsDate(mskVenc(7).Text) And mskVenc(7).Enabled = True Then
     qd(12) = Format(mskVenc(7).Text, "mm/dd/yyyy")
Else
     qd(12) = Null
End If
If IsDate(mskVenc(8).Text) And mskVenc(8).Enabled = True Then
     qd(13) = Format(mskVenc(8).Text, "mm/dd/yyyy")
Else
     qd(13) = Null
End If
If IsDate(mskVenc(9).Text) And mskVenc(9).Enabled = True Then
     qd(14) = Format(mskVenc(9).Text, "mm/dd/yyyy")
Else
     qd(14) = Null
End If
If IsDate(mskVenc(10).Text) And mskVenc(10).Enabled = True Then
     qd(15) = Format(mskVenc(10).Text, "mm/dd/yyyy")
Else
     qd(15) = Null
End If
If IsDate(mskVenc(11).Text) And mskVenc(11).Enabled = True Then
     qd(16) = Format(mskVenc(11).Text, "mm/dd/yyyy")
Else
     qd(16) = Null
End If
If IsDate(mskVenc(12).Text) And mskVenc(12).Enabled = True Then
     qd(17) = Format(mskVenc(12).Text, "mm/dd/yyyy")
Else
     qd(17) = Null
End If

Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
RdoAux.Close

End Sub

Private Sub cmdNovo_Click()
   Eventos "INCLUIR"
   Evento = "Novo"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Le()

Sql = "SELECT ANO,QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,VENC01,VENC02,VENC03,VENC04,VENC05,"
Sql = Sql & "VENC06,VENC07,VENC08,VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA WHERE CODTIPO="
Sql = Sql & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND ANO=" & cmbAno.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
       If .RowCount > 0 Then
            txtQtde.Text = !qtdeparcela
            If !PARCELAUNICA = "S" Then
                 cmbUnica.ListIndex = 0
            Else
                cmbUnica.ListIndex = 1
            End If
            txtPerc.Text = Format(!DESCONTOUNICA, "#0.00")
            If Not IsNull(!VENCUNICA) Then mskVencUnica.Text = Format(!VENCUNICA, "dd/mm/yyyy")
            If Not IsNull(!VENC01) Then mskVenc(1).Text = Format(!VENC01, "dd/mm/yyyy")
            If Not IsNull(!VENC02) Then mskVenc(2).Text = Format(!VENC02, "dd/mm/yyyy")
            If Not IsNull(!VENC03) Then mskVenc(3).Text = Format(!VENC03, "dd/mm/yyyy")
            If Not IsNull(!VENC04) Then mskVenc(4).Text = Format(!VENC04, "dd/mm/yyyy")
            If Not IsNull(!VENC05) Then mskVenc(5).Text = Format(!VENC05, "dd/mm/yyyy")
            If Not IsNull(!VENC06) Then mskVenc(6).Text = Format(!VENC06, "dd/mm/yyyy")
            If Not IsNull(!VENC07) Then mskVenc(7).Text = Format(!VENC07, "dd/mm/yyyy")
            If Not IsNull(!VENC08) Then mskVenc(8).Text = Format(!VENC08, "dd/mm/yyyy")
            If Not IsNull(!VENC09) Then mskVenc(9).Text = Format(!VENC09, "dd/mm/yyyy")
            If Not IsNull(!VENC10) Then mskVenc(10).Text = Format(!VENC10, "dd/mm/yyyy")
            If Not IsNull(!VENC11) Then mskVenc(11).Text = Format(!VENC11, "dd/mm/yyyy")
            If Not IsNull(!VENC12) Then mskVenc(12).Text = Format(!VENC12, "dd/mm/yyyy")
       Else
            Limpa
       End If
End With

End Sub

Private Sub Limpa()
Dim x As Integer

txtQtde.Text = ""
txtPerc.Text = ""
cmbUnica.ListIndex = 0
LimpaMascara mskVencUnica
For x = 1 To 12
      LimpaMascara mskVenc(x)
Next

End Sub

Private Sub Form_Load()
Dim x As Integer

For x = 2004 To Year(Now) + 1
    cmbAno.AddItem CStr(x)
Next

Centraliza Me
sRet = RetEventUserForm(Me.Name)
Eventos "INICIAR"

Sql = "SELECT CODTIPO,DESCTIPO FROM TIPOCARNE"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !DESCTIPO
        cmbTipo.ItemData(cmbTipo.NewIndex) = !CodTipo
       .MoveNext
    Loop
End With
cmbTipo.ListIndex = 0
cmbAno.Text = Year(Now)
cmbUnica.ListIndex = 0
End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

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
   For Each Ct In frmParamParcela
       If TypeOf Ct Is TextBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   cmbUnica.Enabled = False
   cmbUnica.BackColor = Kde
   cmbAno.Enabled = True
   cmbAno.BackColor = Branco
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmParamParcela
       If TypeOf Ct Is TextBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = Branco
         Ct.Enabled = True
       End If
   Next
   cmbUnica.Enabled = True
   cmbUnica.BackColor = Branco
   cmbAno.Enabled = False
   cmbAno.BackColor = Kde
End If

FormHagana

End Sub

Private Sub mskVenc_GotFocus(Index As Integer)
mskVenc(Index).SelStart = 0
mskVenc(Index).SelLength = 10
End Sub

Private Sub mskVenc_KeyPress(Index As Integer, KeyAscii As Integer)

If Index > 1 And KeyAscii <> vbKeyTab Then
     If mskVenc(Index - 1).ClipText = "" Then
          KeyAscii = 0
          MsgBox "Digite o vencimento anterior.", vbExclamation, "Atenção"
          mskVenc(Index - 1).SetFocus
          LimpaMascara mskVenc(Index)
     End If
End If

End Sub

Private Sub mskVenc_LostFocus(Index As Integer)

If mskVenc(Index).Enabled = False Then Exit Sub
If mskVenc(Index).ClipText <> "" Then
     If Not IsDate(mskVenc(Index).Text) Then
          MsgBox "Data inválida.", vbExclamation, "Atenção"
          mskVenc(Index).SetFocus
     Else
          If Index > 1 Then
               If Not IsDate(mskVenc(Index - 1).Text) Then
                  mskVenc(Index - 1).SetFocus
                  Exit Sub
               End If
               If (CDate(mskVenc(Index).Text) < CDate(mskVenc(Index - 1).Text)) Then
                    MsgBox "A data do vencimento " & Index & " tem que ser maior que a do vencimento anterior", vbExclamation, "Atenção"
                    mskVenc(Index).SetFocus
                    Exit Sub
               End If
               For x = 1 To Index - 1
                    If Month(CDate(mskVenc(x).Text)) = Month(CDate(mskVenc(Index).Text)) And Year(CDate(mskVenc(x).Text)) = Year(CDate(mskVenc(Index).Text)) Then
                        MsgBox "Já foi cadastrado uma parcela para este mês de vencimento.", vbExclamation, "Atenção"
                        Exit Sub
                    End If
               Next
          End If
     End If
End If

End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
Tweak txtPerc, KeyAscii, DecimalPositive

End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)
Tweak txtQtde, KeyAscii, IntegerPositive
End Sub

Private Sub txtQtde_LostFocus()

Dim x As Integer

If Val(txtQtde.Text) > 12 Then
     MsgBox "Máximo 12 parcelas.", vbExclamation, "atenção"
     txtQtde.SetFocus
     Exit Sub
End If

For x = 1 To 12
      lblVenc(x).Enabled = True
      mskVenc(x).Enabled = True
      mskVenc(x).BackColor = Branco
Next

For x = Val(txtQtde.Text) + 1 To 12
      lblVenc(x).Enabled = False
      mskVenc(x).Enabled = False
      mskVenc(x).BackColor = Kde
Next

End Sub
