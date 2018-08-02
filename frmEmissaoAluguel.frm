VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmEmissaoAluguel 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão dos Boletos de Cobrança de Aluguel"
   ClientHeight    =   4845
   ClientLeft      =   4395
   ClientTop       =   3495
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   7515
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Index           =   1
      Left            =   6375
      TabIndex        =   17
      Top             =   810
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":0000
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
      Left            =   6375
      TabIndex        =   18
      Top             =   1140
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":001C
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
      Left            =   6375
      TabIndex        =   19
      Top             =   1470
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":0038
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
      Left            =   6375
      TabIndex        =   20
      Top             =   1785
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":0054
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
      Left            =   6375
      TabIndex        =   21
      Top             =   2115
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":0070
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
      Left            =   6375
      TabIndex        =   22
      Top             =   2445
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":008C
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
      Left            =   6375
      TabIndex        =   23
      Top             =   2775
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":00A8
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
      Left            =   6375
      TabIndex        =   24
      Top             =   3105
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":00C4
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
      Left            =   6375
      TabIndex        =   25
      Top             =   3420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":00E0
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
      Left            =   6375
      TabIndex        =   26
      Top             =   3750
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":00FC
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
      Left            =   6375
      TabIndex        =   27
      Top             =   4080
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":0118
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
      Left            =   6375
      TabIndex        =   28
      Top             =   4410
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   503
      MouseIcon       =   "frmEmissaoAluguel.frx":0134
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
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   3540
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   90
      Width           =   3810
   End
   Begin VB.ComboBox cmbAno 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEmissaoAluguel.frx":0150
      Left            =   750
      List            =   "frmEmissaoAluguel.frx":0152
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   90
      Width           =   915
   End
   Begin prjChameleon.chameleonButton cmdHelp 
      Height          =   330
      Left            =   2715
      TabIndex        =   1
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   4425
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmEmissaoAluguel.frx":0154
      PICN            =   "frmEmissaoAluguel.frx":0170
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
      Height          =   330
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   4425
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
      MICON           =   "frmEmissaoAluguel.frx":02CA
      PICN            =   "frmEmissaoAluguel.frx":02E6
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
      Height          =   330
      Left            =   1575
      TabIndex        =   3
      ToolTipText     =   "Emissão dos boletos de aluguel"
      Top             =   4425
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   582
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
      MICON           =   "frmEmissaoAluguel.frx":0354
      PICN            =   "frmEmissaoAluguel.frx":0370
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   2805
      Left            =   30
      TabIndex        =   0
      Top             =   525
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4948
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código       |<Nome                                                                        "
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1260
      Left            =   0
      TabIndex        =   12
      Top             =   4920
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
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 07:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   5190
      TabIndex        =   40
      Top             =   2820
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 08:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   5190
      TabIndex        =   39
      Top             =   3150
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 09:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   5190
      TabIndex        =   38
      Top             =   3480
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 10:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   10
      Left            =   5190
      TabIndex        =   37
      Top             =   3810
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 11:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   11
      Left            =   5190
      TabIndex        =   36
      Top             =   4140
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 12:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   12
      Left            =   5190
      TabIndex        =   35
      Top             =   4470
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 01:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   5190
      TabIndex        =   34
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 02:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   5190
      TabIndex        =   33
      Top             =   1170
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 03:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   5190
      TabIndex        =   32
      Top             =   1500
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 04:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   5190
      TabIndex        =   31
      Top             =   1830
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 05:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   5190
      TabIndex        =   30
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento 06:"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   5190
      TabIndex        =   29
      Top             =   2490
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Aluguel..:"
      Height          =   195
      Index           =   1
      Left            =   2205
      TabIndex        =   16
      Top             =   135
      Width           =   1320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   195
      Left            =   240
      TabIndex        =   15
      Top             =   150
      Width           =   570
   End
   Begin VB.Label lblValorParc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   3705
      TabIndex        =   11
      Top             =   4020
      Width           =   990
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Parcela.:"
      Height          =   225
      Left            =   2625
      TabIndex        =   10
      Top             =   4020
      Width           =   1035
   End
   Begin VB.Label lblValorTotal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1185
      TabIndex        =   9
      Top             =   4020
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total....:"
      Height          =   225
      Left            =   90
      TabIndex        =   8
      Top             =   4020
      Width           =   1065
   End
   Begin VB.Label lblNumParc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1185
      TabIndex        =   7
      Top             =   3735
      Width           =   330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Parcelas..:"
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   3735
      Width           =   1080
   End
   Begin VB.Label lblDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1185
      TabIndex        =   5
      Top             =   3450
      Width           =   3765
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição.....:"
      Height          =   225
      Left            =   90
      TabIndex        =   4
      Top             =   3450
      Width           =   1080
   End
End
Attribute VB_Name = "frmEmissaoAluguel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim nMesesFaltando As Integer, sVencto1 As String
Dim xImovel As clsImovel
Dim sNumInsc As String
Dim sEndImovel As String
Dim nNumImovel As Integer
Dim sComplImovel As String
Dim sBairroImovel As String
Dim sEndEntrega As String
Dim nNumEntrega As Integer
Dim sBairroEntrega As String
Dim sComplEntrega As String
Dim sCepEntrega As String
Dim sCidadeEntrega As String
Dim sUFEntrega As String
Dim sNomeResp As String

Private Sub cmbAno_Click()
grdMain_Click
End Sub

Private Sub cmbTipo_Click()
Limpa
If cmbTipo.ListIndex = -1 Then Exit Sub
grdMain.Rows = 1
Sql = "SELECT CODREDUZIDO,NOME FROM MANUTENCAOALUGUEL WHERE CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddItem Format(!CODREDUZIDO, "000000") & Chr(9) & !Nome
       .MoveNext
    Loop
End With

If grdMain.Rows > 1 Then
    grdMain_Click
End If

End Sub

Private Sub cmdBaixa_Click()
Dim bAchou As Boolean, nCodReduz As Long
If grdMain.Rows = 1 Then
    MsgBox "Não existem registros.", vbCritical, "ERRO"
    Exit Sub
End If

For x = 1 To mskVenc.Count
    If mskVenc(x).ClipText <> "" Then
        If Not IsDate(mskVenc(x).Text) Then
            MsgBox "Data inválida: " & mskVenc(x).Text
            Exit Sub
        End If
    Else
        Exit For
    End If
Next
nCodReduz = Val(grdMain.TextMatrix(grdMain.Row, 0))
Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND "
Sql = Sql & "CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND ANOEXERCICIO=" & Val(cmbAno.Text) & " AND NUMPARCELA<>12"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If MsgBox("Ja foram emitidos boletos de aluguel deste inquilino para o ano informado." & vbCrLf & "A nova geração irá cancelar os boletos antigos e gerar novos lancamentos." & vbCrLf & "Deseja continuar ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            bAchou = False
            Do Until .EOF
                If !statuslanc = 2 Then
                    bAchou = True
                    Exit Do
                End If
               .MoveNext
            Loop
            If bAchou Then
                MsgBox "Alguns dos boletos já foram pagos." & vbCrLf & "Cancele os pagamentos para poder emitir novos boletos.", vbExclamation, "Atenção"
            Else
               
                CarregaEnd nCodReduz
                'GravaCarneTmp
                EmiteBoleto nCodReduz
            End If
        End If
    Else
        CarregaEnd nCodReduz
'        GravaCarneTmp
        EmiteBoleto nCodReduz
    End If
   .Close
End With

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer

For x = 2004 To Year(Now) + 1
    cmbAno.AddItem CStr(x)
Next

Set xImovel = New clsImovel
Centraliza Me
Sql = "SELECT TIPOALUGUEL.CODLANCAMENTO,LANCAMENTO.DESCFULL FROM TIPOALUGUEL INNER JOIN "
Sql = Sql & "LANCAMENTO ON TIPOALUGUEL.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !DESCFULL
        cmbTipo.ItemData(cmbTipo.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
cmbAno.Text = Year(Now)
End Sub

Private Sub grdMain_Click()
Limpa
With grdMain
    If .Rows = 1 Then Exit Sub
    Sql = "SELECT * FROM MANUTENCAOALUGUEL WHERE CODREDUZIDO=" & Val(.TextMatrix(.Row, 0)) & " AND "
    Sql = Sql & "CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        lblValorTotal.Caption = FormatNumber(!ValorTotal, 2)
        sVencto1 = Format(!DataVencto, "dd/mm/yyyy")
        lblDesc.Caption = SubNull(!descricao)
        CarregaDados
       .Close
    End With
End With
End Sub

Private Sub Limpa()

lblDesc.Caption = ""
lblValorParc.Caption = ""
lblValorTotal.Caption = ""
lblNumParc.Caption = ""
For x = 1 To 12
    LimpaMascara mskVenc(x)
Next

End Sub

Private Sub CarregaDados()
Dim nMesAtual As Integer, dData1 As Date, x As Integer
Dim nDia As Integer, nMes As Integer, nAno As Integer, sData As String
If cmbAno.ListIndex = -1 Then cmbAno.Text = Year(Now)
dData1 = CDate(sVencto1)
If Year(dData1) < Val(cmbAno.Text) Then 'antigo
   If Year(Now) = Year(cmbAno.Text) Then
      nMesAtual = Month(Now)
   Else
      nMesAtual = 2
   End If
Else 'mesmo ano
   nMesAtual = Month(dData1)
End If
nMesesFaltando = 14 - nMesAtual

nDia = Day(dData1)
nMes = nMesAtual
nAno = Val(cmbAno.Text)

If Month(dData1) > 1 Then
    nMesesFaltando = nMesesFaltando - 1
End If

If nMesesFaltando = 13 Then nMesesFaltando = 12
For x = 1 To nMesesFaltando
    If nDia > 26 Then
        Select Case nMes
            Case 2
                nDia = 28
            Case 4, 6, 9, 11
                nDia = 30
            Case 1, 3, 5, 7, 8, 10, 12
                nDia = 31
        End Select
    End If
    sData = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
    mskVenc(x).Text = sData
    nMes = nMes + 1
    If nMes = 13 Then
        nMes = 1
        nAno = nAno + 1
    End If
Next

If Val(cmbAno) = 2014 Then
    mskVenc(1).Text = "31/01/2014"
    mskVenc(2).Text = "28/02/2014"
    mskVenc(3).Text = "31/03/2014"
    mskVenc(4).Text = "30/04/2014"
    mskVenc(5).Text = "31/05/2014"
    mskVenc(6).Text = "30/06/2014"
    mskVenc(7).Text = "31/07/2014"
    mskVenc(8).Text = "31/08/2014"
    mskVenc(9).Text = "30/09/2014"
    mskVenc(10).Text = "31/10/2014"
    mskVenc(11).Text = "30/11/2014"
    mskVenc(12).Text = "31/12/2014"
ElseIf Val(cmbAno) = 2016 Then
    mskVenc(1).Text = "29/01/2016"
    mskVenc(2).Text = "29/02/2016"
    mskVenc(3).Text = "31/03/2016"
    mskVenc(4).Text = "30/04/2016"
    mskVenc(5).Text = "31/05/2016"
    mskVenc(6).Text = "30/06/2016"
    mskVenc(7).Text = "31/07/2016"
    mskVenc(8).Text = "31/08/2016"
    mskVenc(9).Text = "30/09/2016"
    mskVenc(10).Text = "31/10/2016"
    mskVenc(11).Text = "30/11/2016"
    mskVenc(12).Text = "30/12/2016"
End If

lblNumParc.Caption = nMesesFaltando
CalculaParcela
End Sub

Private Sub CarregaDadosOld()
Dim nMesAtual As Integer, dData1 As Date, x As Integer
Dim nDia As Integer, nMes As Integer, nAno As Integer, sData As String
If cmbAno.ListIndex = -1 Then cmbAno.Text = Year(Now)
dData1 = CDate(sVencto1)
If Year(dData1) < Val(cmbAno.Text) Then 'antigo
   If Year(Now) = Year(cmbAno.Text) Then
      nMesAtual = Month(Now)
   Else
      nMesAtual = 1
   End If
Else 'mesmo ano
   nMesAtual = Month(dData1)
End If
nMesesFaltando = 13 - nMesAtual

nDia = Day(dData1)
nMes = nMesAtual
nAno = Val(cmbAno.Text)

For x = 1 To nMesesFaltando
    
    sData = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
    mskVenc(x).Text = sData
    nMes = nMes + 1
Next

lblNumParc.Caption = nMesesFaltando
CalculaParcela
End Sub

Private Sub mskVenc_GotFocus(Index As Integer)

On Error Resume Next
mskVenc(Index).SelStart = 0
mskVenc(Index).SelLength = 10
mskVenc(Index).SetFocus
End Sub

Private Sub mskVenc_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
If mskVenc(Index).Locked = True Then Exit Sub
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
On Error Resume Next
If mskVenc(Index).ClipText <> "" Then
     If Not IsDate(mskVenc(Index).Text) Then
          MsgBox "Data inválida.", vbExclamation, "Atenção"
          mskVenc(Index).SetFocus
          Exit Sub
     Else
          If Mid(mskVenc(Index).Text, 4, 2) > 12 Then
            MsgBox "Data inválida.", vbExclamation, "Atenção"
            mskVenc(Index).SetFocus
            Exit Sub
          End If
          
          If Index > 1 Then
               If Not IsDate(mskVenc(Index - 1).Text) Then
                  mskVenc(Index - 1).SetFocus
                  Exit Sub
               End If
               If (CDate(mskVenc(Index).Text) < CDate(mskVenc(Index - 1).Text)) Then
                    MsgBox "A data do vencimento " & Index & " tem que ser maior que a do vencimento anterior", vbExclamation, "Atenção"
'                    mskVenc(Index).SetFocus
                    Exit Sub
               End If
          End If
     End If

    df = ValidaFeriado(CDate(mskVenc(Index).Text))
    If df = 1 Then
        If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVenc(Index).Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            mskVenc(Index).Text = Format(RetornaDiaUtil(CDate(mskVenc(Index).Text)), "dd/mm/yyyy")
        Else
            Exit Sub
        End If
    ElseIf df = 2 Then
        If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVenc(Index).Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            mskVenc(Index).Text = Format(RetornaDiaUtil(CDate(mskVenc(Index).Text)), "dd/mm/yyyy")
        Else
            Exit Sub
        End If
    ElseIf df = 3 Then
        Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
        Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
        Sql = Sql & " Where DIA = " & Day(CDate(mskVenc(Index).Text))
        Sql = Sql & " AND MES=" & Month(CDate(mskVenc(Index).Text)) & " AND ANO=" & Year(CDate(mskVenc(Index).Text))
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskVenc(Index).Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                    mskVenc(Index).Text = RetornaDiaUtil(CDate(mskVenc(Index).Text))
                Else
                    Exit Sub
                End If
              .Close
            End If
        End With
    End If
Else
    If Index < 12 Then
        If mskVenc(Index + 1).ClipText <> "" Then
           MsgBox "Digite uma data válida.", vbExclamation, "atenção"
           mskVenc(Index).SetFocus
        End If
    End If
End If

For x = 1 To 12
    If mskVenc(x).ClipText = "" Then
        Exit For
    End If
Next

lblNumParc.Caption = x - 1
CalculaParcela

End Sub

Private Sub CalculaParcela()
Dim nValorTotal As Double, nValorParc As Double, nNumParc As Integer

nNumParc = Val(lblNumParc.Caption)
nValorTotal = CDbl(lblValorTotal.Caption)
nValorParc = nValorTotal / Val(lblNumParc.Caption)
lblValorParc.Caption = FormatNumber(nValorParc, 2)

End Sub

Private Sub GravaCarneTmp()
On Error GoTo Erro

Dim x As Integer
Dim RdoAux2 As rdoResultset, qd As New rdoQuery
Dim nCodReduz As Long
Dim nCodTributo As Integer
Dim sTipoImposto As String
Dim sDescImposto As String
Dim nAno As Integer
Dim sNumProc As String
Dim dDataProc As Date
Dim dDataVencto As Date
Dim nNumDoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim sValorParc As String
Dim sVencimento As String
Dim nCodLanc As Integer
Dim nSeq As Integer
Dim nComplemento As Integer
Dim nValorTotal As Double
Dim nValorParc As Double
Dim nValorParcUnica As Double
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
Dim sFullTrib As String

If MsgBox("Confirma criação da Guia ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nValorParc = 0
nValorParcUnica = 0
'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
   nLastCod = 0
Else
   nLastCod = RdoAux!maximo
End If
RdoAux.Close

nCodReduz = Val(grdMain.TextMatrix(grdMain.Row, 0))
If nCodReduz < 100000 Then
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
Else
    Sql = "SELECT * FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
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
            sNomeResp = !nomecidadao
        End If
    End With
End If


'CARREGA GRID TEMPORARIO
grdTemp.Rows = 1
nCodLanc = cmbTipo.ItemData(cmbTipo.ListIndex)

Select Case nCodLanc
    Case 33
        nCodTributo = 110
    Case 23
        nCodTributo = 111
    Case 34
        nCodTributo = 92
    Case 35
        nCodTributo = 101
    Case 60
        nCodTributo = 534
    Case 66
        nCodTributo = 528
    Case 77
        nCodTributo = 664
End Select
sFullTrib = cmbTipo.Text

ReDim aTributos(0)
ReDim aTributosU(0)

nValorTxExpParc = 0
nValorTxExpUnica = 0

'CALCULA O VALOR PARCELADO

nValorTotal = CDbl(lblValorParc.Caption)
nValorParc = FormatNumber(nValorTotal, 2)

'MONTA TRIBUTOS
sDadosLanc = cmbTipo.Text

For nNumParc = 1 To Val(lblNumParc.Caption)
    sVencimento = mskVenc(nNumParc).Text
    nAno = Year(CDate(sVencimento))
    'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
    Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
    Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nNumParc
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux2!SEQMAXIMA) Then
               nSeq = 0
            Else
               nSeq = RdoAux2!SEQMAXIMA + 1
            End If
        Else
            nSeq = 0
        End If
       .Close
    End With
    nLastCod = nLastCod + 1
    grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & nNumParc & Chr(9) & 0 & Chr(9) & sVencimento & Chr(9) & FormatNumber(IIf(nNumParc = 0, nValorTotal, nValorParc), 2) & Chr(9) & nLastCod

proximo:
Next
'Exit Sub
'DADOS CABEÇALHO
sNumProc = Format(nCodReduz, "000000") & "/" & CStr(Year(Now))
dDataProc = Format(Now, "dd/mm/yyyy")
sDescImposto = sDadosLanc
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
          Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
 ''         Sql = Sql & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
          Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
          Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & ","
          Sql = Sql & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
          Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
          Sql = Sql & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
          cn.Execute Sql, rdExecDirect
          sFullTrib = ""
         'GRAVA DEBITOTRIBUTO
          Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
          Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
          Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nCodTributo & "," & Virg2Ponto(CStr(nValorParc)) & ")"
          cn.Execute Sql, rdExecDirect
         
'          Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
'          Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & ","
'          Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 3 & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ")"
'          cn.Execute Sql, rdExecDirect'

         'GRAVA NUMDOCUMENTO
          Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
          Sql = Sql & .TextMatrix(x, 8) & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ")"
          cn.Execute Sql, rdExecDirect
         'GRAVA PARCELADOCUMENTO
          Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
          Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & .TextMatrix(x, 0) & ","
          Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ","
          Sql = Sql & .TextMatrix(x, 5) & "," & .TextMatrix(x, 8) & ")"
          cn.Execute Sql, rdExecDirect
    Next
End With

'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn

'GRAVA TEMPORARIO
With grdTemp
    For x = 1 To .Rows - 1
        nAno = .TextMatrix(x, 1)
        nCodLanc = .TextMatrix(x, 2)
        nSeq = .TextMatrix(x, 3)
        nNumParc = .TextMatrix(x, 4)
        nComplemento = .TextMatrix(x, 5)
        dDataVencto = CDate(.TextMatrix(x, 6))
        sValorParc = CDbl(.TextMatrix(x, 7)) + CDbl(nValorTxExpParc)
        nNumDoc = .TextMatrix(x, 8)
        'NumBarra2 = Gera2of5Cod(sValorParc, dDataVencto, nNumDoc & RetornaDVNumDoc(nNumDoc), nNumParc, nCodLanc, nSeq, nComplemento)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        sValorParc = sValorParc - CDbl(nValorTxExpParc)

        Sql = "INSERT CARNETMP(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,ENDIMOVEL,NUMIMOVEL,COMPLIMOVEL,"
        Sql = Sql & "BAIRROIMOVEL,ENDENTREGA,NUMENTREGA,COMPLENTREGA,BAIRROENTREGA,CEPENTREGA,CIDADEENTREGA,UFENTREGA,"
        Sql = Sql & "DESCIMPOSTO,EXERCICIO,NUMPROCESSO,DATAPROCESSO,NUMDOCUMENTO,DV,QUADRA,LOTE,DATAVENCTO,NUMPARCELA,"
        Sql = Sql & "NUMTOTPARCELA,VALORPARCELA,STRBARRA1,STRBARRA2,NUMBARRA1,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
        Sql = Sql & "DADOSLANCAMENTO,TAXAEXP,SAIR) VALUES('" & NomeDoUsuario & "'," & x & ",'" & sNumInsc & "','" & nCodReduz & "','"
        Sql = Sql & sTipoImposto & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left$(sEndImovel, 40) & "'," & nNumImovel & ",'" & Left$(sComplImovel, 30) & "','"
        Sql = Sql & Left$(sBairroImovel, 25) & "','" & Left$(sEndEntrega, 40) & "'," & nNumEntrega & ",'" & Left$(sComplEntrega, 30) & "','" & Left$(sBairroEntrega, 25) & "','"
        Sql = Sql & sCepEntrega & "','" & sCidadeEntrega & "','" & sUFEntrega & "','" & Left$(sDescImposto, 30) & "'," & nAno & ",'" & Left$(sNumProc, 25) & "','"
        Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "'," & nNumDoc & "," & RetornaDVNumDoc(nNumDoc) & ",'" & sQuadra & "','"
        Sql = Sql & sLote & "','" & Format(dDataVencto, "mm/dd/yyyy") & "'," & IIf(nNumParc = 0, 1, nNumParc) & "," & Val(grdTemp.Rows - 1) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(sValorParc)) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
        Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & sDadosLanc & "'," & IIf(nNumParc = 0, Virg2Ponto(CStr(nValorTxExpUnica)), Virg2Ponto(CStr(nValorTxExpParc))) & "," & "0" & ")"
        cn.Execute Sql, rdExecDirect

    Next
End With

frmReport.ShowReport "Carne", frmMdi.hwnd, Me.hwnd

'DELETA TEMPORARIO
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

bGerado = True
Exit Sub

Erro:
For x = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(x).Description
Next
Resume Next
End Sub

Private Sub CarregaEnd(nCodReduz As Long)
Dim nCodImovel As Long

nCodImovel = Val(nCodReduz)

Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & nCodImovel
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        CarregaImovel nCodImovel
    Else
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodImovel
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               sNumInsc = SubNull(!INSCESTADUAL)
               sNomeResp = !RazaoSocial
               sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
               nNumImovel = Val(SubNull(!Numero))
               sComplImovel = SubNull(!Complemento)
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
               Sql = "SELECT NOMELOGRADOURO,NUMIMOVEL,COMPLEMENTO,UF,CIDADE.DESCCIDADE AS DESCCIDADE1,"
               Sql = Sql & "BAIRRO.DESCBAIRRO AS DESCBAIRRO1,CEP,MOBILIARIOENDENTREGA.DESCBAIRRO,"
               Sql = Sql & "MOBILIARIOENDENTREGA.DESCCIDADE FROM CIDADE INNER JOIN BAIRRO ON "
               Sql = Sql & "CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE RIGHT OUTER Join "
               Sql = Sql & "MOBILIARIOENDENTREGA ON BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND "
               Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & nCodImovel
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                    If .RowCount > 0 Then
                        sEndEntrega = SubNull(!NomeLogradouro)
                        nNumEntrega = SubNull(!NUMIMOVEL)
                        sComplEntrega = SubNull(!Complemento)
                        sBairroEntrega = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        sCidadeEntrega = IIf(IsNull(!descCidade), SubNull(!DESCCIDADE1), SubNull(!descCidade))
                        sCepEntrega = SubNull(!Cep)
                        sUFEntrega = SubNull(!UF)
                    Else
                        sEndEntrega = sEndImovel
                        nNumEntrega = nNumImovel
                        sComplEntrega = sComplImovel
                        sBairroEntrega = sBairroImovel
                        sCidadeEntrega = "JABOTICABAL"
                        sCepEntrega = "14870-000"
                        sUFEntrega = "SP"
                    End If
                   .Close
               End With
            Else
               Sql = "SELECT CIDADAO.CODCIDADAO,CIDADAO.NOMECIDADAO,CIDADAO.CPF, CIDADAO.CNPJ, CIDADAO.CODLOGRADOURO,vwLOGRADOURO.ABREVTIPOLOG,"
               Sql = Sql & "vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,CIDADAO.NUMIMOVEL, CIDADAO.COMPLEMENTO,CIDADAO.CODBAIRRO, BAIRRO.DESCBAIRRO,"
               Sql = Sql & "CIDADAO.CODCIDADE, CIDADE.DESCCIDADE,CIDADAO.SIGLAUF, UF.DESCUF, CIDADAO.CEP,CIDADAO.NOMELOGRADOURO AS RUA2 "
               Sql = Sql & "FROM vwLOGRADOURO RIGHT OUTER JOIN CIDADAO ON vwLOGRADOURO.CODLOGRADOURO = CIDADAO.CODLOGRADOURO "
               Sql = Sql & "LEFT OUTER JOIN CIDADE INNER JOIN BAIRRO ON CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE INNER JOIN "
               Sql = Sql & "UF ON CIDADE.SIGLAUF = UF.SIGLAUF ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO WHERE CODCIDADAO=" & nCodImovel
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                   If .RowCount > 0 Then
                       sNomeResp = !nomecidadao
                       If !CodLogradouro > 0 Then
                          sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                       Else
                          sEndImovel = SubNull(!RUA2)
                       End If
                       nNumImovel = Val(SubNull(!NUMIMOVEL))
                       sComplImovel = SubNull(!Complemento)
                       sBairroImovel = SubNull(!DescBairro)
                       nNumImovel = Val(SubNull(!NUMIMOVEL))
                       sComplImovel = SubNull(!Complemento)
                       sBairroImovel = SubNull(!DescBairro)

                       sEndEntrega = sEndImovel
                       nNumEntrega = nNumImovel
                       sComplEntrega = sComplImovel
                       sBairroEntrega = sBairroImovel
                       sCidadeEntrega = SubNull(!descCidade)
                       sCepEntrega = ""
                       sUFEntrega = SubNull(!SiglaUF)
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
          sNumInsc = .Inscricao
          sNomeResp = .NomePropPrincipal
          sEndImovel = Trim$(.AbrevTipoLog) & " " & Trim$(.AbrevTitLog) & " " & .NomeLogradouro
          nNumImovel = .Li_Num
          sComplImovel = .Li_Compl
          sBairroImovel = .DescBairro
          Select Case .Ee_TipoEnd
            Case 0
                sEndEntrega = sEndImovel
                nNumEntrega = nNumImovel
                sComplEntrega = sComplImovel
                sBairroEntrega = sBairroImovel
                sCidadeEntrega = "JABOTICABAL"
                sCepEntrega = "14870-000"
                sUFEntrega = "SP"
            Case 1
                CarregaEndCidadao .CodPropPrincipal
            Case 2
                sEndEntrega = .Ee_NomeLog
                nNumEntrega = .Ee_NumImovel
                sComplEntrega = .Ee_Complemento
                Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & .Ee_Uf & "' AND CODCIDADE=" & .Ee_Cidade & " AND CODBAIRRO=" & .Ee_Bairro
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sBairroEntrega = !DescBairro
                    End If
                   .Close
                End With
                sCidadeEntrega = .Ee_Cidade
                sCepEntrega = .Ee_Cep
                sUFEntrega = .Ee_Uf
          End Select
    End If
End With

fim:
Liberado

End Sub

Private Sub CarregaEndCidadao(nCodigo As Long)

Sql = "SELECT CIDADAO.CODCIDADAO,vwLOGRADOUROCEP.ABREVTIPOLOG,vwLOGRADOUROCEP.ABREVTITLOG,"
Sql = Sql & "vwLOGRADOUROCEP.NOMELOGRADOURO,vwLOGRADOUROCEP.CEP, CIDADAO.NUMIMOVEL,"
Sql = Sql & "CIDADAO.COMPLEMENTO, CIDADAO.CODBAIRRO,CIDADAO.CODCIDADE, CIDADAO.SIGLAUF,"
Sql = Sql & "Cidade.DESCCIDADE , BAIRRO.DescBairro FROM CIDADAO INNER JOIN vwLOGRADOUROCEP ON "
Sql = Sql & "CIDADAO.CODLOGRADOURO = vwLOGRADOUROCEP.CODLOGRADOURO Inner Join BAIRRO ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND "
Sql = Sql & "CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO INNER JOIN "
Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = Cidade.CODCIDADE "
Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & nCodigo
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    sEndEntrega = SubNull(!NomeLogradouro)
    nNumEntrega = SubNull(!NUMIMOVEL)
    sComplEntrega = SubNull(!Complemento)
    sBairroEntrega = SubNull(!DescBairro)
    sCidadeEntrega = SubNull(!descCidade)
    sCepEntrega = SubNull(!Cep)
    sUFEntrega = SubNull(!Cep)
   .Close
End With

End Sub

Private Sub EmiteBoleto(nCodReduz As Long)
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataDam As String, sDataVencto As String
Dim sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, sQuadras As String, sLotes As String, nSeqLanc As Integer
Dim sUsuario As String, nNumDoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String, bGerado As Boolean
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date
Dim sBarra As String, sDigitavel2 As String, nValorDam As Double, nValorPrincDam As Double, nNumGuia As Long, nValorParc As Double
Dim sTipoEnd As String, nCodLanc As Integer, nCodTributo As Integer, sVencimento As String, sCPF As String
Dim bBoleto As Boolean
Dim sValor As String, dDataVencto As Date, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String

bBoleto = False
If MsgBox("Confirma Emissão dos boletos?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If


sTipoImposto = "ALUGUEL"
Select Case nCodReduz
    Case 1 To 99999
        xImovel.CarregaImovel nCodReduz
        sNumInsc = xImovel.Inscricao
        sCodReduz = txtCod.Text
        sNomeResp = xImovel.NomePropPrincipal
        sQuadra = xImovel.Li_Quadras
        sLote = xImovel.Li_Lotes
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
        sEndImovel = xImovel.Endereco
        nNumImovel = xImovel.Numero
        sComplImovel = xImovel.Complemento
        sBairroImovel = xImovel.Bairro
        
        sEndEntrega = xImovel.Ee_NomeLog
        nNumEntrega = xImovel.Ee_NumImovel
        sComplEntrega = xImovel.Ee_Complemento
        sBairroEntrega = xImovel.Ee_Bairro
        sCidadeEntrega = "JABOTICABAL"
        sUFEntrega = "SP"
        sCepEntrega = xImovel.Ee_Cep
        Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg "
        Sql = Sql & "FROM cidadao INNER JOIN proprietario ON cidadao.codcidadao = proprietario.codcidadao "
        Sql = Sql & "WHERE(proprietario.codreduzido = " & nCodReduz & ") AND (proprietario.tipoprop = 'P') AND (proprietario.principal = 1)"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sCPF = SubNull(!CPF)
                If Trim(sCPF) = "" Then
                   sCPF = SubNull(!Cnpj)
                End If
             Else
                sCPF = ""
             End If
             sRG = SubNull(!rg)
            .Close
        End With

    Case 100000 To 500000
        sNomeResp = grdMain.TextMatrix(grdMain.Row, 1)
        sNumInsc = nCodReduz
        sCodReduz = nCodReduz
        sLote = ""
        sQuadra = ""
        
        xImovel.RetornaEndereco nCodReduz, Mobiliario, Localizacao
        sEndImovel = xImovel.Endereco
        nNumImovel = xImovel.Numero
        sComplImovel = xImovel.Complemento
        sBairroImovel = xImovel.Bairro
        
        sEndEntrega = xImovel.Ee_NomeLog
        nNumEntrega = xImovel.Ee_NumImovel
        sComplEntrega = xImovel.Ee_Complemento
        sBairroEntrega = xImovel.Bairro
        sCidadeEntrega = xImovel.Cidade
        sUFEntrega = xImovel.UF
        sCepEntrega = xImovel.Ee_Cep
        Sql = "SELECT codigomob, inscestadual, cnpj, cpf From mobiliario WHERE codigomob = " & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sCPF = SubNull(!CPF)
                If Trim(sCPF) = "" Then
                   sCPF = SubNull(!Cnpj)
                End If
             Else
                sCPF = ""
             End If
             sRG = SubNull(!INSCESTADUAL)
            .Close
        End With
        
    Case 500000 To 800000
        sNomeResp = grdMain.TextMatrix(grdMain.Row, 1)
        sNumInsc = nCodReduz
        sCodReduz = nCodReduz
        sLote = ""
        sQuadra = ""
        
        xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
        sEndImovel = xImovel.Endereco
        nNumImovel = Val(xImovel.Numero)
        sComplImovel = xImovel.Complemento
        sBairroImovel = xImovel.Bairro
        
        sEndEntrega = sEndImovel
        nNumEntrega = nNumImovel
        sComplEntrega = sComplImovel
        sBairroEntrega = sBairroImovel
        sCidadeEntrega = xImovel.Cidade
        sUFEntrega = xImovel.UF
        sCepEntrega = xImovel.Cep
        
        Sql = "SELECT codcidadao,nomecidadao,cpf,cnpj,rg from cidadao WHERE CODCIDADAO=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sCPF = SubNull(!CPF)
                If Trim(sCPF) = "" Then
                   sCPF = SubNull(!Cnpj)
                End If
             Else
                sCPF = ""
             End If
             sRG = SubNull(!rg)
            .Close
        End With
End Select



'CARREGA GRID TEMPORARIO
nLanc = cmbTipo.ItemData(cmbTipo.ListIndex)

Select Case nLanc
    Case 33
        nCodTributo = 110
    Case 23
        nCodTributo = 111
    Case 34
        nCodTributo = 92
    Case 35
        nCodTributo = 101
    Case 60
        nCodTributo = 534
    Case 66
        nCodTributo = 528
    Case 77
        nCodTributo = 664
End Select



sFullTrib = cmbTipo.Text
nValorParc = CDbl(lblValorParc.Caption)

nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "insert boletoguiacapa(usuario,computer,sid,seq,codtributo,desctributo,valor) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & 1 & ","
Sql = Sql & cmbTipo.ItemData(cmbTipo.ListIndex) & ",'" & Left(cmbTipo.Text, 50) & "'," & Virg2Ponto(RemovePonto(lblValorTotal.Caption)) & ")"
cn.Execute Sql, rdExecDirect


'GERAÇÃO DOS DÉBITOS
For nParc = 1 To Val(lblNumParc.Caption)
    sVencimento = mskVenc(nParc).Text
    nAno = Year(CDate(sVencimento))
    nCompl = 0
    
    Sql = "delete from debitoparcela WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND seqlancamento=" & nSeq & " AND NUMPARCELA=" & nParc
    cn.Execute Sql, rdExecDirect
    Sql = "delete from debitotributo WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND seqlancamento=" & nSeq & " AND NUMPARCELA=" & nParc
    cn.Execute Sql, rdExecDirect
    Sql = "delete from parceladocumento WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND seqlancamento=" & nSeq & " AND NUMPARCELA=" & nParc
    cn.Execute Sql, rdExecDirect
    
    'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
    Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
    Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND ANOEXERCICIO=" & nAno
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux2!SEQMAXIMA) Then
               nSeq = 0
            Else
               nSeq = RdoAux2!SEQMAXIMA + 1
            End If
        Else
            nSeq = 0
        End If
       .Close
    End With
          
    'GRAVA DEBITOPARCELA
'    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,"
'    Sql = Sql & "USUARIO) VALUES(" & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
'    Sql = Sql & 3 & ",'" & Format(sVencimento, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "','"
 '   Sql = Sql & Left$(NomeDeLogin, 25) & "')"
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,"
    Sql = Sql & "USERID) VALUES(" & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ","
    Sql = Sql & 3 & ",'" & Format(sVencimento, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
    Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect

    'GRAVA DEBITOTRIBUTO
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nCodTributo & "," & Virg2Ponto(CStr(nValorParc)) & ")"
    cn.Execute Sql, rdExecDirect
    
    'RETORNA ULTIMO DOCUMENTO
    Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nNumDoc = RdoAux!maximo + 1
    RdoAux.Close
    
    'GRAVA NUMDOCUMENTO
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO) VALUES(" & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
    
    'GRAVA PARCELADOCUMENTO
    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & ")"
    cn.Execute Sql, rdExecDirect

    'EMITE BOLETO
    nValorGuia = nValorParc
    nNumGuia = nNumDoc

    sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
    sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
    sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)
    
    If bBoleto Then
        '**** GERADOR DE CÓDIGO DE BARRAS ********
        sNossoNumero = "2678478"
        sDigitavel = "001900000"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv
        sDataDam = sVencimento
        dDataBase = "07/10/1997"
        nFatorVencto = CDate(sDataDam) - dDataBase
        sQuintoGrupo = Format(nFatorVencto, "0000")
        sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
        sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "00000026784780"
        sBarra = sBarra & sNumDoc3 & "18"
        sDv = Trim(Calculo_DV11(sBarra))
        sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
        
        sDigitavel = sDigitavel & sDv & sQuintoGrupo
        
        sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
        sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
        sBarra = Gera2of5Str(sBarra)
    Else
        sValor = nValorGuia
        dDataVencto = CDate(sVencimento)
       ' nNumDoc = Val(sNumDoc2)
        sDadosLanc = "ALUGUEL"
        NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumDoc, nCodReduz)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
    
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        sBarra = StrBarra2
    End If
    '*******************************************
    
    Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
    Sql = Sql & "digitavel,codbarra,valorguia,obs,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nParc & "," & nCodReduz & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
    Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(sComplImovel, 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sComplImovel) & "','" & sUF & "','" & Mask(sFullTrib) & "','"
    Sql = Sql & CStr(nNumGuia) & "'," & nParc & "," & Val(lblNumParc.Caption) & ",'" & Format(sVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
    Sql = Sql & Virg2Ponto(Format(nValorParc, "#0.00")) & ",'" & "" & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
    cn.Execute Sql, rdExecDirect

Next

'EXIBE RELATORIO
If bBoleto Then
    If frmMdi.frTeste.Visible = False Then
        frmReport.ShowReport2 "BOLETOGUIA", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
    End If
Else
    If frmMdi.frTeste.Visible = False Then
        frmReport.ShowReport2 "BOLETOGUIA_V4", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
    Else
        frmReport.ShowReport2 "BOLETOGUIA_V4TMP", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
    End If
End If

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect


End Sub

