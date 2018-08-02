VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCancelParcAuto 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento Automático de Parcelamento"
   ClientHeight    =   3525
   ClientLeft      =   7905
   ClientTop       =   1365
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   11190
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   6
      Top             =   3915
      Width           =   1275
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   4215
      Width           =   1275
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   9840
      TabIndex        =   2
      ToolTipText     =   "Relatório com todos os parcelamentos cancelados automaticamente"
      Top             =   3090
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCancelParcAuto.frx":0000
      PICN            =   "frmCancelParcAuto.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCarregar 
      Height          =   345
      Left            =   8490
      TabIndex        =   0
      ToolTipText     =   "Executar o Cancelamento"
      Top             =   3090
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Carregar"
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
      MICON           =   "frmCancelParcAuto.frx":0176
      PICN            =   "frmCancelParcAuto.frx":0192
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
      Left            =   7140
      TabIndex        =   1
      ToolTipText     =   "Executar o Cancelamento"
      Top             =   3090
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Executar"
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
      MICON           =   "frmCancelParcAuto.frx":036C
      PICN            =   "frmCancelParcAuto.frx":0388
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdParc 
      Height          =   2955
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   $"frmCancelParcAuto.frx":0427
   End
   Begin prjChameleon.chameleonButton cmdCancelar 
      Height          =   315
      Left            =   6120
      TabIndex        =   7
      ToolTipText     =   "Cancelar o parcelamento"
      Top             =   9375
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Cance&lar  "
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
      MICON           =   "frmCancelParcAuto.frx":04D6
      PICN            =   "frmCancelParcAuto.frx":04F2
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
      Left            =   7350
      TabIndex        =   8
      ToolTipText     =   "Sair da Tela"
      Top             =   9375
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmCancelParcAuto.frx":0591
      PICN            =   "frmCancelParcAuto.frx":05AD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdOrigem 
      Height          =   1485
      Left            =   90
      TabIndex        =   9
      Top             =   7575
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   2619
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      BackColorBkg    =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Ano      |^Lanc.  |^Seq   |^Parc. |^Compl. |^Vencto.         |>Vl.Lançado      |>Valor Parcela    |<Situação                     "
   End
   Begin MSFlexGridLib.MSFlexGrid grdDestino 
      Height          =   1485
      Left            =   120
      TabIndex        =   10
      Top             =   5805
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2619
      _Version        =   393216
      Rows            =   1
      Cols            =   11
      FixedCols       =   0
      BackColorBkg    =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmCancelParcAuto.frx":061B
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "CANCELADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   90
      TabIndex        =   45
      Top             =   9375
      Visible         =   0   'False
      Width           =   4365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Parcelas........:"
      Height          =   225
      Index           =   5
      Left            =   210
      TabIndex        =   44
      Top             =   4560
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Parcelas Pagas.:"
      Height          =   225
      Index           =   4
      Left            =   2970
      TabIndex        =   43
      Top             =   4560
      Width           =   1635
   End
   Begin VB.Label lblQtdePago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4680
      TabIndex        =   42
      Top             =   4545
      Width           =   945
   End
   Begin VB.Label lblQtdeParc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   41
      Top             =   4545
      Width           =   1125
   End
   Begin VB.Label lblValorPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1695
      TabIndex        =   40
      Top             =   4875
      Width           =   1125
   End
   Begin VB.Label lblValorTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4680
      TabIndex        =   39
      Top             =   4875
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total Pago...:"
      Height          =   225
      Index           =   1
      Left            =   210
      TabIndex        =   38
      Top             =   4875
      Width           =   1485
   End
   Begin VB.Label lblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3060
      TabIndex        =   37
      Top             =   3915
      Width           =   5745
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido...:"
      Height          =   225
      Index           =   7
      Left            =   210
      TabIndex        =   36
      Top             =   3945
      Width           =   1485
   End
   Begin VB.Label lblDataParc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5160
      TabIndex        =   35
      Top             =   4215
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Parcelamento:"
      Height          =   225
      Index           =   2
      Left            =   3450
      TabIndex        =   34
      Top             =   4215
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo......:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   33
      Top             =   4245
      Width           =   1485
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
      Left            =   90
      TabIndex        =   32
      Top             =   3645
      Width           =   2910
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   " Parcelas de Origem"
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
      Left            =   90
      TabIndex        =   31
      Top             =   7305
      Width           =   2910
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Caption         =   " Parcelas de Destino"
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
      Index           =   2
      Left            =   120
      TabIndex        =   30
      Top             =   5535
      Width           =   2910
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Honorários....:"
      Height          =   225
      Index           =   6
      Left            =   210
      TabIndex        =   29
      Top             =   5205
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Juros Aplicado..:"
      Height          =   225
      Index           =   8
      Left            =   2970
      TabIndex        =   28
      Top             =   5205
      Width           =   1605
   End
   Begin VB.Label lblValorJuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4680
      TabIndex        =   27
      Top             =   5205
      Width           =   945
   End
   Begin VB.Label lblValorHonorario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1680
      TabIndex        =   26
      Top             =   5205
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor à Compensar.:"
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   10
      Left            =   5760
      TabIndex        =   25
      Top             =   5205
      Width           =   1605
   End
   Begin VB.Label lblValorCompensar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7260
      TabIndex        =   24
      Top             =   5205
      Width           =   1125
   End
   Begin VB.Label lblValorExpediente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4950
      TabIndex        =   23
      Top             =   5505
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelado por.:"
      Height          =   225
      Index           =   11
      Left            =   120
      TabIndex        =   22
      Top             =   9105
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelado em.:"
      Height          =   225
      Index           =   12
      Left            =   4980
      TabIndex        =   21
      Top             =   9105
      Width           =   1215
   End
   Begin VB.Label lblCanceladoPor 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1320
      TabIndex        =   20
      Top             =   9105
      Width           =   1215
   End
   Begin VB.Label lblDataCancel 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   6120
      TabIndex        =   19
      Top             =   9105
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Expediente....:"
      Height          =   225
      Index           =   9
      Left            =   3450
      TabIndex        =   18
      Top             =   5505
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor não Pago..........:"
      Height          =   225
      Index           =   3
      Left            =   2970
      TabIndex        =   17
      Top             =   4875
      Width           =   1605
   End
   Begin VB.Label lblPerc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7260
      TabIndex        =   16
      Top             =   4545
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%Últ.Parc.Compen..:"
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   13
      Left            =   5760
      TabIndex        =   15
      Top             =   4560
      Width           =   1605
   End
   Begin VB.Label lblValorCorrecao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7410
      TabIndex        =   14
      Top             =   4875
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Correção Aplic.:"
      Height          =   225
      Index           =   14
      Left            =   5700
      TabIndex        =   13
      Top             =   4875
      Width           =   1605
   End
   Begin VB.Label lblNP 
      Caption         =   "Label3"
      Height          =   285
      Left            =   1875
      TabIndex        =   12
      Top             =   10050
      Width           =   825
   End
   Begin VB.Label lblVlNComp 
      Caption         =   "Label4"
      Height          =   285
      Left            =   4035
      TabIndex        =   11
      Top             =   10005
      Width           =   870
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      Left            =   90
      TabIndex        =   4
      Top             =   3120
      Width           =   6765
   End
End
Attribute VB_Name = "frmCancelParcAuto"
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
    nValorTributo As Double
    nPercentual As Double
End Type
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, aTributo() As TRIBUTO
Dim sNumProc As String, nNumproc As Long, nAnoproc As Integer, nLinhaOriginal As Integer

Private Sub cmdCancelar_Click()
Dim nCodReduz As Long, nSeq As Integer, nSeq2 As Integer, nNumDoc As Long, nValorTxExp As Double, RdoAux2 As rdoResultset
Dim sData As String, sObs As String, sNumProc As String, nSeqObs As Integer

nCodReduz = Val(txtCod.Text)
sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)


Ocupado
'CANCELAMENTO DAS PARCELAS DE DESTINO
With grdDestino
    For x = 1 To .Rows - 1
        If Not IsDate(.TextMatrix(x, 9)) Then
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(x, 0) & " AND CODLANCAMENTO=" & .TextMatrix(x, 1) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(x, 2) & " AND NUMPARCELA=" & .TextMatrix(x, 3) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & .TextMatrix(x, 4)
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With



'ATUALIZAÇÃO DAS PARCELAS DE ORIGEM
With grdOrigem
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 6) <> "N/A" Then
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & Val(Left$(.TextMatrix(x, 8), 2)) & " WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(x, 0) & " AND CODLANCAMENTO=" & .TextMatrix(x, 1) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(x, 2) & " AND NUMPARCELA=" & .TextMatrix(x, 3) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & .TextMatrix(x, 4)
            cn.Execute Sql, rdExecDirect
        Else
            'CARREGA ORIGINAL PARCELA COMPLEMENTO
            Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(nLinhaOriginal, 0) & " AND CODLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 1) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 2) & " AND NUMPARCELA=" & .TextMatrix(nLinhaOriginal, 3) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(nLinhaOriginal, 4))
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                'GRAVA COMPLEMENTO PARCELA
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
'                Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,"
'                Sql = Sql & "DATAAJUIZA,VALORJUROS,NUMPROCESSO,USUARIO) VALUES(" & Val(txtCod.Text) & "," & !AnoExercicio & "," & !CodLancamento & ","
'                Sql = Sql & !SeqLancamento & "," & !NumParcela & "," & Val(grdOrigem.TextMatrix(x, 4)) & "," & Val(Left$(grdOrigem.TextMatrix(x, 8), 2)) & ",'" & Format(grdOrigem.TextMatrix(nLinhaOriginal, 5), "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',"
'                Sql = Sql & Val(SubNull(!CODMOEDA)) & "," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & "," & Val(SubNull(!numcertidao)) & "," & IIf(IsNull(!datainscricao), "Null", "'" & Format(!datainscricao, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!dataajuiza), "Null", "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'") & "," & !ValorJuros & ",'"
'                Sql = Sql & txtNumProc.Text & "','GTI')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,"
                Sql = Sql & "DATAAJUIZA,VALORJUROS,NUMPROCESSO,USERID) VALUES(" & Val(txtCod.Text) & "," & !AnoExercicio & "," & !CodLancamento & ","
                Sql = Sql & !SeqLancamento & "," & !NumParcela & "," & Val(grdOrigem.TextMatrix(x, 4)) & "," & Val(Left$(grdOrigem.TextMatrix(x, 8), 2)) & ",'" & Format(grdOrigem.TextMatrix(nLinhaOriginal, 5), "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',"
                Sql = Sql & Val(SubNull(!CODMOEDA)) & "," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & "," & Val(SubNull(!numcertidao)) & "," & IIf(IsNull(!datainscricao), "Null", "'" & Format(!datainscricao, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!dataajuiza), "Null", "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'") & "," & Val(SubNull(!ValorJuros)) & ",'"
                Sql = Sql & txtNumProc.Text & "',236)"
                cn.Execute Sql, rdExecDirect
            
               'GRAVA OBS PARCELA
                Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & !AnoExercicio
                Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
                Sql = Sql & " AND CODCOMPLEMENTO=" & Val(grdOrigem.TextMatrix(x, 4))
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
                sObs = "Débito remanescente do parcelamento com processo número " & txtNumProc.Text & " com percentual remanescente de " & lblPerc.Caption & "."
'                Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & Val(txtCod.Text) & "," & !AnoExercicio & ","
'                Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & Val(grdOrigem.TextMatrix(x, 4)) & "," & nSeq2 & ",'" & sObs & "','" & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
                Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & Val(txtCod.Text) & "," & !AnoExercicio & ","
                Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & Val(grdOrigem.TextMatrix(x, 4)) & "," & nSeq2 & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "')"
                cn.Execute Sql, rdExecDirect
            End With
            
            'CARREGA ORIGINAL TRIBUTO COMPLEMENTO
            Sql = "SELECT sum(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(nLinhaOriginal, 0) & " AND CODLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 1) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 2) & " AND NUMPARCELA=" & .TextMatrix(nLinhaOriginal, 3) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(nLinhaOriginal, 4)) & " AND CODTRIBUTO <>3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If Not IsNull(!soma) Then
                   nValorTotal = !soma
                Else
                    nValorTotal = 0
                End If
              .Close
           End With
           nValorComplemento = CDbl(grdOrigem.TextMatrix(x, 7))
           ReDim aTributo(0)
           Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND "
           Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(nLinhaOriginal, 0) & " AND CODLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 1) & " AND "
           Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 2) & " AND NUMPARCELA=" & .TextMatrix(nLinhaOriginal, 3) & " AND "
           Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(nLinhaOriginal, 4)) & " AND CODTRIBUTO <>3"
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          With RdoAux2
               If .RowCount > 0 Then
               nCodReduz = !CODREDUZIDO
               nAno = !AnoExercicio
               nLanc = !CodLancamento
               nSeq = !SeqLancamento
               nParc = !NumParcela
               nCompl = !CODCOMPLEMENTO
               Do Until .EOF
                   ReDim Preserve aTributo(UBound(aTributo) + 1)
                   aTributo(UBound(aTributo)).nCodTributo = !CodTributo
                   aTributo(UBound(aTributo)).nPercentual = (!ValorTributo * 100) / nValorTotal
                  .MoveNext
               Loop
               End If
              .Close
           End With
            
           For TY = 1 To UBound(aTributo)
               aTributo(TY).nValorTributo = Format((nValorComplemento * aTributo(TY).nPercentual) / 100, "#0.00")
               'GRAVA COMPLEMENTO TRIBUTO
               Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
               Sql = Sql & "CODTRIBUTO,VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAno & "," & nLanc & ","
               Sql = Sql & nSeq & "," & nParc & "," & Val(grdOrigem.TextMatrix(x, 4)) & "," & aTributo(TY).nCodTributo & "," & Virg2Ponto(CStr(aTributo(TY).nValorTributo)) & " )"
               cn.Execute Sql, rdExecDirect
           Next
        End If
    Next
End With

'CANCELAMENTO DAS PARCELAS BLOQUEADAS
Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE NUMPROCESSO='" & sNumProc & "' AND STATUSLANC=18"
cn.Execute Sql, rdExecDirect

'CANCELAMENTO DO PROCESSO
Sql = "UPDATE PROCESSOREPARC SET CANCELADO=1,DATACANCEL='" & Format(Now, "mm/dd/yyyy") & "',FUNCIONARIOCANCEL='" & NomeDeLogin & "' WHERE ANOPROC=" & nAnoproc & " AND NUMPROC=" & nNumproc
cn.Execute Sql, rdExecDirect


'***INTEGRATIVA ****
If frmMdi.frTeste.Visible = False Then
    ConectaIntegrativa
    
    'GRAVA NA TABELA ACORDOSTATUS
    Sql = "insert acordostatus(idacordo,anoacordo,dtocorrencia,ocorrencia,dtgeracao) values("
    Sql = Sql & nNumproc & "," & nAnoproc & ",'" & Format(Now, "mm/dd/yyyy") & "','PARCEL.CANCELADO','" & Format(Now, "mm/dd/yyyy") & "')"
    cnInt.Execute Sql, rdExecDirect
    
    cnInt.Close
End If
'*******************

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeqObs = 1
    Else
        nSeqObs = !maximo + 1
    End If
   .Close
End With

sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
sObs = "PARCELAMENTO " & (nNumproc & RetornaDVProcesso(nNumproc) & "/" & nAnoproc) & " FOI CANCELADO AUTOMÁTICAMENTO PELO SISTEMA GTI POR ESTAR COM " & grdParc.TextMatrix(grdParc.Row, 3) & " PARCELAS EM ATRASO."
'Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USUARIO,DATAOBS,OBS) VALUES(" & nCodReduz & "," & nSeqObs & ",'"
'Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USERID,DATAOBS,OBS) VALUES(" & nCodReduz & "," & nSeqObs & ","
Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
cn.Execute Sql, rdExecDirect
'GRAVA DETALHES DO CANCELAMENTO
If nValorComplemento = "" Then nValorComplemento = 0
Sql = "INSERT PARCELAMENTOCANCEL(CODIGO,ANOPROC,NUMPROC,NUMPROCESSO,DATAPARC,PARCATRASADA,QTDEPARC,QTDEPAGA,VALORPAGO,VALORNPAGO,VALORCOMP,DATACANCEL) VALUES("
Sql = Sql & nCodReduz & "," & nAnoproc & "," & nNumproc & ",'" & CStr(nNumproc) & "-" & RetornaDVProcesso(nNumproc) & "/" & CStr(nAnoproc) & "','" & Format(grdParc.TextMatrix(grdParc.Row, 8), "mm/dd/yyyy") & "'," & Val(grdParc.TextMatrix(grdParc.Row, 3)) & ","
'Sql = Sql & Val(.TextMatrix(x, 4)) & "," & Val(.TextMatrix(x, 5)) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 6))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 7))) & "," & Virg2Ponto(CStr(nValorComplemento)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
Sql = Sql & Val(grdParc.TextMatrix(grdParc.Row, 4)) & "," & Val(grdParc.TextMatrix(grdParc.Row, 5)) & "," & Virg2Ponto(RemovePonto(grdParc.TextMatrix(grdParc.Row, 6))) & "," & 0 & "," & Virg2Ponto(CStr(nValorComplemento)) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect

Liberado
'MsgBox "O cancelamento do reparcelamento foi executado com sucesso.", vbExclamation, "Atenção"
grdDestino.Rows = 1
grdOrigem.Rows = 1

End Sub

Private Sub cmdCarregar_Click()
Dim x As Integer, nContador As Integer

Ocupado
grdParc.Enabled = False
cmdBaixa.Enabled = False
grdParc.Rows = 1: grdOrigem.Rows = 1
Sql = "SELECT debitoparcela.codreduzido, COUNT(debitoparcela.codreduzido) AS contador, processoreparc.numprocesso, processoreparc.qtdeparcela "
Sql = Sql & "FROM debitoparcela INNER JOIN processoreparc ON debitoparcela.numprocesso = processoreparc.numprocesso WHERE processoreparc.novo = 1 AND "
Sql = Sql & "debitoparcela.codlancamento = 20 AND (debitoparcela.statuslanc = 3 or debitoparcela.statuslanc = 18) AND datediff(day,debitoparcela.datavencimento, GETDATE())>30 "
Sql = Sql & "GROUP BY debitoparcela.codreduzido, processoreparc.numprocesso, processoreparc.qtdeparcela Having (Count(debitoparcela.CODREDUZIDO) > 2) ORDER BY debitoparcela.codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    x = 1: nContador = .RowCount
    If .RowCount = 0 Then
        MsgBox "Não existem parcelamentos vencidos a mais de 100 dias.", vbExclamation, "Atenção"
    End If
    Do Until .EOF
'    If !CODREDUZIDO <> 538957 Then
'        GoTo PROXIMO
'    End If
    If x > 10 Then Exit Do
        If !CODREDUZIDO < 100000 Then
           Sql = "SELECT NOMECIDADAO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & !CODREDUZIDO
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
               sNome = SubNull(!nomecidadao)
              .Close
           End With
        ElseIf !CODREDUZIDO >= 100000 And !CODREDUZIDO < 300000 Then
           Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & !CODREDUZIDO
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
               sNome = SubNull(!RazaoSocial)
              .Close
           End With
        Else
           Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & !CODREDUZIDO
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
               sNome = SubNull(!nomecidadao)
              .Close
           End With
        End If
        nNumproc = Val(Left$(!numprocesso, InStr(1, !numprocesso, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(!numprocesso, 4))
        sNumProc = nNumproc & "/" & nAnoproc
        
        lblStatus.Caption = "Carregando processo nº " & (nNumproc & RetornaDVProcesso(nNumproc) & "/" & nAnoproc) & " (" & x & " de " & nContador & ")"
        If cGetInputState() <> 0 Then DoEvents
        
        grdParc.AddItem !CODREDUZIDO & Chr(9) & sNome & Chr(9) & (nNumproc & RetornaDVProcesso(nNumproc) & "/" & nAnoproc) & Chr(9) & !contador & " de " & !qtdeparcela
        CarregaProcessos sNumProc
        x = x + 1
proximo:
       .MoveNext
       'Exit Do
    Loop
   .Close
End With
cmdBaixa.Enabled = True
grdParc.Enabled = True
Liberado
End Sub

Private Sub cmdPrint_Click()
Dim sNome As String, sNumProc As String, nNumproc As Long, nAnoproc As Integer

frmReport.ShowReport "PROCESSOCANCELADO", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub cmdBaixa_Click()
Dim sNumProc As String, nNumproc As Long, nAnoproc As Integer, x As Integer, nCodReduz As Long, nQtdePago As Integer, sObs As String, nSeq As Integer
Dim nStatus As Integer, RdoAux2 As rdoResultset, aTributo() As TRIBUTO, nValorComplemento As Double

If grdParc.Rows = 1 Then
    MsgBox "Nada a cancelar.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Deseja cancelar estes processos?", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE CANCELAMENTO") = vbNo Then Exit Sub
With grdParc
    For x = 1 To .Rows - 1
       .Row = x
       .ColSel = 8
       txtCod.Text = .TextMatrix(x, 0)
       txtNumProc.Text = .TextMatrix(x, 2)
       txtNumProc_LostFocus
       cmdCancelar_Click
       If cGetInputState() <> 0 Then DoEvents
    Next
End With

grdParc.Rows = 1: grdOrigem.Rows = 1
lblStatus.Caption = "Todos os Processos acima foram cancelados."

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim sNome As String, x As Integer
Centraliza Me

End Sub

Private Sub CarregaProcessos(sNumProc As String)
On Error GoTo Erro

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nLinha As Integer
Dim nValorLanc As Double, nQtdePago As Integer, nQtdeParc As Integer
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date
Dim dDataPag As Date
Dim nValorPago As Double, nValorNaoPago As Double
Dim nSomaValorTributo As Double
Dim nSomaPago As Double, nSomaNaoPago As Double, nSomaLancado As Double
Dim nTotalACompensar As Double
Dim nTotalAtual As Double
Dim nValorAChecar As Double
Dim nSobra As Double
Dim nCodCompl As Integer
Dim x As Integer
Dim dDataPagto As Date, sDataPagto As String
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, Achou As Boolean

nLinha = grdParc.Rows - 1
ReDim aDebito(0)
'dDataPag = CDate(lblDataParc.Caption)
'grdOrigem.Rows = 1: grdDestino.Rows = 1

Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & grdParc.TextMatrix(nLinha, 0) & " AND CODLANCAMENTO=20 AND (STATUSLANC=2 or statuslanc=7) AND NUMPROCESSO='" & sNumProc & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    nQtdePago = .RowCount
   .Close
End With

Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & sNumProc & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    dDataPag = !datareparc
    nQtdeParc = Val(SubNull(!qtdeparcela))
    Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
    Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
    Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
    Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
    Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
    Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND (STATUSLANC=2 OR STATUSLANC=7)"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            If Not IsNull(!ValorTributo) Then
                nValorPago = !ValorTributo
            Else
                nValorPago = 0
            End If
        Else
            nValorPago = 0
        End If
       .Close
    End With
   .Close
End With
grdParc.TextMatrix(nLinha, 4) = nQtdeParc
grdParc.TextMatrix(nLinha, 5) = nQtdePago
grdParc.TextMatrix(nLinha, 6) = Round(nValorPago, 2)
grdParc.TextMatrix(nLinha, 8) = Format(dDataPag, "dd/mm/yyyy")

Exit Sub
Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub grdParc_Click()
CarregaOrigem
End Sub

Private Sub CarregaOrigem()
    On Error GoTo Erro

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double, nCodReduz As Long
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date
Dim dDataPag As Date
Dim nValorPago As Double, nValorNaoPago As Double
Dim nSomaValorTributo As Double, nSomaValorTributoJuros As Double
Dim nSomaPago As Double, nSomaNaoPago As Double, nSomaLancado As Double
Dim nTotalACompensar As Double, nSomaCorrecao As Double
Dim nTotalAtual As Double
Dim nValorAChecar As Double
Dim nSobra As Double
Dim nCodCompl As Integer
Dim x As Integer, nPerc As Double, nValorPerc As Double
Dim dDataPagto As Date, sDataPagto As String
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, Achou As Boolean


If txtNumProc.Text = "" Then
    Exit Sub
End If

nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
nAnoproc = Right$(txtNumProc.Text, 4)
sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
nCodReduz = Val(grdParc.TextMatrix(grdParc.Row, 0))

Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=20 AND (STATUSLANC=2 or statuslanc=7) AND NUMPROCESSO='" & sNumProc & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    lblQtdePago.Caption = .RowCount
    .Close
End With

ReDim aDebito(0)
dDataPag = CDate(Format(grdParc.TextMatrix(grdParc.Row, 8), "dd/mm/yyyy"))
grdOrigem.Rows = 1: grdDestino.Rows = 1
Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & sNumProc & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorPago = 0
    nSomaPago = 0: nSomaNaoPago = 0: nSomaLancado = 0
    Do Until .EOF
         dDataVencto = Format(!DATADEBASE, "dd/mm/yyyy")
         'BUSCA VALOR LANÇADO
         Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
         Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
         Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
         Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
         Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
         Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3 AND CODTRIBUTO<>90  AND CODTRIBUTO<>585  AND CODTRIBUTO<>587"
         Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            nValorLanc = !ValorTributo
            If (dDataPag > dDataVencto) Then
'                nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
'                nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
'                nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
            Else
                nValorCorrecao = 0
                nValorJuros = 0
                nValorMulta = 0
            End If
            nSomaValorTributo = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
            .Close
         End With
            
         Sql = "SELECT jurosapl, honorario From destinoreparc WHERE codreduzido = " & nCodReduz & " AND NUMPROCESSO='" & CStr(nNumproc) & "/" & CStr(nAnoproc) & "'"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            nSomaValorTributoJuros = !jurosapl
            .Close
         End With
            
         Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
         Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
         Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
         Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
         Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
         Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=587"
         Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            If .RowCount > 0 Then
                nSomaCorrecao = !ValorTributo
            Else
                nSomaCorrecao = 0
            End If
            .Close
         End With
            
            
         'BUSCA VALORPAGO
         Sql = "SELECT VALORPAGOREAL,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
         Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !numsequencia
         Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND SEQPAG=0"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
              If .RowCount > 0 Then
                  nValorPago = !valorpagoreal
                  dDataPagto = !DataPagamento
                  sDataPagto = Format(!DataPagamento, "dd/mm/yyyy")
              Else
                  Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
                  Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
                  Sql = Sql & "WHERE CODREDUZIDO=" & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO = " & RdoAux!AnoExercicio
                  Sql = Sql & " AND CODLANCAMENTO=" & RdoAux!CodLancamento & " AND NUMPARCELA=" & RdoAux!NumParcela & " AND SEQLANCAMENTO=" & RdoAux!numsequencia
                  Sql = Sql & " AND CODCOMPLEMENTO=" & RdoAux!CODCOMPLEMENTO & " AND VALORPAGO>0"
                  Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                  With RdoAux2
                       If .RowCount > 0 Then
                            nValorPago = !ValorPago
                            sDataPagto = "Pago sem Data"
                       Else
                            nValorPago = 0
                            dDataPagto = CDate("01/01/1900")
                            sDataPagto = "Não Pago"
                       End If
                      .Close
                  End With
                  
              End If
             .Close
         End With
           
         If nValorPago > 0 Then
            'BUSCA TAXA
            Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO "
            Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
            Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !numsequencia
            Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    'nValorPago = nValorPago + !VALORTRIBUTO
                    If nValorPago > 0 Then
                        nSomaLancado = nSomaLancado + nSomaValorTributo + !ValorTributo
                    End If
                Else
                    If nValorPago > 0 Then
                        nSomaLancado = nSomaLancado + nSomaValorTributo
                    End If
                End If
            End With
            nSomaPago = nSomaPago + nValorPago
         End If
                            
        
         grdDestino.AddItem !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!numsequencia, "00") & Chr(9) & _
         Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & FormatNumber(nSomaValorTributoJuros, 2) & Chr(9) & _
         FormatNumber(nSomaCorrecao, 2) & Chr(9) & sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
'         nSomaLancado = nSomaLancado + nSomaValorTributo
        .MoveNext
    Loop
   .Close
End With

'nSomaLancado = 0
'lblValorPago.Caption = Format(nSomaLancado, "#0.00")


'PREENCHE GRID DE ORIGEM
bVenctoNulo = False
Sql = "SELECT * FROM vwCNSREPARCELAMENTOO WHERE NUMPROCESSO='" & sNumProc & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
    
        'SE ALGUMA PARCELA NÃO FOR LOCALIZADA NÃO PERMITE O CANCELAMENTO
        If IsNull(!DataVencimento) Then bVenctoNulo = True
        
        'CARREGA OS TRIBUTOS DE CADA UM DOS LANCAMENTOS
        Set qd.ActiveConnection = cn
        On Error Resume Next
        RdoAux3.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = !CODREDUZIDO
        qd(1) = !CODREDUZIDO 'codigo
        qd(2) = !AnoExercicio
        qd(3) = !AnoExercicio  'ano
        qd(4) = !CodLancamento
        qd(5) = !CodLancamento 'lancamento
        qd(6) = !numsequencia
        qd(7) = !numsequencia 'sequencia
        qd(8) = !NumParcela
        qd(9) = !NumParcela 'numparcela
        qd(10) = !CODCOMPLEMENTO
        qd(11) = !CODCOMPLEMENTO 'complemento
        qd(12) = 1
        qd(13) = 99 'statuslanc
        qd(14) = Format(dDataPag, "mm/dd/yyyy")
        qd(15) = NomeDoUsuario
        
        Set RdoAux3 = qd.OpenResultset(rdOpenKeyset)
        With RdoAux3
            Do Until .EOF
                'CARREGA MATRIZ DE DÉBITO
                nEval = UBound(aDebito)
                Achou = False
                For x = 1 To nEval
                    If aDebito(x).nCodReduzido = !CODREDUZIDO And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                       aDebito(x).nSeq = !SeqLancamento And _
                       aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
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
                   aDebito(nEval).nLanc = !CodLancamento
                   aDebito(nEval).nSeq = !SeqLancamento
                   aDebito(nEval).nParc = !NumParcela
                   aDebito(nEval).nCompl = !CODCOMPLEMENTO
                   aDebito(nEval).nSituacao = !statuslanc
                   aDebito(nEval).sSituacao = !Situacao
                   aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                   aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                   aDebito(nEval).nValorAtual = !ValorTotal
                Else
                    'SE ENCONTRAR ADICIONAR O VALOR AO JA EXISTENTE
                    If !statuslanc = 3 Or !statuslanc = 4 Or !statuslanc = 6 Then
                        aDebito(x).nValorAtual = aDebito(x).nValorAtual + !ValorTotal
                    End If
                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                End If
               .MoveNext
            Loop
           .Close
        End With
       .MoveNext
    Loop
End With
'ADICIONA OS DEBITOS AO GRID DE ORIGEM
nSomaNaoPago = 0
For x = 1 To UBound(aDebito)
    With aDebito(x)
        nSomaNaoPago = nSomaNaoPago + .nValorAtual
        grdOrigem.AddItem .nAno & Chr(9) & Format(.nLanc, "00") & Chr(9) & Format(.nSeq, "00") & Chr(9) & _
        Format(.nParc, "00") & Chr(9) & Format(.nCompl, "00") & Chr(9) & .sVencto & Chr(9) & FormatNumber(.nValorTributo, 2) & Chr(9) & _
        FormatNumber(.nValorAtual, 2) & Chr(9) & "03-NÃO PAGO"
    End With
Next
nSomaPago = CDbl(lblValorPago.Caption)
'lblValorNPago.Caption = FormatNumber(nSomaNaoPago - nSomaPago, 2)

'VERIFICA SE TEM COMPENSAÇÃO
If Val(lblValorCompensar.Caption) > 0 Then
    nTotalACompensar = CDbl(lblValorCompensar.Caption)
    nTotalAtual = 0
    
    nSobra = nTotalACompensar
    With grdOrigem
        For x = 1 To .Rows - 1
             nValorAChecar = CDbl(.TextMatrix(x, 7))
             nTotalAtual = nTotalAtual + nValorAChecar
             If nSobra > nValorAChecar Then
                .TextMatrix(x, 8) = "06-COMPENSADO"
                nSobra = nSobra - nValorAChecar
             ElseIf nSobra > 0 And nSobra < nValorAChecar Then
                 nValorAChecar = CDbl(.TextMatrix(x, 7))
                 nPerc = 1 - (nSobra / nValorAChecar)
                 nValorPerc = FormatNumber(nValorAChecar * nPerc, 2)
                 lblPerc.Caption = FormatNumber(nPerc * 100, 2) & "%"
                 nValorAChecar = CDbl(.TextMatrix(x, 6))
'                 nPerc = 1 - (nSobra / nValorAChecar)
                 nValorPerc = FormatNumber(nValorAChecar * nPerc, 2)
                .TextMatrix(x, 8) = "06-COMPENSADO"
                 'busca o novo codigo do complemento
                 Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXCOMPL FROM DEBITOPARCELA WHERE "
                 Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & .TextMatrix(x, 0) & " AND "
                 Sql = Sql & "CODLANCAMENTO=" & .TextMatrix(x, 1) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 2) & " AND "
                 Sql = Sql & "NUMPARCELA=" & .TextMatrix(x, 3)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 nCodCompl = RdoAux!MAXCOMPL + 1
                 RdoAux.Close
                 'cria uma parcela de compensação
                 nLinhaOriginal = x
                .AddItem nCodReduz & Chr(9) & .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & _
                 Format(nCodCompl, "00") & Chr(9) & .TextMatrix(x, 5) & Chr(9) & "N/A" & Chr(9) & _
                 FormatNumber(nValorPerc, 2) & Chr(9) & "03-NÃO PAGO"
                 lblValorExt.Caption = FormatNumber((nValorAChecar - (nSobra)), 2)
                 nSobra = 0
             Else
                .TextMatrix(x, 8) = "03-NÃO PAGO"
             End If
        Next
        
    End With
Else
    'SE NÃO TEM COMPENSAÇÃO, QUER DIZER QUE NENHUMA PARCELA FOI PAGA
    'NESTE CASO BASTA CANCELAR TODAS AS PARCELAS
    With grdOrigem
        For x = 1 To .Rows - 1
            .TextMatrix(x, 8) = "03-NÃO PAGO"
        Next
    End With
End If

nValorAChecar = 0: nValorNaoPago = 0
For x = 1 To grdOrigem.Rows - 1
    If grdOrigem.TextMatrix(x, 8) = "06-COMPENSADO" Then
        nValorAChecar = nValorAChecar + grdOrigem.TextMatrix(x, 7)
    ElseIf grdOrigem.TextMatrix(x, 8) = "03-NÃO PAGO" And grdOrigem.TextMatrix(x, 7) <> "N/A" Then
        nValorNaoPago = nValorNaoPago + grdOrigem.TextMatrix(x, 7)
    End If
Next
lblNP.Caption = FormatNumber(nValorAChecar, 2)
lblVlNComp.Caption = FormatNumber(nValorNaoPago, 2)
With grdOrigem
    If .TextMatrix(.Rows - 1, 8) = "06-COMPENSADO" Then
        If Val(lblValorNaoPago.Caption) > 0 Then
             .AddItem .TextMatrix(.Rows - 1, 0) & Chr(9) & .TextMatrix(.Rows - 1, 1) & Chr(9) & .TextMatrix(.Rows - 1, 2) & Chr(9) & .TextMatrix(.Rows - 1, 3) & Chr(9) & _
            .TextMatrix(.Rows - 1, 4) & Chr(9) & Format(nCodCompl + 1, "00") & Chr(9) & Format(mskDataParc.Text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
              FormatNumber(CDbl(lblValorNaoPago.Caption), 2) & Chr(9) & "03-NÃO PAGO"
        End If
    Else
        If CDbl(lblValorNaoPago.Caption) > CDbl(lblVlNComp.Caption) Then
            .TextMatrix(.Rows - 1, 7) = FormatNumber(CDbl(lblValorNaoPago.Caption) - CDbl(lblVlNComp.Caption), 2)
''             .TextMatrix(.Rows - 1, 8) = FormatNumber(CDbl(lblValorExt.Caption), 2)
       Else
           If lblValorExt.Caption > 0 Then
                .TextMatrix(.Rows - 1, 7) = FormatNumber(CDbl(lblValorExt.Caption), 2)
           End If
       End If
    End If
End With

With grdOrigem
     If .TextMatrix(.Rows - 1, 6) = "N/A" Then
        .FillStyle = flexFillRepeat
        .Row = .Rows - 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &H9FFFC0
     End If
End With


Exit Sub
Erro:
MsgBox Err.Description
'Resume Next


End Sub

Private Sub txtNumProc_LostFocus()
Dim nValorPago As Double, nNovo As Integer, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nQtdeParc As Integer, nValorCalc As Double, x As Integer, nValorCorrecao As Double
On Error Resume Next
Ocupado
nValorPago = 0
If Trim$(txtNumProc.Text) <> "" Then
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumproc = Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))
        nAnoproc = Val(Right$(txtNumProc.Text, 4))
        lblNumProc.Caption = nNumproc
        lblAnoProc.Caption = nAnoproc
        Sql = "SELECT NUMPROC,ANOPROC,DATAREPARC,QTDEPARCELA,NOVO,CANCELADO,DATACANCEL,FUNCIONARIOCANCEL FROM PROCESSOREPARC  WHERE CODIGORESP=" & Val(txtCod.Text) & " AND NUMPROC=" & nNumproc & " AND ANOPROC=" & nAnoproc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
                lblValorNPago.Caption = "0,00"
                lblDataParc.Caption = ""
                lblValorPago.Caption = "0,00"
                lblQtdePago.Caption = "0"
                lblQtdeParc.Caption = "0"
                txtNumProc.SetFocus
                Liberado
                Exit Sub
            Else
                'aqui
                lblDataParc.Caption = Format(!datareparc, "dd/mm/yyyy")
                lblQtdeParc.Caption = !qtdeparcela
                lblNovo.Caption = IIf(IsNull(!Novo), 0, 1)
                nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
                nAnoproc = Right$(txtNumProc.Text, 4)
                sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
                
                Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND (STATUSLANC=2 or statuslanc=7) AND NUMPROCESSO='" & sNumProc & "'"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    lblQtdePago.Caption = .RowCount
                    .Close
                End With
                lblCancel.Visible = !Cancelado
                If !Cancelado = True Then
                    lblDataCancel.Caption = Format(!DataCancel, "dd/mm/yyyy")
                    lblCanceladoPor.Caption = SubNull(!FUNCIONARIOCANCEL)
                Else
                    lblDataCancel.Caption = ""
                    lblCanceladoPor.Caption = ""
                End If
            End If
           .Close
        End With
Ini:
        If lblQtdePago.Caption > 0 Then
            CarregaGrid
            nQtdeParc = 0: nValorCalc = 0: nValorPago = 0: nValorCorrecao = 0
            For x = 1 To grdDestino.Rows - 1
                nValorCalc = nValorCalc + CDbl(grdDestino.TextMatrix(x, 6))
                If Val(grdDestino.TextMatrix(x, 10)) > 0 Then
                    nValorPago = nValorPago + CDbl(grdDestino.TextMatrix(x, 6))
                    nValorCorrecao = nValorCorrecao + CDbl(grdDestino.TextMatrix(x, 8))
                    nQtdeParc = nQtdeParc + 1
                End If
            Next
            
            lblValorCorrecao.Caption = FormatNumber(nValorCorrecao, 2)
            Sql = "SELECT jurosapl, honorario From destinoreparc WHERE codreduzido = " & Val(txtCod.Text) & " AND NUMPROCESSO='" & CStr(nNumproc) & "/" & CStr(nAnoproc) & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!SOMAPRINCIPAL) Then
    '                Corrige
    '                GoTo Ini
                End If
                lblValorHonorario.Caption = FormatNumber(!honorario * nQtdeParc, 2)
                lblValorJuros.Caption = FormatNumber(!jurosapl * nQtdeParc, 2)
                
                lblValorPago.Caption = FormatNumber(nValorPago + CDbl(lblValorJuros.Caption) + CDbl(lblValorCorrecao.Caption), 2)
                lblValorCompensar.Caption = FormatNumber(CDbl(lblValorPago.Caption) - CDbl(lblValorJuros.Caption) - CDbl(lblValorCorrecao.Caption), 2)
                lblValorTotal.Caption = FormatNumber(nValorCalc - nValorPago, 2)
                
               .MoveNext
               .Close
            End With
        Else
            lblValorHonorario.Caption = 0
            lblValorExpediente.Caption = 0
            lblValorJuros.Caption = 0
            lblValorCompensar.Caption = 0
            lblValorPago.Caption = 0
            lblValorTotal.Caption = 0
        End If
        CarregaGrid
    Else
        MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
        lblValorNPago.Caption = "0,00"
        lblDataParc.Caption = ""
        lblValorPago.Caption = "0,00"
        lblQtdePago.Caption = "0"
        lblQtdeParc.Caption = "0"
        txtNumProc.SetFocus
    End If
End If
Liberado

End Sub

Private Sub CarregaGrid()
    On Error GoTo Erro

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date
Dim dDataPag As Date
Dim nValorPago As Double, nValorNaoPago As Double
Dim nSomaValorTributo As Double, nSomaValorTributoJuros As Double
Dim nSomaPago As Double, nSomaNaoPago As Double, nSomaLancado As Double
Dim nTotalACompensar As Double, nSomaCorrecao As Double
Dim nTotalAtual As Double
Dim nValorAChecar As Double
Dim nSobra As Double
Dim nCodCompl As Integer
Dim x As Integer, nPerc As Double, nValorPerc As Double
Dim dDataPagto As Date, sDataPagto As String
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, Achou As Boolean

ReDim aDebito(0)
dDataPag = CDate(lblDataParc.Caption)
grdOrigem.Rows = 1: grdDestino.Rows = 1
Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & sNumProc & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nValorPago = 0
    nSomaPago = 0: nSomaNaoPago = 0: nSomaLancado = 0
    Do Until .EOF
'         lblDataProc.Caption = Format(!DATAPROCESSO, "dd/mm/yyyy")
         dDataVencto = Format(!DATADEBASE, "dd/mm/yyyy")
      '   dDataPag = Format(!DATAREPARC, "dd/mm/yyyy")
         dDataPag = CDate(lblDataParc.Caption)
         'BUSCA VALOR LANÇADO
         Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
         Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
         Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
         Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
         Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
         Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3 AND CODTRIBUTO<>90  AND CODTRIBUTO<>585  AND CODTRIBUTO<>587 AND CODTRIBUTO<>609"
         Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            nValorLanc = !ValorTributo
            If (dDataPag > dDataVencto) Then
'                nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
'                nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
'                nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
            Else
                nValorCorrecao = 0
                nValorJuros = 0
                nValorMulta = 0
            End If
            nSomaValorTributo = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
            .Close
         End With
            
         Sql = "SELECT jurosapl, honorario From destinoreparc WHERE codreduzido = " & Val(txtCod.Text) & " AND NUMPROCESSO='" & CStr(nNumproc) & "/" & CStr(nAnoproc) & "'"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            If Not IsNull(!jurosapl) Then
                nSomaValorTributoJuros = !jurosapl
            Else
                nSomaValorTributoJuros = 0
            End If
            .Close
         End With
            
         Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
         Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
         Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
         Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
         Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
         Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=587"
         Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            If .RowCount > 0 Then
                nSomaCorrecao = !ValorTributo
            Else
                nSomaCorrecao = 0
            End If
            .Close
         End With
            
            
         'BUSCA VALORPAGO
         Sql = "SELECT VALORPAGOREAL,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
         Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !numsequencia
         Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND SEQPAG=0"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
              If .RowCount > 0 Then
                  nValorPago = !valorpagoreal
                  dDataPagto = !DataPagamento
                  sDataPagto = Format(!DataPagamento, "dd/mm/yyyy")
              Else
                  Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
                  Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
                  Sql = Sql & "WHERE CODREDUZIDO=" & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO = " & RdoAux!AnoExercicio
                  Sql = Sql & " AND CODLANCAMENTO=" & RdoAux!CodLancamento & " AND NUMPARCELA=" & RdoAux!NumParcela & " AND SEQLANCAMENTO=" & RdoAux!numsequencia
                  Sql = Sql & " AND CODCOMPLEMENTO=" & RdoAux!CODCOMPLEMENTO & " AND VALORPAGO>0"
                  Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                  With RdoAux2
                       If .RowCount > 0 Then
                            nValorPago = !ValorPago
                            sDataPagto = "Pago sem Data"
                       Else
                            nValorPago = 0
                            dDataPagto = CDate("01/01/1900")
                            sDataPagto = "Não Pago"
                       End If
                      .Close
                  End With
                  
              End If
             .Close
         End With
           
         If nValorPago > 0 Then
            'BUSCA TAXA
            Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO "
            Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
            Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !numsequencia
            Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    'nValorPago = nValorPago + !VALORTRIBUTO
                    If nValorPago > 0 Then
                        nSomaLancado = nSomaLancado + nSomaValorTributo + !ValorTributo
                    End If
                Else
                    If nValorPago > 0 Then
                        nSomaLancado = nSomaLancado + nSomaValorTributo
                    End If
                End If
            End With
            nSomaPago = nSomaPago + nValorPago
         End If
                            
        
         grdDestino.AddItem !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!numsequencia, "00") & Chr(9) & _
         Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & FormatNumber(nSomaValorTributoJuros, 2) & Chr(9) & _
         FormatNumber(nSomaCorrecao, 2) & Chr(9) & sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
'         nSomaLancado = nSomaLancado + nSomaValorTributo
        .MoveNext
        DoEvents
    Loop
   .Close
End With

'nSomaLancado = 0
'lblValorPago.Caption = Format(nSomaLancado, "#0.00")


'PREENCHE GRID DE ORIGEM
bVenctoNulo = False
Sql = "SELECT * FROM vwCNSREPARCELAMENTOO WHERE NUMPROCESSO='" & sNumProc & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
    
        'SE ALGUMA PARCELA NÃO FOR LOCALIZADA NÃO PERMITE O CANCELAMENTO
        If IsNull(!DataVencimento) Then bVenctoNulo = True
        
        'CARREGA OS TRIBUTOS DE CADA UM DOS LANCAMENTOS
        Set qd.ActiveConnection = cn
        On Error Resume Next
        RdoAux3.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = !CODREDUZIDO
        qd(1) = !CODREDUZIDO 'codigo
        qd(2) = !AnoExercicio
        qd(3) = !AnoExercicio  'ano
        qd(4) = !CodLancamento
        qd(5) = !CodLancamento 'lancamento
        qd(6) = !numsequencia
        qd(7) = !numsequencia 'sequencia
        qd(8) = !NumParcela
        qd(9) = !NumParcela 'numparcela
        qd(10) = !CODCOMPLEMENTO
        qd(11) = !CODCOMPLEMENTO 'complemento
        qd(12) = 1
        qd(13) = 99 'statuslanc
        qd(14) = Format(dDataPag, "mm/dd/yyyy")
        qd(15) = NomeDoUsuario
        Set RdoAux3 = qd.OpenResultset(rdOpenKeyset)
        With RdoAux3
            Do Until .EOF
                'CARREGA MATRIZ DE DÉBITO
                nEval = UBound(aDebito)
                Achou = False
                For x = 1 To nEval
                    If aDebito(x).nCodReduzido = !CODREDUZIDO And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                       aDebito(x).nSeq = !SeqLancamento And _
                       aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
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
                   aDebito(nEval).nLanc = !CodLancamento
                   aDebito(nEval).nSeq = !SeqLancamento
                   aDebito(nEval).nParc = !NumParcela
                   aDebito(nEval).nCompl = !CODCOMPLEMENTO
                   aDebito(nEval).nSituacao = !statuslanc
                   aDebito(nEval).sSituacao = !Situacao
                   aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                   aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                   aDebito(nEval).nValorAtual = !ValorTotal
                Else
                    'SE ENCONTRAR ADICIONAR O VALOR AO JA EXISTENTE
                    If !statuslanc = 3 Or !statuslanc = 4 Or !statuslanc = 6 Then
                        aDebito(x).nValorAtual = aDebito(x).nValorAtual + !ValorTotal
                    End If
                    aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                End If
               .MoveNext
            Loop
           .Close
        End With
       .MoveNext
    Loop
End With
'ADICIONA OS DEBITOS AO GRID DE ORIGEM
nSomaNaoPago = 0
For x = 1 To UBound(aDebito)
    With aDebito(x)
        nSomaNaoPago = nSomaNaoPago + .nValorAtual
        grdOrigem.AddItem .nAno & Chr(9) & Format(.nLanc, "00") & Chr(9) & Format(.nSeq, "00") & Chr(9) & _
        Format(.nParc, "00") & Chr(9) & Format(.nCompl, "00") & Chr(9) & .sVencto & Chr(9) & FormatNumber(.nValorTributo, 2) & Chr(9) & _
        FormatNumber(.nValorAtual, 2) & Chr(9) & "03-NÃO PAGO"
    End With
Next
nSomaPago = CDbl(lblValorPago.Caption)
'lblValorNPago.Caption = FormatNumber(nSomaNaoPago - nSomaPago, 2)

'VERIFICA SE TEM COMPENSAÇÃO
If Val(lblValorCompensar.Caption) > 0 Then
    nTotalACompensar = CDbl(lblValorCompensar.Caption)
    nTotalAtual = 0
    
    nSobra = nTotalACompensar
    With grdOrigem
        For x = 1 To .Rows - 1
             nValorAChecar = CDbl(.TextMatrix(x, 7))
             nTotalAtual = nTotalAtual + nValorAChecar
             If nSobra > nValorAChecar Then
                .TextMatrix(x, 8) = "06-COMPENSADO"
                nSobra = nSobra - nValorAChecar
             ElseIf nSobra > 0 And nSobra < nValorAChecar Then
                 nValorAChecar = CDbl(.TextMatrix(x, 7))
                 nPerc = 1 - (nSobra / nValorAChecar)
                 nValorPerc = FormatNumber(nValorAChecar * nPerc, 2)
                 lblPerc.Caption = FormatNumber(nPerc * 100, 2) & "%"
                 nValorAChecar = CDbl(.TextMatrix(x, 6))
'                 nPerc = 1 - (nSobra / nValorAChecar)
                 nValorPerc = FormatNumber(nValorAChecar * nPerc, 2)
                .TextMatrix(x, 8) = "06-COMPENSADO"
                 'busca o novo codigo do complemento
                 Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXCOMPL FROM DEBITOPARCELA WHERE "
                 Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO=" & .TextMatrix(x, 0) & " AND "
                 Sql = Sql & "CODLANCAMENTO=" & .TextMatrix(x, 1) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 2) & " AND "
                 Sql = Sql & "NUMPARCELA=" & .TextMatrix(x, 3)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 nCodCompl = RdoAux!MAXCOMPL + 1
                 RdoAux.Close
                 'cria uma parcela de compensação
                 nLinhaOriginal = x
                .AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & _
                 Format(nCodCompl, "00") & Chr(9) & .TextMatrix(x, 5) & Chr(9) & "N/A" & Chr(9) & _
                 FormatNumber(nValorPerc, 2) & Chr(9) & "03-NÃO PAGO"
                 lblValorExt.Caption = FormatNumber((nValorAChecar - (nSobra)), 2)
                 nSobra = 0
             Else
                .TextMatrix(x, 8) = "03-NÃO PAGO"
             End If
        Next
        
    End With
Else
    'SE NÃO TEM COMPENSAÇÃO, QUER DIZER QUE NENHUMA PARCELA FOI PAGA
    'NESTE CASO BASTA CANCELAR TODAS AS PARCELAS
    With grdOrigem
        For x = 1 To .Rows - 1
            .TextMatrix(x, 8) = "03-NÃO PAGO"
        Next
    End With
End If

nValorAChecar = 0: nValorNaoPago = 0
For x = 1 To grdOrigem.Rows - 1
    If grdOrigem.TextMatrix(x, 8) = "06-COMPENSADO" Then
        nValorAChecar = nValorAChecar + grdOrigem.TextMatrix(x, 7)
    ElseIf grdOrigem.TextMatrix(x, 8) = "03-NÃO PAGO" And grdOrigem.TextMatrix(x, 7) <> "N/A" Then
        nValorNaoPago = nValorNaoPago + grdOrigem.TextMatrix(x, 7)
    End If
Next
lblNP.Caption = FormatNumber(nValorAChecar, 2)
lblVlNComp.Caption = FormatNumber(nValorNaoPago, 2)
With grdOrigem
    If .TextMatrix(.Rows - 1, 8) = "06-COMPENSADO" Then
        If Val(lblValorNaoPago.Caption) > 0 Then
             .AddItem .TextMatrix(.Rows - 1, 0) & Chr(9) & .TextMatrix(.Rows - 1, 1) & Chr(9) & .TextMatrix(.Rows - 1, 2) & Chr(9) & .TextMatrix(.Rows - 1, 3) & Chr(9) & _
            .TextMatrix(.Rows - 1, 4) & Chr(9) & Format(nCodCompl + 1, "00") & Chr(9) & Format(mskDataParc.Text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
              FormatNumber(CDbl(lblValorNaoPago.Caption), 2) & Chr(9) & "03-NÃO PAGO"
        End If
    Else
        If CDbl(lblValorNaoPago.Caption) > CDbl(lblVlNComp.Caption) Then
            .TextMatrix(.Rows - 1, 7) = FormatNumber(CDbl(lblValorNaoPago.Caption) - CDbl(lblVlNComp.Caption), 2)
''             .TextMatrix(.Rows - 1, 8) = FormatNumber(CDbl(lblValorExt.Caption), 2)
       Else
           If lblValorExt.Caption > 0 Then
                .TextMatrix(.Rows - 1, 7) = FormatNumber(CDbl(lblValorExt.Caption), 2)
           End If
       End If
    End If
End With

With grdOrigem
     If .TextMatrix(.Rows - 1, 6) = "N/A" Then
        .FillStyle = flexFillRepeat
        .Row = .Rows - 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &H9FFFC0
     End If
End With

Exit Sub
Erro:
MsgBox Err.Description
Resume Next

End Sub


