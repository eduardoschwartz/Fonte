VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCancelReparc 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento de Reparcelamento"
   ClientHeight    =   6000
   ClientLeft      =   8895
   ClientTop       =   1995
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9705
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1950
      MaxLength       =   9
      TabIndex        =   19
      Top             =   210
      Width           =   1125
   End
   Begin VB.ComboBox cmbProc 
      Height          =   315
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   540
      Width           =   1725
   End
   Begin MSFlexGridLib.MSFlexGrid grdOrigem 
      Height          =   1485
      Left            =   30
      TabIndex        =   11
      Top             =   3510
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   2619
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmCancelReparc.frx":0000
   End
   Begin MSFlexGridLib.MSFlexGrid grdDestino 
      Height          =   1485
      Left            =   60
      TabIndex        =   6
      Top             =   1260
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   2619
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedCols       =   0
      BackColorBkg    =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código     |^Ano     |^Lanc. |^Seq  |^Parc. |^Compl. |^Vencimento      |>Vl.Lançado     |^Data Pagto.      |>Valor Pago         "
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   315
      Left            =   5760
      TabIndex        =   1
      ToolTipText     =   "Executar o Cancelamento"
      Top             =   5550
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Executar o Cancelamento"
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
      MICON           =   "frmCancelReparc.frx":009A
      PICN            =   "frmCancelReparc.frx":00B6
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
      Left            =   8550
      TabIndex        =   0
      ToolTipText     =   "Sair da Tela"
      Top             =   5550
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
      MICON           =   "frmCancelReparc.frx":0155
      PICN            =   "frmCancelReparc.frx":0171
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   3120
      TabIndex        =   20
      ToolTipText     =   "Consulta Imóvel"
      Top             =   180
      Width           =   435
      _ExtentX        =   767
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCancelReparc.frx":01DF
      PICN            =   "frmCancelReparc.frx":01FB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdFindProc 
      Height          =   315
      Left            =   3120
      TabIndex        =   21
      ToolTipText     =   "Selecionar Nº de Processo"
      Top             =   570
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      BTYPE           =   14
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
      MICON           =   "frmCancelReparc.frx":0355
      PICN            =   "frmCancelReparc.frx":0371
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataParc 
      Height          =   285
      Left            =   8220
      TabIndex        =   22
      Top             =   630
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmCancelReparc.frx":04CB
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
   Begin VB.Label lblDataProc 
      Caption         =   "Label4"
      Height          =   225
      Left            =   5040
      TabIndex        =   29
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Imóvel a ser Consultado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   28
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label lblProp 
      BackStyle       =   0  'Transparent
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
      Left            =   3510
      TabIndex        =   27
      Top             =   240
      Width           =   4365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Processo:"
      Height          =   225
      Index           =   10
      Left            =   3660
      TabIndex        =   26
      Top             =   660
      Width           =   1395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Parcelamento:"
      Height          =   225
      Index           =   9
      Left            =   6480
      TabIndex        =   25
      Top             =   660
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   630
      Width           =   1395
   End
   Begin VB.Label lblCancel 
      BackStyle       =   0  'Transparent
      Caption         =   "CANCELADO"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   810
      Left            =   4050
      TabIndex        =   23
      Top             =   150
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lblValorExt 
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   4905
      TabIndex        =   17
      Top             =   5100
      Width           =   795
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor excedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3375
      TabIndex        =   16
      Top             =   5100
      Width           =   1380
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total compensado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   15
      Top             =   5100
      Width           =   2085
   End
   Begin VB.Label lblNP 
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   2310
      TabIndex        =   14
      Top             =   5100
      Width           =   1020
   End
   Begin VB.Label lblVlNComp 
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   8145
      TabIndex        =   13
      Top             =   5100
      Width           =   1020
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Não Compensado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   6390
      TabIndex        =   12
      Top             =   5100
      Width           =   1680
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total pago no reparcelamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   360
      TabIndex        =   10
      Top             =   2850
      Width           =   3225
   End
   Begin VB.Label lblValorPago 
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
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   3690
      TabIndex        =   9
      Top             =   2850
      Width           =   1425
   End
   Begin VB.Label lblValorNaoPago 
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
      Height          =   225
      Left            =   8070
      TabIndex        =   8
      Top             =   2850
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total da divida Não Paga:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   5205
      TabIndex        =   7
      Top             =   2850
      Width           =   2730
   End
   Begin VB.Label lblFuncCancel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1860
      TabIndex        =   5
      Top             =   5730
      Width           =   2715
   End
   Begin VB.Label lblDataCancel 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1860
      TabIndex        =   4
      Top             =   5490
      Width           =   2715
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelado por..............:"
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   5730
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Cancelamento.:"
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   5490
      Width           =   1755
   End
End
Attribute VB_Name = "frmCancelReparc"
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

Dim RdoAux As rdoResultset, sRet As String
Dim Sql As String, bResize As Boolean, bVenctoNulo As Boolean
Dim aTributo() As TRIBUTO, nLinhaOriginal As Integer

Private Sub cmbProc_Click()
lblDataCancel.Caption = ""
lblFuncCancel.Caption = ""
Sql = "SELECT CANCELADO,DATACANCEL,FUNCIONARIOCANCEL FROM PROCESSOREPARC WHERE NUMPROCESSO='" & cmbProc.Text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
        If .RowCount = 0 Then
            lblCancel.Visible = False
        Else
            If !Cancelado Then
                lblCancel.Visible = True
                lblDataCancel.Caption = Format(!DataCancel, "dd/mm/yyyy")
                lblFuncCancel.Caption = SubNull(!FUNCIONARIOCANCEL)
            Else
                lblCancel.Visible = False
            End If
        End If
       .Close
End With
CarregaProcesso2
End Sub

Private Sub cmdCnsImovel_Click()
sForm = "CR"
frmCnsImovel.show
frmCnsImovel.ZOrder 0
End Sub

Private Sub cmdExec_Click()
Dim x As Integer
Dim RdoAux2 As rdoResultset, evDel As Integer
Dim bAchou As Boolean, nCodReduz As Long, nSeq As Integer
Dim nValorComplemento As Double, nValorTotal As Double
Dim nAno As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer
Dim sUser As String

evDel = 4

If cmbProc.ListIndex = -1 Then Exit Sub
If grdDestino.Rows = 1 Then
    MsgBox "Não foi possível carregar parcelas de destino.", vbCritical, "ERRO"
    Exit Sub
End If
If lblCancel.Visible = True Then
   If Val(grdDestino.TextMatrix(1, 0)) < 500000 Then
        If MsgBox("Este Reparcelamento já foi cancelado." & vbCrLf & "Deseja continuar ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
           'sUser = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8)
           sUser = NomeDeLogin
           If sUser = "SCHWARTZ" Or sUser = "GLEISE" Or sUser = "ROSE" Then
           Else
              MsgBox "Permissão negada para este usuário.", vbExclamation, "Atenção"
              Exit Sub
           End If
        Else
           Exit Sub
        End If
   End If
End If

bAchou = False
For x = 1 To grdDestino.Rows - 1
    If grdDestino.TextMatrix(x, 8) = "Não Pago" Then
        bAchou = True
        Exit For
    End If
Next

If Not bAchou Then
   MsgBox "Este Reparcelamento já foi totalmente pago.", vbExclamation, "Atenção"
   Exit Sub
End If

If Not IsDate(mskDataParc.Text) Then
    MsgBox "Data de parcelamento invalida.", vbExclamation, "atenção"
    Exit Sub
End If


If MsgBox("Os débitos do reparcelamento serão cancelados." & vbCrLf & vbCrLf & "Deseja continuar ?", vbQuestion + vbYesNo, "CONFIRMAÇÂO DE CANCELAMENTO !!!") = vbNo Then Exit Sub
Ocupado
'CANCELAMENTO DAS PARCELAS DE DESTINO
With grdDestino
    For x = 1 To .Rows - 1
        If Not IsDate(.TextMatrix(x, 8)) Then
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'ATUALIZAÇÃO DAS PARCELAS DE ORIGEM

With grdOrigem
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 7) <> "N/A" Then
            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & Val(Left$(.TextMatrix(x, 9), 2)) & " WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND NUMPARCELA=" & .TextMatrix(x, 4) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            cn.Execute Sql, rdExecDirect
        Else
            'CARREGA ORIGINAL PARCELA COMPLEMENTO
            nLinhaOriginal = .Rows - 2
            Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(nLinhaOriginal, 1) & " AND CODLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 2) & " AND "
            'Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND CODLANCAMENTO=" & .TextMatrix(x - 1, 2) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 3) & " AND NUMPARCELA=" & .TextMatrix(nLinhaOriginal, 4) & " AND "
            'Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(x, 5)) - 1
            Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(nLinhaOriginal, 5))
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                'GRAVA COMPLEMENTO PARCELA
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
'                Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,"
'               Sql = Sql & "DATAAJUIZA,VALORJUROS,NUMPROCESSO,USUARIO) VALUES(" & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 0)) & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 1)) & "," & !CodLancamento & ","
'               Sql = Sql & !SeqLancamento & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 4)) & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 5)) & "," & Val(Left$(grdOrigem.TextMatrix(x, 9), 2)) & ",'" & Format(grdOrigem.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',"
               Sql = Sql & Val(SubNull(!CODMOEDA)) & "," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & "," & Val(SubNull(!numcertidao)) & "," & IIf(IsNull(!datainscricao), "Null", "'" & Format(!datainscricao, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!dataajuiza), "Null", "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'") & "," & !ValorJuros & ",'"
 ''              Sql = Sql & cmbProc.Text & "','" & Left$(NomeDeLogin, 25) & "')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,"
               Sql = Sql & "DATAAJUIZA,VALORJUROS,NUMPROCESSO,USERID) VALUES(" & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 0)) & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 1)) & "," & !CodLancamento & ","
               Sql = Sql & !SeqLancamento & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 4)) & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 5)) & "," & Val(Left$(grdOrigem.TextMatrix(x, 9), 2)) & ",'" & Format(grdOrigem.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',"
               Sql = Sql & Val(SubNull(!CODMOEDA)) & "," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & "," & Val(SubNull(!numcertidao)) & "," & IIf(IsNull(!datainscricao), "Null", "'" & Format(!datainscricao, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!dataajuiza), "Null", "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'") & "," & !ValorJuros & ",'"
               Sql = Sql & cmbProc.Text & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
                cn.Execute Sql, rdExecDirect
            End With
            'CARREGA ORIGINAL TRIBUTO COMPLEMENTO
            Sql = "SELECT sum(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & .TextMatrix(x, 0) & " AND "
            Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(nLinhaOriginal, 1) & " AND CODLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 2) & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(nLinhaOriginal, 3) & " AND NUMPARCELA=" & .TextMatrix(nLinhaOriginal, 4) & " AND "
            Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(nLinhaOriginal, 5)) & " AND CODTRIBUTO <>3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If Not IsNull(!soma) Then
                   nValorTotal = !soma
                Else
                    nValorTotal = 0
                End If
              .Close
           End With
           nValorComplemento = CDbl(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 8))
           'nValorComplemento = CDbl(lblValorNaoPago.Caption)
           ReDim aTributo(0)
           Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & .TextMatrix(grdOrigem.Rows - 1, 0) & " AND "
           Sql = Sql & "ANOEXERCICIO=" & .TextMatrix(grdOrigem.Rows - 1, 1) & " AND CODLANCAMENTO=" & .TextMatrix(grdOrigem.Rows - 1, 2) & " AND "
           Sql = Sql & "SEQLANCAMENTO=" & .TextMatrix(grdOrigem.Rows - 1, 3) & " AND NUMPARCELA=" & .TextMatrix(grdOrigem.Rows - 1, 4) & " AND "
           Sql = Sql & "CODCOMPLEMENTO=" & Val(.TextMatrix(grdOrigem.Rows - 2, 5)) & " AND CODTRIBUTO <>3"
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
               Sql = Sql & nSeq & "," & nParc & "," & Val(grdOrigem.TextMatrix(grdOrigem.Rows - 1, 5)) & "," & aTributo(TY).nCodTributo & "," & Virg2Ponto(CStr(aTributo(TY).nValorTributo)) & " )"
               cn.Execute Sql, rdExecDirect
           Next
        End If
    Next
End With

'CANCELAMENTO DO PROCESSO
If Right$(cmbProc.Text, 4) <> "SMAR" Then
    'Sql = "UPDATE PROCESSOREPARC SET CANCELADO=1,DATACANCEL='" & Format(Now, "mm/dd/yyyy") & "',FUNCIONARIOCANCEL='" & Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) & "' WHERE NUMPROCESSO='" & cmbProc.text & "'"
    Sql = "UPDATE PROCESSOREPARC SET CANCELADO=1,DATACANCEL='" & Format(Now, "mm/dd/yyyy") & "',FUNCIONARIOCANCEL='" & NomeDeLogin & "' WHERE NUMPROCESSO='" & cmbProc.Text & "'"
    cn.Execute Sql, rdExecDirect
Else
    For x = 1 To grdDestino.Rows - 1
        nCodReduz = grdDestino.TextMatrix(x, 0)
        nSeq = grdDestino.TextMatrix(x, 3)
        Sql = "UPDATE REPARCTMP SET CODSIT=1 WHERE CODREDUZD=" & nCodReduz & " AND CODSEQD=" & nSeq
        cn.Execute Sql, rdExecDirect
    Next
    
    'Sql = "UPDATE REPARC2TMP SET DATACANCEL='" & Format(Now, "mm/dd/yyyy") & "', FUNCIONARIOCANCEL='" & Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) & "'"
    Sql = "UPDATE REPARC2TMP SET DATACANCEL='" & Format(Now, "mm/dd/yyyy") & "', FUNCIONARIOCANCEL='" & NomeDeLogin & "'"
    Sql = Sql & " WHERE CODREDUZ=" & Val(txtCod.Text) & " AND NUMSEQ=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    cn.Execute Sql, rdExecDirect
End If

Liberado
MsgBox "O cancelamento do reparcelamento foi executado com sucesso.", vbExclamation, "Atenção"
grdDestino.Rows = 1
grdOrigem.Rows = 1
cmbProc.Clear
txtCod.SetFocus

End Sub

Private Sub cmdFindProc_Click()
frmCnsNumProc.show
frmCnsNumProc.ZOrder 0
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
     txtCod.Text = Left$(CodImovel, 7)
     CodImovel = 0
     txtCod_LostFocus

End If
bResize = True
End Sub

Private Sub Form_Load()
Ocupado
sRet = RetEventUserForm(Me.Name)
Centraliza Me
Liberado

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CodImovel = 0
End Sub

Private Sub txtCod_GotFocus()

txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    txtCod_LostFocus
    Exit Sub
End If

Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Public Sub txtCod_LostFocus()
Dim nCodImovel As Long
On Error Resume Next
If Val(txtCod.Text) = 0 Then Exit Sub
txtCod.Text = Format(txtCod.Text, "0000000")
nCodImovel = Val(txtCod.Text)
lblDataCancel.Caption = ""
lblFuncCancel.Caption = ""
cmbProc.SetFocus
CarregaImovel nCodImovel
CarregaProc


End Sub

Private Sub CarregaImovel(nCodigoImovel As Long)

Ocupado

Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & nCodigoImovel
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
With RdoAux
    If RdoAux.RowCount > 0 Then
         lblProp.Caption = !nomecidadao
    Else
       Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodigoImovel
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
           If .RowCount > 0 Then
               lblProp.Caption = !RazaoSocial
           End If
       End With
   
    End If
    .Close
End With

Liberado

End Sub

Private Sub CarregaProc()
    
lblCancel.Visible = False
grdOrigem.Rows = 1: grdDestino.Rows = 1
lblVlNComp.Caption = "0,00"
lblValorPago.Caption = "0,00"
lblValorNaoPago.Caption = "0,00"
lblValorExt.Caption = "0,00"
lblNP.Caption = "0,00"
cmbProc.Clear
Sql = "SELECT DISTINCT NUMPROCESSO FROM ORIGEMREPARC WHERE "
Sql = Sql & "CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
      Do Until .EOF
            cmbProc.AddItem !numprocesso
           .MoveNext
      Loop
     .Close
End With
'VERIFICA OS REPARCELAMENTOS DA SMAR
'Sql = "SELECT * From REPARC2TMP Where CODREDUZ =" & Val(txtCod.text)
Sql = "SELECT DISTINCT(CODSEQD) From REPARCTMP Where CODREDUZD =" & Val(txtCod.Text) & " Or CODREDUZO = " & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        lblDataProc.Caption = ""
        lblDataParc.Caption = ""
        LimpaMascara mskDataParc
        grdOrigem.Rows = 1
        grdDestino.Rows = 1
        lblValorExt.Caption = FormatNumber(0, 2)
        lblValorPago.Caption = FormatNumber(0, 2)
        
        Exit Sub
    Else
        Do Until .EOF
           cmbProc.AddItem CStr(!CODSEQD) & "/SMAR"
          .MoveNext
        Loop
        cmbProc.ListIndex = 0
    End If
   .Close
End With

    
If cmbProc.ListCount = 0 Then
    MsgBox "Não existem processos de reparcelamento para este imóvel.", vbExclamation, "Atenção"
End If
    
End Sub

'Private Sub CarregaProcesso()
'
'Dim RdoAux2 As rdoResultset
'Dim nValorLanc As Double
'Dim nValorJuros As Double
'Dim nValorMulta As Double
'Dim nValorCorrecao As Double
'Dim nValorAtual As Double
'Dim dDataBase As Date
'Dim dDataVencto As Date
'Dim dDataPag As Date
'Dim nValorPago As Double
'Dim nSomaValorTributo As Double
'Dim nSomaPago As Double
'Dim nTotalACompensar As Double
'Dim nTotalAtual As Double
'Dim nValorAChecar As Double
'Dim nSobra As Double
'Dim nCodCompl As Integer
'Dim x As Integer
'Dim dDataPagto As Date, sDataPagto As String
'
'
'grdOrigem.Rows = 1
'grdDestino.Rows = 1
'Ocupado
''PREENCHE GRID DE DESTINO
'If Right$(cmbProc.text, 4) <> "SMAR" Then
'    Sql = "SELECT numprocesso, datareparc From processoreparc WHERE numprocesso = '" & cmbProc.text & "'"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    mskDataParc.text = Format(RdoAux!DATAREPARC, "dd/mm/yyyy")
'    RdoAux.Close
'
'    Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & cmbProc.text & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
'   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'
'        nValorPago = 0
'        nSomaPago = 0
'        Do Until .EOF
'             lblDataProc.Caption = Format(!DATAPROCESSO, "dd/mm/yyyy")
'             dDataVencto = Format(!DATADEBASE, "dd/mm/yyyy")
'             dDataPag = Format(!DATAREPARC, "dd/mm/yyyy")
'             'BUSCA VALOR LANÇADO
'             Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
'             Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
'             Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
'             Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
'             Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !NUMSEQUENCIA
'             Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
'             Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
'             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'             With RdoAux2
'                nValorLanc = !VALORTRIBUTO
'                If (dDataPag > dDataVencto) Then
'                    nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
'                    nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
'                    nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
'                Else
'                    nValorCorrecao = 0
'                    nValorJuros = 0
'                    nValorMulta = 0
'                End If
'                nSomaValorTributo = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
'                .Close
'             End With
'
'             'BUSCA VALORPAGO
'             Sql = "SELECT VALORPAGO,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
'             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !NUMSEQUENCIA
'             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND SEQPAG=0"
'             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'             With RdoAux2
'                  If .RowCount > 0 Then
'                      nValorPago = !ValorPago
'                      dDataPagto = !DataPagamento
'                      sDataPagto = Format(!DataPagamento, "dd/mm/yyyy")
'                  Else
'                      Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
'                      Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
'                      Sql = Sql & "WHERE CODREDUZIDO=" & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO = " & RdoAux!AnoExercicio
'                      Sql = Sql & " AND CODLANCAMENTO=" & RdoAux!CodLancamento & " AND NUMPARCELA=" & RdoAux!NumParcela & " AND SEQLANCAMENTO=" & RdoAux!NUMSEQUENCIA
'                      Sql = Sql & " AND CODCOMPLEMENTO=" & RdoAux!CODCOMPLEMENTO & " AND VALORPAGO>0"
'                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                      With RdoAux2
'                           If .RowCount > 0 Then
'                                nValorPago = !ValorPago
'                                sDataPagto = "Pago sem Data"
'                           Else
'                                nValorPago = 0
'                                dDataPagto = CDate("01/01/1900")
'                                sDataPagto = "Não Pago"
'                           End If
'                          .Close
'                      End With
'
'                  End If
'                 .Close
'             End With
'
'             If nValorPago > 0 Then
'                'BUSCA TAXA
'                Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO "
'                Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
'                Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !NUMSEQUENCIA
'                Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
'                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                With RdoAux2
'                    If .RowCount > 0 Then
'                        nValorPago = nValorPago + !VALORTRIBUTO
'                    End If
'                End With
'             End If
'
'             nSomaPago = nSomaPago + nValorPago
'             grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!NUMSEQUENCIA, "00") & Chr(9) & _
'             Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!datavencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & _
'             sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
'            .MoveNext
'        Loop
'    End With
'Else
'
'    Sql = "SELECT CODREDUZD FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.text) & " AND CODSEQD=" & Left$(cmbProc.text, Len(cmbProc) - 5)
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    nCodReduz = RdoAux!CODREDUZD
'    RdoAux.Close
'
'    Sql = "SELECT * FROM REPARC2TMP WHERE CODREDUZ=" & nCodReduz & " AND CODSEQ=" & Left$(cmbProc.text, Len(cmbProc) - 5)
'    Sql = Sql & " ORDER BY ANOEXERC,CODLANC,CODSEQ,PARCELAS"
'
'
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If .RowCount = 0 Then Exit Sub
'        mskDataParc.text = Format(!DataVencto, "dd/mm/yyyy")
'        dDataPag = CDate(mskDataParc.text)
'       .Close
'    End With
'
'    Sql = "SELECT DISTINCT * FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.text) & " AND CODSEQD=" & Left$(cmbProc.text, Len(cmbProc) - 5)
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        If .RowCount > 0 Then
'            If !CODSIT > 0 Then
'                lblCancel.Visible = True
'                Sql = "SELECT CODREDUZ,DATACANCEL,FUNCIONARIOCANCEL FROM REPARC2TMP WHERE CODREDUZ=" & Val(txtCod.text) & " AND NUMSEQ=" & Left$(cmbProc.text, Len(cmbProc) - 5)
'                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                If IsNull(RdoAux2!DATACANCEL) Then
'                    lblDataCancel.Caption = "N/A"
'                    lblFuncCancel.Caption = "N/A"
'                Else
'                    lblDataCancel.Caption = Format(RdoAux2!DATACANCEL, "dd/mm/yyyy")
'                    lblFuncCancel.Caption = SubNull(RdoAux2!FUNCIONARIOCANCEL)
'                End If
'                RdoAux2.Close
'            Else
'                lblCancel.Visible = False
'            End If
'        Else
'            lblCancel.Visible = False
'        End If
'        Do Until .EOF
'            Sql = "SELECT CODREDUZIDO,DATAVENCIMENTO,DATADEBASE FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZO & " AND ANOEXERCICIO=" & !ANOEXERCO & " AND CODLANCAMENTO=" & !CODLANCO & " AND SEQLANCAMENTO=" & !CODSEQO & " AND NUMPARCELA=" & !NUMPARCO & " AND CODCOMPLEMENTO=" & !CODCOMPLO
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            With RdoAux2
'                If .RowCount > 0 Then
'                   dDataBase = Format(!DATADEBASE, "dd/mm/yyyy")
'                   dDataVencto = Format(!datavencimento, "dd/mm/yyyy")
'                Else
'                    GoTo prox
'                End If
'               .Close
'            End With
'            'BUSCA VALOR LANCADO
'             Sql = "SELECT SUM(TOTALLANCADO) AS TOTAL FROM VWCNSLANCAMENTO WHERE CODREDUZIDO=" & !CODREDUZO & " AND ANOEXERCICIO = " & !ANOEXERCO
'             Sql = Sql & " AND CODLANCAMENTO=" & !CODLANCO & " AND NUMPARCELA=" & !NUMPARCO & " AND SEQLANCAMENTO=" & !CODSEQO
'             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLO & " AND CODTRIBUTO<>3"
'             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'             With RdoAux2
'                  If Not IsNull(!TOTAL) Then
'                      nValorLanc = !TOTAL
'                  Else
'                      nValorLanc = 0
'                  End If
'                 .Close
'             End With
'
'            If (dDataPag > dDataVencto) Then
'                If IsDate(lblDataCancel.Caption) Then
'                    nValorCorrecao = FormatNumber(CalculaCorrecao3(nValorLanc, dDataVencto), 2)
'                    nValorJuros = FormatNumber(CalculaJuros3(nValorLanc + nValorCorrecao, dDataVencto), 2)
'                    nValorMulta = FormatNumber(CalculaMulta3(nValorLanc + nValorCorrecao, dDataVencto), 2)
'                Else
'                    nValorCorrecao = FormatNumber(CalculaCorrecao(nValorLanc, dDataVencto), 2)
'                    nValorJuros = FormatNumber(CalculaJuros(nValorLanc + nValorCorrecao, dDataVencto), 2)
'                    nValorMulta = FormatNumber(CalculaMulta(nValorLanc + nValorCorrecao, dDataVencto), 2)
'                End If
'            Else
'                nValorCorrecao = 0
'                nValorJuros = 0
'                nValorMulta = 0
'            End If
'            nValorAtual = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
'
'
'             grdOrigem.AddItem Format(!CODREDUZO, "0000000") & Chr(9) & !ANOEXERCO & Chr(9) & Format(!CODLANCO, "00") & Chr(9) & Format(!CODSEQO, "00") & Chr(9) & _
'             Format(!NUMPARCO, "00") & Chr(9) & Format(!CODCOMPLO, "00") & Chr(9) & dDataVencto & Chr(9) & FormatNumber(nValorLanc, 2) & Chr(9) & _
'             FormatNumber(nValorAtual, 2)
'prox:
'           .MoveNext
'        Loop
'       .Close
'    End With
'    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & Left$(cmbProc.text, Len(cmbProc) - 5) & " AND STATUSLANC<>5"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        Do Until .EOF
'
'             'BUSCA VALOR LANÇADO
'             Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO FROM DEBITOTRIBUTO "
'             Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
'             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
'             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
'             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'             With RdoAux2
'                 nSomaValorTributo = !VALORTRIBUTO
'                .Close
'             End With
'
'             'BUSCA VALORPAGO
'             Sql = "SELECT VALORPAGO,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
'             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
'             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND SEQPAG=0"
'             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'             With RdoAux2
'                  If .RowCount > 0 Then
'                      nValorPago = !ValorPago
'                      dDataPagto = !DataPagamento
'                  Else
'
'                      nValorPago = 0
'                      dDataPagto = CDate("01/01/1900")
'                  End If
'                 .Close
'             End With
'
'             nSomaPago = nSomaPago + nValorPago
'
'           If dDataPagto = "01/01/1900" Then
'                Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
'                Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
'                Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
'                Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
'                Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND VALORPAGO>0"
'                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                With RdoAux2
'                    If .RowCount > 0 Then
'                        nValorPago = FormatNumber(!ValorPago, 2)
'                        sDataPagto = "Pago sem Data"
'                    Else
'                        sDataPagto = "Não Pago"
'                    End If
'                   .Close
'                End With
'
'           Else
'                If nValorPago = 0 Then
'                   sDataPagto = "Pago sem Data"
'                Else
'                    sDataPagto = Format(dDataPagto, "dd/mm/yyyy")
'                End If
'           End If
'
'           grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & _
'           Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!datavencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & _
'           sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
'          .MoveNext
'        Loop
'       .Close
'    End With
'End If
'lblValorPago.Caption = FormatNumber(nSomaPago, 2)
'
'If Right$(cmbProc.text, 4) <> "SMAR" Then
'    'PREENCHE GRID DE ORIGEM
'    bVenctoNulo = False
'    Sql = "SELECT * FROM vwCNSREPARCELAMENTOO WHERE NUMPROCESSO='" & cmbProc.text & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        Do Until .EOF
'
'            'SE ALGUMA PARCELA NÃO FOR LOCALIZADA NÃO PERMITE O CANCELAMENTO
'            If IsNull(!datavencimento) Then bVenctoNulo = True
'
'
'            'BUSCA VALOR LANCADO
'             nValorLanc = 0
'             Sql = "SELECT sum(TOTALLANCADO) AS TOTAL FROM VWCNSLANCAMENTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
'             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !NUMSEQUENCIA
'             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
'             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'             With RdoAux2
'                  If .RowCount > 0 Then
'                     If Not IsNull(!TOTAL) Then
'                        nValorLanc = !TOTAL
'                     Else
'                        nValorLanc = 0
'                     End If
'                  Else
'                      Exit Do
'                  End If
'                 .Close
'             End With
'             If IsNull(!DATADEBASE) Then GoTo PROXIMO
'             dDataBase = !DATADEBASE
'             dDataVencto = !datavencimento
'             If mskDataParc.ClipText <> "" Then
'               dDataPag = CDate(mskDataParc.text)
'             End If
'             If (dDataPag > dDataVencto) Then
'                nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
'                nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
'                nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
'            Else
'                nValorCorrecao = 0
'                nValorJuros = 0
'                nValorMulta = 0
'            End If
'            nValorAtual = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
'            grdOrigem.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!NUMSEQUENCIA, "00") & Chr(9) & _
'            Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!datavencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nValorLanc, 2) & Chr(9) & _
'            FormatNumber(nValorAtual, 2)
'PROXIMO:
'            .MoveNext
'        Loop
'    End With
'End If
'
''VERIFICA SE TEM COMPENSAÇÃO
'If Val(lblValorPago.Caption) > 0 Then '
'    nTotalACompensar = CDbl(lblValorPago.Caption)
'    nTotalAtual = 0
'
'    nSobra = nTotalACompensar
'    With grdOrigem
'        For x = 1 To .Rows - 1
'             nValorAChecar = CDbl(.TextMatrix(x, 8))
'             nTotalAtual = nTotalAtual + nValorAChecar
'             If nSobra > nValorAChecar Then
'                .TextMatrix(x, 9) = "06-COMPENSADO"
'                nSobra = nSobra - nValorAChecar
'             ElseIf nSobra > 0 And nSobra < nValorAChecar Then
'                .TextMatrix(x, 9) = "06-COMPENSADO"
'                 'busca o novo codigo do complemento
'                 Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXCOMPL FROM DEBITOPARCELA WHERE "
'                 Sql = Sql & "CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND "
'                 Sql = Sql & "CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND "
'                 Sql = Sql & "NUMPARCELA=" & .TextMatrix(x, 4)
'                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                 nCodCompl = RdoAux!MAXCOMPL + 1
'                 RdoAux.Close
'                 'cria uma parcela de compensação
'                .AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & _
'                .TextMatrix(x, 4) & Chr(9) & Format(nCodCompl, "00") & Chr(9) & Format(mskDataParc.text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
'                 FormatNumber(nValorAChecar - (nSobra), 2) & Chr(9) & "03-NÃO PAGO"
'                 nSobra = 0
'             Else
'                .TextMatrix(x, 9) = "03-NÃO PAGO"
'             End If
'        Next
'        If nTotalACompensar > nTotalAtual Then
'            lblValorExt.Caption = FormatNumber(nTotalACompensar - nTotalAtual, 2)
'        End If
'    End With
'Else
'    'SE NÃO TEM COMPENSAÇÃO, QUER DIZER QUE NENHUMA PARCELA FOI PAGA
'    'NESTE CASO BASTA CANCELAR TODAS AS PARCELAS
'    With grdOrigem
'        For x = 1 To .Rows - 1
'            .TextMatrix(x, 9) = "03-NÃO PAGO"
'        Next
'    End With
'End If
'
'nValorAChecar = 0
'For x = 1 To grdOrigem.Rows - 2
'    If grdOrigem.TextMatrix(x, 9) = "06-COMPENSADO" Then
'        nValorAChecar = nValorAChecar + grdOrigem.TextMatrix(x, 8)
'    End If
'Next
'lblNP.Caption = FormatNumber(nValorAChecar, 2)
'
'
'With grdOrigem
'     If .TextMatrix(.Rows - 1, 7) = "N/A" Then
'        .FillStyle = flexFillRepeat
'        .Row = .Rows - 1
'        .Col = 0
'        .RowSel = .Rows - 1
'        .ColSel = .Cols - 1
'        .CellBackColor = &H9FFFC0
'     End If
'End With
'
'Liberado
'End Sub
'

Private Function CalculaJuros2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String

If Year(dDataPagto) > Year(Now) Then
    CalculaJuros2 = 0
    Exit Function
End If

If dDataVencto >= dDataPagto Then
    CalculaJuros2 = 0
    Exit Function
End If
nNumMes = Int((DateDiff("d", dDataVencto, dDataPagto)) / 30)
Sql = "SELECT PERCJUROS FROM JUROS WHERE ANOJUROS=" & Year(dDataPagto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
        CalculaJuros2 = 0
        Exit Function
    Else
        nValorPerc = !PERCJUROS
    End If
   .Close
End With
nValorPerc = nValorPerc / 100

CalculaJuros2 = nValorDebito * nValorPerc * nNumMes
If CalculaJuros2 > 0 Then
   CalculaJuros2 = FormatNumber(CalculaJuros2, 3)
End If

End Function

Private Function CalculaMulta2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String


If dDataVencto >= dDataPagto Then
    CalculaMulta2 = 0
    Exit Function
End If

nNumDia = Abs(DateDiff("d", dDataPagto, dDataVencto))

If nNumDia = 0 Then
   CalculaMulta2 = 0
   Exit Function
End If

Sql = "SELECT MINDIA,MAXDIA,PERCDIA FROM MULTA WHERE ANOMULTA=" & Year(dDataVencto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         If nNumDia >= !MINDIA And nNumDia <= !MAXDIA Then
             nValorPerc = !PERCDIA
             Exit Do
         ElseIf nNumDia >= !MINDIA And !MAXDIA = 0 Then
             nValorPerc = !PERCDIA
             Exit Do
         End If
        .MoveNext
    Loop
End With

nValorPerc = nValorPerc / 100
CalculaMulta2 = nValorDebito * nValorPerc
If CalculaMulta2 > 0 Then
   CalculaMulta2 = FormatNumber(CalculaMulta2, 3)
End If

End Function

Private Function CalculaCorrecao2(nValorDebito As Double, dDataBase As Date, dDataVencto As Date) As Double

Dim UfirAtual As Double
Dim UfirBase As Double

If Year(dDataVencto) > Year(mskDataParc.Text) Then
   CalculaCorrecao2 = 0
   Exit Function
End If
UfirAtual = RetornaUFIR(Year(dDataVencto))
UfirBase = RetornaUFIR(Year(dDataBase))

CalculaCorrecao2 = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecao2 > 0 Then
   CalculaCorrecao2 = FormatNumber(CalculaCorrecao2, 2)
End If
End Function

Public Function CalculaJuros3(nValorDebito As Double, dDataVencto As Date, Optional dDataNow As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double

If Not IsDate(lblDataCancel.Caption) Then
    dDataNow = Now
Else
    dDataNow = CDate(lblDataCancel.Caption)
End If

If dDataVencto >= dDataNow Then
    CalculaJuros3 = 0
    Exit Function
End If
nNumMes = Int((DateDiff("d", dDataVencto, dDataNow)) / 30)

If Not dcJuros.Exists(Year(dDataNow)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJuros3 = 0
   Exit Function
End If
nValorPerc = dcJuros.Item(Year(dDataNow))

nValorPerc = nValorPerc / 100

CalculaJuros3 = nValorDebito * nValorPerc * nNumMes
If CalculaJuros3 > 0 Then
   CalculaJuros3 = FormatNumber(CalculaJuros3, 3)
End If

End Function

Public Function CalculaMulta3(nValorDebito As Double, dDataVencto As Date, Optional dDataNow As Date) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double

If Not IsDate(lblDataCancel.Caption) Then
    dDataNow = Now
Else
    dDataNow = CDate(lblDataCancel.Caption)
End If

If dDataVencto >= dDataNow Then
    CalculaMulta3 = 0
    Exit Function
End If
On Error Resume Next
nNumDia = Abs(DateDiff("d", dDataNow, dDataVencto))

If nNumDia = 0 Then
   CalculaMulta3 = 0
   Exit Function
End If

For x = 1 To UBound(aMulta)
    If aMulta(x).nAno = Year(dDataVencto) Then
        If nNumDia >= aMulta(x).nMin And nNumDia <= aMulta(x).nMax Then
            nValorPerc = aMulta(x).nValor
            Exit For
        ElseIf nNumDia >= aMulta(x).nMin And aMulta(x).nMax = 0 Then
            nValorPerc = aMulta(x).nValor
            Exit For
        End If
    End If
Next

nValorPerc = nValorPerc / 100
CalculaMulta3 = nValorDebito * nValorPerc
If CalculaMulta3 > 0 Then
   CalculaMulta3 = FormatNumber(CalculaMulta3, 3)
End If

End Function

Public Function CalculaCorrecao3(nValorDebito As Double, dDataBase As Date) As Double

Dim dDataNow As Date
Dim UfirAtual As Double
Dim UfirBase As Double

If Not IsDate(lblDataCancel.Caption) Then
    dDataNow = Now
Else
    dDataNow = CDate(lblDataCancel.Caption)
End If

If Year(dDataBase) > Year(dDataNow) Then
    CalculaCorrecao3 = 0
    Exit Function
End If

UfirAtual = RetornaUFIR(Year(dDataNow))
If UfirAtual = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano atual.", vbCritical, "Alerta !!!"
    CalculaCorrecao3 = 0
    Exit Function
End If

UfirBase = RetornaUFIR(Year(dDataBase))
If UfirBase = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano base.", vbCritical, "Alerta !!!"
    CalculaCorrecao3 = 0
    Exit Function
End If

CalculaCorrecao3 = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecao3 > 0 Then
   CalculaCorrecao3 = FormatNumber(CalculaCorrecao3, 2)
End If
End Function


Private Sub CarregaProcesso2()

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date
Dim dDataPag As Date
Dim nValorPago As Double, nValorNaoPago As Double
Dim nSomaValorTributo As Double
Dim nSomaPago As Double, nSomaNaoPago As Double
Dim nTotalACompensar As Double
Dim nTotalAtual As Double
Dim nValorAChecar As Double
Dim nSobra As Double
Dim nCodCompl As Integer
Dim x As Integer
Dim dDataPagto As Date, sDataPagto As String
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, Achou As Boolean

ReDim aDebito(0)

grdOrigem.Rows = 1
grdDestino.Rows = 1
Ocupado
'PREENCHE GRID DE DESTINO
If Right$(cmbProc.Text, 4) <> "SMAR" Then
    Sql = "SELECT numprocesso, datareparc,novo From processoreparc WHERE numprocesso = '" & cmbProc.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    mskDataParc.Text = Format(RdoAux!datareparc, "dd/mm/yyyy")
    If RdoAux!Novo = True Then
        Liberado
        MsgBox "Este processo pertence a nova Lei de Parcelamento " & vbCrLf & "e não pode ser cancelado por esta tela.", vbExclamation, "Atenção"
        Exit Sub
    End If
    RdoAux.Close
    
    Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & cmbProc.Text & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        
        nValorPago = 0
        nSomaPago = 0: nSomaNaoPago = 0
        Do Until .EOF
             lblDataProc.Caption = Format(!DATAPROCESSO, "dd/mm/yyyy")
             dDataVencto = Format(!DATADEBASE, "dd/mm/yyyy")
          '   dDataPag = Format(!DATAREPARC, "dd/mm/yyyy")
          dDataPag = CDate(mskDataParc.Text)
             'BUSCA VALOR LANÇADO
             Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
             Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
             Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
             Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
             Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3 AND CODTRIBUTO<>90 "
             Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                nValorLanc = !ValorTributo
                If (dDataPag > dDataVencto) Then
                    nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
                    nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                    nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                Else
                    nValorCorrecao = 0
                    nValorJuros = 0
                    nValorMulta = 0
                End If
                nSomaValorTributo = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
                .Close
             End With
                
             'BUSCA VALORPAGO
             Sql = "SELECT VALORPAGO,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !numsequencia
             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND SEQPAG=0"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                  If .RowCount > 0 Then
                      nValorPago = !ValorPago
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
                        nValorPago = nValorPago + !ValorTributo
                    End If
                End With
 '            Else
  '              nValorNaoPago = nValorNaoPago + !VALORTRIBUTO
             End If
                                
             nSomaPago = nSomaPago + nValorPago
             grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!numsequencia, "00") & Chr(9) & _
             Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & _
             sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
            .MoveNext
        Loop
    End With
Else
    'BUSCA O CONTRIBUINTE RESPONSAVEL NA SMAR
    Sql = "SELECT CODREDUZD FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.Text) & " AND CODSEQD=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nCodReduz = RdoAux!CODREDUZD
    RdoAux.Close
    
    'CARREGA A DATA DO REPARCELAMENTO
    Sql = "SELECT * FROM REPARC2TMP WHERE CODREDUZ=" & nCodReduz & " AND CODSEQ=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Sql = Sql & " ORDER BY ANOEXERC,CODLANC,CODSEQ,PARCELAS"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            Liberado
            MsgBox "Parcelamento nào localizado em REPARC2TMP"
            Exit Sub
        End If
        mskDataParc.Text = Format(!DataVencto, "dd/mm/yyyy")
        dDataPag = CDate(mskDataParc.Text)
       .Close
    End With
    
    'CARREGA TODOS OS LANCAMENTOS DO REPARCELAMENTO
    Sql = "SELECT DISTINCT * FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.Text) & " AND CODSEQD=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            'VERIFICA SE FOI OU NÃO CANCELADO E POR QUEM
            If !CODSIT > 0 Then
                lblCancel.Visible = True
                Sql = "SELECT CODREDUZ,DATACANCEL,FUNCIONARIOCANCEL FROM REPARC2TMP WHERE CODREDUZ=" & Val(txtCod.Text) & " AND NUMSEQ=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If IsNull(RdoAux2!DataCancel) Then
                    lblDataCancel.Caption = "N/A"
                    lblFuncCancel.Caption = "N/A"
                Else
                    If IsDate(RdoAux2!DataCancel) Then
                        lblDataCancel.Caption = Format(RdoAux2!DataCancel, "dd/mm/yyyy")
                    Else
                        lblDataCancel.Caption = ""
                    End If
                    lblFuncCancel.Caption = SubNull(RdoAux2!FUNCIONARIOCANCEL)
                End If
                RdoAux2.Close
            Else
                lblCancel.Visible = False
            End If
        Else
            lblCancel.Visible = False
        End If
        
        Do Until .EOF
        
            'CARREGA OS TRIBUTOS DE CADA UM DOS LANCAMENTOS
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux3.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = !CODREDUZO
            qd(1) = !CODREDUZO 'codigo
            qd(2) = !ANOEXERCO
            qd(3) = !ANOEXERCO  'ano
            qd(4) = !CODLANCO
            qd(5) = !CODLANCO 'lancamento
            qd(6) = !CODSEQO
            qd(7) = !CODSEQO 'sequencia
            qd(8) = !NUMPARCO
            qd(9) = !NUMPARCO 'numparcela
            qd(10) = !CODCOMPLO
            qd(11) = !CODCOMPLO 'complemento
            qd(12) = 1
            qd(13) = 99 'statuslanc
            qd(14) = Format(dDataPag, "mm/dd/yyyy") 'data atual
            qd(15) = NomeDoUsuario
            Set RdoAux3 = qd.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
            With RdoAux3
                Do Until .EOF
                    'CARREGA MATRIZ DE DÉBITO
                    
                    nEval = UBound(aDebito)
                    Achou = False
                    For x = 1 To nEval
                        If aDebito(x).nCodReduzido = !CODREDUZIDO And aDebito(x).nAno = RdoAux!ANOEXERCO And aDebito(x).nLanc = RdoAux!CODLANCO And _
                           aDebito(x).nSeq = RdoAux!CODSEQO And _
                           aDebito(x).nParc = RdoAux!NUMPARCO And aDebito(x).nCompl = RdoAux!CODCOMPLO Then
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
       .Close
    End With

    'ADICIONA OS DEBITOS AO GRID DE ORIGEM
    For x = 1 To UBound(aDebito)
        With aDebito(x)
            grdOrigem.AddItem Format(.nCodReduzido, "0000000") & Chr(9) & .nAno & Chr(9) & Format(.nLanc, "00") & Chr(9) & Format(.nSeq, "00") & Chr(9) & _
            Format(.nParc, "00") & Chr(9) & Format(.nCompl, "00") & Chr(9) & .sVencto & Chr(9) & FormatNumber(.nValorTributo, 2) & Chr(9) & _
            FormatNumber(.nValorAtual, 2)
        End With
    Next

    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & Left$(cmbProc.Text, Len(cmbProc) - 5) & " AND STATUSLANC<>5"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
             'BUSCA VALOR LANÇADO
             Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO FROM DEBITOTRIBUTO "
             Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                 If Not IsNull(!ValorTributo) Then
                    nSomaValorTributo = !ValorTributo
                 Else
                    nSomaValorTributo = 0
                 End If
                .Close
             End With

             'BUSCA VALORPAGO
             Sql = "SELECT VALORPAGOREAL,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                  If .RowCount > 0 Then
                      nValorPago = !valorpagoreal
                      dDataPagto = !DataPagamento
                  Else
                      nValorPago = 0
                      dDataPagto = CDate("01/01/1900")
                  End If
                 .Close
             End With

      '       If nValorPago > 0 Then
                'BUSCA TAXA
    '            Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO "
    '            Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
    '            Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
    '            Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
    '            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    '            With RdoAux2
    '                If .RowCount > 0 Then
'                        nValorPago = nValorPago - !VALORTRIBUTO
    '                End If
    '            End With
     '        End If

'             nSomaPago = nSomaPago + nValorPago

           If dDataPagto = "01/01/1900" Then
               ' sDataPagto = "Não Pago"
                Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
                Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
                Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
                Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
                Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND VALORPAGO>0"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                   If .RowCount > 0 Then
                        nValorPago = FormatNumber(!ValorPago, 2)
                        sDataPagto = "Pago sem Data"
                    Else
                        sDataPagto = "Não Pago"
                    End If
                   .Close
                End With
           Else
                If nValorPago = 0 Then
                   sDataPagto = "Pago sem Data"
                Else
                    sDataPagto = Format(dDataPagto, "dd/mm/yyyy")
                End If
           End If
           nSomaPago = nSomaPago + nValorPago

           grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & _
           Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & _
           sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
          .MoveNext
        Loop
       .Close
   End With
End If

nSomaPago = o
For x = 1 To grdDestino.Rows - 1
    If CDbl(grdDestino.TextMatrix(x, 9)) > 0 Then
        nSomaPago = nSomaPago + CDbl(grdDestino.TextMatrix(x, 7))
    End If
Next

lblValorPago.Caption = FormatNumber(nSomaPago, 2)

With grdDestino
    For x = 1 To grdDestino.Rows - 1
        If .TextMatrix(x, 8) = "Não Pago" Then
            nSomaNaoPago = nSomaNaoPago + CDbl(.TextMatrix(x, 7))
        End If
    Next
End With
lblValorNaoPago.Caption = FormatNumber(nSomaNaoPago, 2)

If Right$(cmbProc.Text, 4) <> "SMAR" Then
    'PREENCHE GRID DE ORIGEM
    bVenctoNulo = False
    Sql = "SELECT * FROM vwCNSREPARCELAMENTOO WHERE NUMPROCESSO='" & cmbProc.Text & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
        
            'SE ALGUMA PARCELA NÃO FOR LOCALIZADA NÃO PERMITE O CANCELAMENTO
            If IsNull(!DataVencimento) Then bVenctoNulo = True
            
'****************
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
            qd(14) = IIf(dDataPag = "00:00:00", Format(Now, "mm,dd,yyyy"), Format(dDataPag, "mm/dd/yyyy"))             'data atua
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
                        If !statuslanc = 3 Or !statuslanc = 4 Then
                            aDebito(x).nValorAtual = aDebito(x).nValorAtual + !ValorTotal
                        End If
                        aDebito(x).nValorTributo = FormatNumber(aDebito(x).nValorTributo + !ValorTributo, 2)
                    End If
                   .MoveNext
                Loop
               .Close
            End With
'proximo:
           .MoveNext
        Loop
    End With
    'ADICIONA OS DEBITOS AO GRID DE ORIGEM
    For x = 1 To UBound(aDebito)
        With aDebito(x)
            grdOrigem.AddItem Format(.nCodReduzido, "0000000") & Chr(9) & .nAno & Chr(9) & Format(.nLanc, "00") & Chr(9) & Format(.nSeq, "00") & Chr(9) & _
            Format(.nParc, "00") & Chr(9) & Format(.nCompl, "00") & Chr(9) & .sVencto & Chr(9) & FormatNumber(.nValorTributo, 2) & Chr(9) & _
            FormatNumber(.nValorAtual, 2)
        End With
    Next
    
End If

'VERIFICA SE TEM COMPENSAÇÃO
If Val(lblValorPago.Caption) > 0 Then '
    nTotalACompensar = CDbl(lblValorPago.Caption)
    nTotalAtual = 0
    
    nSobra = nTotalACompensar
    With grdOrigem
        For x = 1 To .Rows - 1
             nValorAChecar = CDbl(.TextMatrix(x, 8))
             nTotalAtual = nTotalAtual + nValorAChecar
             If nSobra > nValorAChecar Then
                .TextMatrix(x, 9) = "06-COMPENSADO"
                nSobra = nSobra - nValorAChecar
             ElseIf nSobra > 0 And nSobra < nValorAChecar Then
                .TextMatrix(x, 9) = "06-COMPENSADO"
                 'busca o novo codigo do complemento
                 Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXCOMPL FROM DEBITOPARCELA WHERE "
                 Sql = Sql & "CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND "
                 Sql = Sql & "CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND "
                 Sql = Sql & "NUMPARCELA=" & .TextMatrix(x, 4)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 nCodCompl = RdoAux!MAXCOMPL + 1
                 RdoAux.Close
                 'cria uma parcela de compensação
                 nLinhaOriginal = x
                .AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & _
               .TextMatrix(x, 4) & Chr(9) & Format(nCodCompl, "00") & Chr(9) & Format(mskDataParc.Text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
                 FormatNumber((nValorAChecar - (nSobra)), 2) & Chr(9) & "03-NÃO PAGO"

                 lblValorExt.Caption = FormatNumber((nValorAChecar - (nSobra)), 2)
                 nSobra = 0
             Else
                .TextMatrix(x, 9) = "03-NÃO PAGO"
             End If
        Next
        
    End With
Else
    'SE NÃO TEM COMPENSAÇÃO, QUER DIZER QUE NENHUMA PARCELA FOI PAGA
    'NESTE CASO BASTA CANCELAR TODAS AS PARCELAS
    With grdOrigem
        For x = 1 To .Rows - 1
            .TextMatrix(x, 9) = "03-NÃO PAGO"
        Next
    End With
End If

nValorAChecar = 0: nValorNaoPago = 0
For x = 1 To grdOrigem.Rows - 1
    If grdOrigem.TextMatrix(x, 9) = "06-COMPENSADO" Then
        nValorAChecar = nValorAChecar + grdOrigem.TextMatrix(x, 8)
    ElseIf grdOrigem.TextMatrix(x, 9) = "03-NÃO PAGO" And grdOrigem.TextMatrix(x, 7) <> "N/A" Then
        nValorNaoPago = nValorNaoPago + grdOrigem.TextMatrix(x, 8)
    End If
Next
lblNP.Caption = FormatNumber(nValorAChecar, 2)
lblVlNComp.Caption = FormatNumber(nValorNaoPago, 2)
With grdOrigem
    If .TextMatrix(.Rows - 1, 9) = "06-COMPENSADO" Then
        If Val(lblValorNaoPago.Caption) > 0 Then
             .AddItem .TextMatrix(.Rows - 1, 0) & Chr(9) & .TextMatrix(.Rows - 1, 1) & Chr(9) & .TextMatrix(.Rows - 1, 2) & Chr(9) & .TextMatrix(.Rows - 1, 3) & Chr(9) & _
            .TextMatrix(.Rows - 1, 4) & Chr(9) & Format(nCodCompl + 1, "00") & Chr(9) & Format(mskDataParc.Text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
              FormatNumber(CDbl(lblValorNaoPago.Caption), 2) & Chr(9) & "03-NÃO PAGO"
        End If
    Else
        If CDbl(lblValorNaoPago.Caption) > CDbl(lblVlNComp.Caption) Then
            .TextMatrix(.Rows - 1, 8) = FormatNumber(CDbl(lblValorNaoPago.Caption) - CDbl(lblVlNComp.Caption), 2)
'             .TextMatrix(.Rows - 1, 8) = FormatNumber(CDbl(lblValorExt.Caption), 2)
       Else
           If lblValorExt.Caption > 0 Then
                .TextMatrix(.Rows - 1, 8) = FormatNumber(CDbl(lblValorExt.Caption), 2)
           End If
       End If
    End If
End With
If lblValorExt.Caption > 0 Then
    grdOrigem.TextMatrix(grdOrigem.Rows - 1, 8) = FormatNumber(CDbl(lblValorExt.Caption), 2)
End If
With grdOrigem
     If .TextMatrix(.Rows - 1, 7) = "N/A" Then
        .FillStyle = flexFillRepeat
        .Row = .Rows - 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &H9FFFC0
     End If
End With

Liberado
End Sub

Private Sub CarregaProcesso2Old()

Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim nValorAtual As Double
Dim dDataVencto As Date
Dim dDataPag As Date
Dim nValorPago As Double, nValorNaoPago As Double
Dim nSomaValorTributo As Double
Dim nSomaPago As Double, nSomaNaoPago As Double
Dim nTotalACompensar As Double
Dim nTotalAtual As Double
Dim nValorAChecar As Double
Dim nSobra As Double
Dim nCodCompl As Integer
Dim x As Integer
Dim dDataPagto As Date, sDataPagto As String
Dim qd As New rdoQuery, aDebito() As Debito, nEval As Integer, Achou As Boolean

ReDim aDebito(0)

grdOrigem.Rows = 1
grdDestino.Rows = 1
Ocupado
'PREENCHE GRID DE DESTINO
If Right$(cmbProc.Text, 4) <> "SMAR" Then
    Sql = "SELECT numprocesso, datareparc From processoreparc WHERE numprocesso = '" & cmbProc.Text & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    mskDataParc.Text = Format(RdoAux!datareparc, "dd/mm/yyyy")
    RdoAux.Close
    
    Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE NUMPROCESSO='" & cmbProc.Text & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        
        nValorPago = 0
        nSomaPago = 0: nSomaNaoPago = 0
        Do Until .EOF
             lblDataProc.Caption = Format(!DATAPROCESSO, "dd/mm/yyyy")
             dDataVencto = Format(!DATADEBASE, "dd/mm/yyyy")
             'dDataPag = Format(!DATAREPARC, "dd/mm/yyyy")
             dDataPag = CDate(mskDataParc.Text)
             'BUSCA VALOR LANÇADO
             Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
             Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
             Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
             Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !numsequencia
             Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
             Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                nValorLanc = !ValorTributo
                If (dDataPag > dDataVencto) Then
                    nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
                    nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                    nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                Else
                    nValorCorrecao = 0
                    nValorJuros = 0
                    nValorMulta = 0
                End If
                nSomaValorTributo = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
                .Close
             End With
                
             'BUSCA VALORPAGO
             Sql = "SELECT VALORPAGO,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !numsequencia
             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND SEQPAG=0"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                  If .RowCount > 0 Then
                      nValorPago = !ValorPago
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
                        nValorPago = nValorPago + !ValorTributo
                    End If
                End With
 '            Else
  '              nValorNaoPago = nValorNaoPago + !VALORTRIBUTO
             End If
                                
             nSomaPago = nSomaPago + nValorPago
             grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!numsequencia, "00") & Chr(9) & _
             Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & _
             sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
            .MoveNext
        Loop
    End With
Else
    'BUSCA O CONTRIBUINTE RESPONSAVEL NA SMAR
    Sql = "SELECT CODREDUZD FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.Text) & " AND CODSEQD=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nCodReduz = RdoAux!CODREDUZD
    RdoAux.Close
    
    'CARREGA A DATA DO REPARCELAMENTO
    Sql = "SELECT * FROM REPARC2TMP WHERE CODREDUZ=" & nCodReduz & " AND CODSEQ=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Sql = Sql & " ORDER BY ANOEXERC,CODLANC,CODSEQ,PARCELAS"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then Exit Sub
        mskDataParc.Text = Format(!DataVencto, "dd/mm/yyyy")
        dDataPag = CDate(mskDataParc.Text)
       .Close
    End With
    
    'CARREGA TODOS OS LANCAMENTOS DO REPARCELAMENTO
    Sql = "SELECT DISTINCT * FROM REPARCTMP WHERE CODREDUZO=" & Val(txtCod.Text) & " AND CODSEQD=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            'VERIFICA SE FOI OU NÃO CANCELADO E POR QUEM
            If !CODSIT > 0 Then
                lblCancel.Visible = True
                Sql = "SELECT CODREDUZ,DATACANCEL,FUNCIONARIOCANCEL FROM REPARC2TMP WHERE CODREDUZ=" & Val(txtCod.Text) & " AND NUMSEQ=" & Left$(cmbProc.Text, Len(cmbProc) - 5)
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If IsNull(RdoAux2!DataCancel) Then
                    lblDataCancel.Caption = "N/A"
                    lblFuncCancel.Caption = "N/A"
                Else
                    If IsDate(RdoAux2!DataCancel) Then
                        lblDataCancel.Caption = Format(RdoAux2!DataCancel, "dd/mm/yyyy")
                    Else
                        lblDataCancel.Caption = ""
                    End If
                    lblFuncCancel.Caption = SubNull(RdoAux2!FUNCIONARIOCANCEL)
                End If
                RdoAux2.Close
            Else
                lblCancel.Visible = False
            End If
        Else
            lblCancel.Visible = False
        End If
        
        Do Until .EOF
            'CARREGA OS TRIBUTOS DE CADA UM DOS LANCAMENTOS
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux3.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = !CODREDUZO
            qd(1) = !CODREDUZO 'codigo
            qd(2) = !ANOEXERCO
            qd(3) = !ANOEXERCO  'ano
            qd(4) = !CODLANCO
            qd(5) = !CODLANCO 'lancamento
            qd(6) = !CODSEQO
            qd(7) = !CODSEQO 'sequencia
            qd(8) = !NUMPARCO
            qd(9) = !NUMPARCO 'numparcela
            qd(10) = !CODCOMPLO
            qd(11) = !CODCOMPLO 'complemento
            qd(12) = 1
            qd(13) = 99 'statuslanc
            qd(14) = Format(dDataPag, "mm/dd/yyyy") 'data atual
            qd(15) = NomeDoUsuario
            Set RdoAux3 = qd.OpenResultset(rdOpenKeyset)
            With RdoAux3
                Do Until .EOF
                    'CARREGA MATRIZ DE DÉBITO
                    nEval = UBound(aDebito)
                    Achou = False
                    For x = 1 To nEval
                        If aDebito(x).nCodReduzido = !CODREDUZIDO And aDebito(x).nAno = RdoAux!ANOEXERCO And aDebito(x).nLanc = RdoAux!CODLANCO And _
                           aDebito(x).nSeq = RdoAux!CODSEQO And _
                           aDebito(x).nParc = RdoAux!NUMPARCO And aDebito(x).nCompl = RdoAux!CODCOMPLO Then
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
       .Close
    End With

    'ADICIONA OS DEBITOS AO GRID DE ORIGEM
    For x = 1 To UBound(aDebito)
        With aDebito(x)
            grdOrigem.AddItem Format(.nCodReduzido, "0000000") & Chr(9) & .nAno & Chr(9) & Format(.nLanc, "00") & Chr(9) & Format(.nSeq, "00") & Chr(9) & _
            Format(.nParc, "00") & Chr(9) & Format(.nCompl, "00") & Chr(9) & .sVencto & Chr(9) & FormatNumber(.nValorTributo, 2) & Chr(9) & _
            FormatNumber(.nValorAtual, 2)
        End With
    Next

    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & Left$(cmbProc.Text, Len(cmbProc) - 5) & " AND STATUSLANC<>5"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
             'BUSCA VALOR LANÇADO
             Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO FROM DEBITOTRIBUTO "
             Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                 nSomaValorTributo = !ValorTributo
                .Close
             End With

             'BUSCA VALORPAGO
             Sql = "SELECT VALORPAGOREAL,DATAPAGAMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
             Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
             Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                  If .RowCount > 0 Then
                      nValorPago = !valorpagoreal
                      dDataPagto = !DataPagamento
                  Else
                      nValorPago = 0
                      dDataPagto = CDate("01/01/1900")
                  End If
                 .Close
             End With

             If nValorPago > 0 Then
                'BUSCA TAXA
                Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO "
                Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
                Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
                Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
'                        nValorPago = nValorPago - !VALORTRIBUTO
                    End If
                End With
             End If

'             nSomaPago = nSomaPago + nValorPago

           If dDataPagto = "01/01/1900" Then
               ' sDataPagto = "Não Pago"
                Sql = "SELECT numdocumento.numdocumento, numdocumento.valorpago "
                Sql = Sql & "FROM parceladocumento INNER JOIN  numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
                Sql = Sql & "WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio
                Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND NUMPARCELA=" & !NumParcela & " AND SEQLANCAMENTO=" & !SeqLancamento
                Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND VALORPAGO>0"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                   If .RowCount > 0 Then
                        nValorPago = FormatNumber(!ValorPago, 2)
                        sDataPagto = "Pago sem Data"
                    Else
                        sDataPagto = "Não Pago"
                    End If
                   .Close
                End With
           Else
                If nValorPago = 0 Then
                   sDataPagto = "Pago sem Data"
                Else
                    sDataPagto = Format(dDataPagto, "dd/mm/yyyy")
                End If
           End If
           nSomaPago = nSomaPago + nValorPago

           grdDestino.AddItem Format(!CODREDUZIDO, "0000000") & Chr(9) & !AnoExercicio & Chr(9) & Format(!CodLancamento, "00") & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & _
           Format(!NumParcela, "00") & Chr(9) & Format(!CODCOMPLEMENTO, "00") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy") & Chr(9) & FormatNumber(nSomaValorTributo, 2) & Chr(9) & _
           sDataPagto & Chr(9) & FormatNumber(nValorPago, 2)
          .MoveNext
        Loop
       .Close
   End With
End If

lblValorPago.Caption = FormatNumber(nSomaPago, 2)

With grdDestino
    For x = 1 To grdDestino.Rows - 1
        If .TextMatrix(x, 8) = "Não Pago" Then
            nSomaNaoPago = nSomaNaoPago + CDbl(.TextMatrix(x, 7))
        End If
    Next
End With
lblValorNaoPago.Caption = FormatNumber(nSomaNaoPago, 2)

If Right$(cmbProc.Text, 4) <> "SMAR" Then
    'PREENCHE GRID DE ORIGEM
    bVenctoNulo = False
    Sql = "SELECT * FROM vwCNSREPARCELAMENTOO WHERE NUMPROCESSO='" & cmbProc.Text & "' ORDER BY ANOEXERCICIO,NUMPARCELA"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
        
            'SE ALGUMA PARCELA NÃO FOR LOCALIZADA NÃO PERMITE O CANCELAMENTO
            If IsNull(!DataVencimento) Then bVenctoNulo = True
            
'****************
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
            qd(14) = IIf(dDataPag = "00:00:00", Format(Now, "mm,dd,yyyy"), Format(dDataPag, "mm/dd/yyyy"))             'data atua
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
                        If !statuslanc = 3 Or !statuslanc = 4 Then
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
    For x = 1 To UBound(aDebito)
        With aDebito(x)
            grdOrigem.AddItem Format(.nCodReduzido, "0000000") & Chr(9) & .nAno & Chr(9) & Format(.nLanc, "00") & Chr(9) & Format(.nSeq, "00") & Chr(9) & _
            Format(.nParc, "00") & Chr(9) & Format(.nCompl, "00") & Chr(9) & .sVencto & Chr(9) & FormatNumber(.nValorTributo, 2) & Chr(9) & _
            FormatNumber(.nValorAtual, 2)
        End With
    Next
    
End If

'VERIFICA SE TEM COMPENSAÇÃO
If Val(lblValorPago.Caption) > 0 Then '
    nTotalACompensar = CDbl(lblValorPago.Caption)
    nTotalAtual = 0
    
    nSobra = nTotalACompensar
    With grdOrigem
        For x = 1 To .Rows - 1
             nValorAChecar = CDbl(.TextMatrix(x, 8))
             nTotalAtual = nTotalAtual + nValorAChecar
             If nSobra > nValorAChecar Then
                .TextMatrix(x, 9) = "06-COMPENSADO"
                nSobra = nSobra - nValorAChecar
             ElseIf nSobra > 0 And nSobra < nValorAChecar Then
                .TextMatrix(x, 9) = "06-COMPENSADO"
                 'busca o novo codigo do complemento
                 Sql = "SELECT MAX(CODCOMPLEMENTO) AS MAXCOMPL FROM DEBITOPARCELA WHERE "
                 Sql = Sql & "CODREDUZIDO=" & .TextMatrix(x, 0) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1) & " AND "
                 Sql = Sql & "CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 3) & " AND "
                 Sql = Sql & "NUMPARCELA=" & .TextMatrix(x, 4)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 nCodCompl = RdoAux!MAXCOMPL + 1
                 RdoAux.Close
                 'cria uma parcela de compensação
                 'alterado, agora o valor da compensação é o valor das parcelas não pagas no carne.
                 'lancamento 64 debito rem.parc.
'                .AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & 64 & Chr(9) & .TextMatrix(x, 3) & Chr(9) & _
               .TextMatrix(x, 4) & Chr(9) & "00" & Chr(9) & Format(mskDataParc.text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
                 FormatNumber(CDbl(lblValorNaoPago.Caption), 2) & Chr(9) & "03-NÃO PAGO"
                 
                .AddItem .TextMatrix(x, 0) & Chr(9) & .TextMatrix(x, 1) & Chr(9) & .TextMatrix(x, 2) & Chr(9) & .TextMatrix(x, 3) & Chr(9) & _
               .TextMatrix(x, 4) & Chr(9) & Format(nCodCompl, "00") & Chr(9) & Format(mskDataParc.Text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
                 FormatNumber((nValorAChecar - (nSobra)), 2) & Chr(9) & "03-NÃO PAGO"
                 '!!! VERIFICAR ANTES DE ALTERAR ESTE VALOR !!!
                 nSobra = 0
             Else
                .TextMatrix(x, 9) = "03-NÃO PAGO"
             End If
        Next
        If nTotalACompensar > nTotalAtual Then
            lblValorExt.Caption = FormatNumber(nTotalACompensar - nTotalAtual, 2)
        End If
    End With
Else
    'SE NÃO TEM COMPENSAÇÃO, QUER DIZER QUE NENHUMA PARCELA FOI PAGA
    'NESTE CASO BASTA CANCELAR TODAS AS PARCELAS
    With grdOrigem
        For x = 1 To .Rows - 1
            .TextMatrix(x, 9) = "03-NÃO PAGO"
        Next
    End With
End If

nValorAChecar = 0
For x = 1 To grdOrigem.Rows - 1
    If grdOrigem.TextMatrix(x, 9) = "06-COMPENSADO" Then
        nValorAChecar = nValorAChecar + grdOrigem.TextMatrix(x, 8)
    End If
Next
lblNP.Caption = FormatNumber(nValorAChecar, 2)


With grdOrigem
     If .TextMatrix(.Rows - 1, 7) = "N/A" Then
        .FillStyle = flexFillRepeat
        .Row = .Rows - 1
        .col = 0
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &H9FFFC0
     End If
End With

Liberado
End Sub


