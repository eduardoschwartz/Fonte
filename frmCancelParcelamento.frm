VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCancelParcelamento 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento Manual de Parcelamento"
   ClientHeight    =   6180
   ClientLeft      =   3495
   ClientTop       =   2250
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   8580
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1650
      TabIndex        =   4
      Top             =   600
      Width           =   1275
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1650
      MaxLength       =   6
      TabIndex        =   3
      Top             =   300
      Width           =   1275
   End
   Begin prjChameleon.chameleonButton cmdCancelar 
      Height          =   315
      Left            =   6090
      TabIndex        =   1
      ToolTipText     =   "Cancelar o parcelamento"
      Top             =   5760
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCancelParcelamento.frx":0000
      PICN            =   "frmCancelParcelamento.frx":001C
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
      Left            =   7320
      TabIndex        =   0
      ToolTipText     =   "Sair da Tela"
      Top             =   5760
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
      MICON           =   "frmCancelParcelamento.frx":00BB
      PICN            =   "frmCancelParcelamento.frx":00D7
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
      Left            =   60
      TabIndex        =   19
      Top             =   3960
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
      Left            =   90
      TabIndex        =   21
      Top             =   2190
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
      FormatString    =   $"frmCancelParcelamento.frx":0145
   End
   Begin VB.Label lblVlNComp 
      Caption         =   "Label4"
      Height          =   285
      Left            =   4005
      TabIndex        =   41
      Top             =   6390
      Width           =   870
   End
   Begin VB.Label lblNP 
      Caption         =   "Label3"
      Height          =   285
      Left            =   1845
      TabIndex        =   40
      Top             =   6435
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Correção Aplic.:"
      Height          =   225
      Index           =   14
      Left            =   5670
      TabIndex        =   39
      Top             =   1260
      Width           =   1605
   End
   Begin VB.Label lblValorCorrecao 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7380
      TabIndex        =   38
      Top             =   1260
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%Últ.Parc.Compen..:"
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   13
      Left            =   5730
      TabIndex        =   37
      Top             =   945
      Width           =   1605
   End
   Begin VB.Label lblPerc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7230
      TabIndex        =   36
      Top             =   930
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor não Pago..........:"
      Height          =   225
      Index           =   3
      Left            =   2940
      TabIndex        =   35
      Top             =   1260
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Expediente....:"
      Height          =   225
      Index           =   9
      Left            =   3420
      TabIndex        =   34
      Top             =   1890
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label lblDataCancel 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   6090
      TabIndex        =   33
      Top             =   5490
      Width           =   1215
   End
   Begin VB.Label lblCanceladoPor 
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   1290
      TabIndex        =   32
      Top             =   5490
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelado em.:"
      Height          =   225
      Index           =   12
      Left            =   4950
      TabIndex        =   31
      Top             =   5490
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cancelado por.:"
      Height          =   225
      Index           =   11
      Left            =   90
      TabIndex        =   30
      Top             =   5490
      Width           =   1215
   End
   Begin VB.Label lblValorExpediente 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   1890
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblValorCompensar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   7230
      TabIndex        =   28
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor à Compensar.:"
      ForeColor       =   &H00000080&
      Height          =   225
      Index           =   10
      Left            =   5730
      TabIndex        =   27
      Top             =   1590
      Width           =   1605
   End
   Begin VB.Label lblValorHonorario 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1650
      TabIndex        =   26
      Top             =   1590
      Width           =   1125
   End
   Begin VB.Label lblValorJuros 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4650
      TabIndex        =   25
      Top             =   1590
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Juros Aplicado..:"
      Height          =   225
      Index           =   8
      Left            =   2940
      TabIndex        =   24
      Top             =   1590
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Honorários....:"
      Height          =   225
      Index           =   6
      Left            =   180
      TabIndex        =   23
      Top             =   1590
      Width           =   1485
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
      Left            =   90
      TabIndex        =   22
      Top             =   1920
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
      Left            =   60
      TabIndex        =   20
      Top             =   3690
      Width           =   2910
   End
   Begin VB.Label lblNovo 
      Height          =   255
      Left            =   10560
      TabIndex        =   18
      Top             =   7140
      Visible         =   0   'False
      Width           =   585
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
      Left            =   60
      TabIndex        =   17
      Top             =   30
      Width           =   2910
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo......:"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   630
      Width           =   1485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Parcelamento:"
      Height          =   225
      Index           =   2
      Left            =   3420
      TabIndex        =   15
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label lblDataParc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   5130
      TabIndex        =   14
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido...:"
      Height          =   225
      Index           =   7
      Left            =   180
      TabIndex        =   13
      Top             =   330
      Width           =   1485
   End
   Begin VB.Label lblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3030
      TabIndex        =   12
      Top             =   300
      Width           =   5745
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total Pago...:"
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   11
      Top             =   1260
      Width           =   1485
   End
   Begin VB.Label lblValorTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4650
      TabIndex        =   10
      Top             =   1260
      Width           =   945
   End
   Begin VB.Label lblValorPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1665
      TabIndex        =   9
      Top             =   1260
      Width           =   1125
   End
   Begin VB.Label lblQtdeParc 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1650
      TabIndex        =   8
      Top             =   930
      Width           =   1125
   End
   Begin VB.Label lblQtdePago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4650
      TabIndex        =   7
      Top             =   930
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Parcelas Pagas.:"
      Height          =   225
      Index           =   4
      Left            =   2940
      TabIndex        =   6
      Top             =   945
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde.Parcelas........:"
      Height          =   225
      Index           =   5
      Left            =   180
      TabIndex        =   5
      Top             =   945
      Width           =   1485
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
      Left            =   60
      TabIndex        =   2
      Top             =   5760
      Visible         =   0   'False
      Width           =   4365
   End
End
Attribute VB_Name = "frmCancelParcelamento"
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

Dim RdoAux As rdoResultset, Sql As String, dDataBase As Date, aTributo() As TRIBUTO
Dim nNumproc As Long, nAnoproc As Integer, sNumProc As String, nLinhaOriginal As Integer
Private Sub cmdCancelar_Click()
Dim nCodReduz As Long, nSeq As Integer, nSeq2 As Integer, nNumDoc As Long, nValorTxExp As Double, RdoAux2 As rdoResultset

If lblCancel.Visible Then
    MsgBox "Este parcelamento já foi cancelado.", vbExclamation, "Atenção"
    Exit Sub
End If

If lblNome.Caption = "" Or lblDataParc.Caption = "" Then
    MsgBox "Selecione o proprietário e o processo de parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(lblNovo.Caption) = 0 Then
    MsgBox "Este processo é antigo e não pode ser " & vbCrLf & "cancelado por esta tela.", vbExclamation, "Atenção"
    Exit Sub
End If

nCodReduz = Val(txtCod.Text)
sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)


If MsgBox("Os débitos do reparcelamento serão cancelados." & vbCrLf & vbCrLf & "Deseja continuar ?", vbQuestion + vbYesNo, "CONFIRMAÇÂO DE CANCELAMENTO !!!") = vbNo Then Exit Sub
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
                Sql = Sql & Val(SubNull(!CODMOEDA)) & "," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & "," & Val(SubNull(!numcertidao)) & "," & IIf(IsNull(!datainscricao), "Null", "'" & Format(!datainscricao, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!dataajuiza), "Null", "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!ValorJuros), 0, !ValorJuros) & ",'"
'                Sql = Sql & txtNumProc.Text & "','" & Left$(NomeDeLogin, 25) & "')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,"
                Sql = Sql & "DATAAJUIZA,VALORJUROS,NUMPROCESSO,USERID) VALUES(" & Val(txtCod.Text) & "," & !AnoExercicio & "," & !CodLancamento & ","
                Sql = Sql & !SeqLancamento & "," & !NumParcela & "," & Val(grdOrigem.TextMatrix(x, 4)) & "," & Val(Left$(grdOrigem.TextMatrix(x, 8), 2)) & ",'" & Format(grdOrigem.TextMatrix(nLinhaOriginal, 5), "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',"
                Sql = Sql & Val(SubNull(!CODMOEDA)) & "," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & "," & Val(SubNull(!numcertidao)) & "," & IIf(IsNull(!datainscricao), "Null", "'" & Format(!datainscricao, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!dataajuiza), "Null", "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'") & "," & IIf(IsNull(!ValorJuros), 0, !ValorJuros) & ",'"
                Sql = Sql & txtNumProc.Text & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
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

Liberado
MsgBox "O cancelamento do reparcelamento foi executado com sucesso.", vbExclamation, "Atenção"
grdDestino.Rows = 1
grdOrigem.Rows = 1

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
End Sub

Private Sub txtCod_Change()
lblNome.Caption = ""
txtNumProc.Text = ""
lblDataParc.Caption = ""
lblCancel.Visible = False
lblValorPago.Caption = "0,00"
lblValorTotal.Caption = "0,00"
lblQtdePago.Caption = "0"
lblQtdeParc.Caption = "0"
grdOrigem.Rows = 1
grdDestino.Rows = 1
End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_LostFocus()
Dim nCodReduz As Long, sTipoCod As String

If Val(txtCod.Text) = 0 Then Exit Sub
If Val(txtCod.Text) = 0 Then
    lblNome.Caption = ""
    Exit Sub
End If
If Val(txtCod.Text) < 100000 Then
    sTipoCod = "I"
ElseIf Val(txtCod.Text) >= 100000 And Val(txtCod.Text) < 500000 Then
    sTipoCod = "M"
ElseIf Val(txtCod.Text) >= 500000 Then
    sTipoCod = "C"
End If
txtCod.Text = Format(txtCod.Text, "000000")
nCodReduz = Val(txtCod.Text)
lblNome.Caption = ""
If sTipoCod = "I" Then
    Sql = "SELECT PROPRIETARIO.CODCIDADAO, CIDADAO.NOMECIDADAO "
    Sql = Sql & "FROM PROPRIETARIO INNER JOIN   CIDADAO ON   PROPRIETARIO.CodCidadao = CIDADAO.CodCidadao "
    Sql = Sql & "Where PROPRIETARIO.CODREDUZIDO =" & nCodReduz & " AND TIPOPROP='P'"
ElseIf sTipoCod = "M" Then
    Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO Where CODIGOMOB =" & nCodReduz
ElseIf sTipoCod = "C" Then
    Sql = "SELECT NOMECIDADAO FROM CIDADAO Where CODCIDADAO =" & nCodReduz
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux.RowCount > 0 Then
         If sTipoCod = "I" Or sTipoCod = "C" Then
            lblNome.Caption = !NomeCidadao
         ElseIf sTipoCod = "M" Then
            lblNome.Caption = !RazaoSocial
         End If
    Else
       MsgBox "Código não Cadastrado.", vbExclamation, "Atenção"
       txtCod.SetFocus
       Exit Sub
    End If
    .Close
End With

End Sub

Private Sub txtNumProc_Change()
lblDataParc.Caption = ""
lblValorPago.Caption = "0,00"
lblValorTotal.Caption = "0,00"
lblQtdePago.Caption = "0"
lblQtdeParc.Caption = "0"
lblCancel.Visible = False
grdOrigem.Rows = 1
grdDestino.Rows = 1
End Sub

Private Sub txtNumProc_GotFocus()
txtNumProc.SelStart = 0
txtNumProc.SelLength = Len(txtNumProc.Text)
End Sub

Private Sub txtNumProc_LostFocus()
Dim nValorPago As Double, nNovo As Integer, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, nQtdeParc As Integer, nValorCalc As Double, x As Integer, nValorCorrecao As Double
On Error Resume Next
Ocupado
nValorPago = 0
txtNumProc.Text = Replace$(txtNumProc.Text, "-", "")
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
         On Error Resume Next
         Sql = "SELECT jurosapl, honorario From destinoreparc WHERE codreduzido = " & Val(txtCod.Text) & " AND NUMPROCESSO='" & CStr(nNumproc) & "/" & CStr(nAnoproc) & "'"
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
                   'If !CodTributo = 565 Then
                   '   aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo * 2, 2)
                   '   aDebito(nEval).nValorAtual = !ValorTotal * 2
                   'Else
                      aDebito(nEval).nValorTributo = FormatNumber(!ValorTributo, 2)
                      aDebito(nEval).nValorAtual = !ValorTotal
                   'End If
                   
                   
                   
                   
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
        If Val(lblValorTotal.Caption) > 0 Then
             .AddItem .TextMatrix(.Rows - 1, 0) & Chr(9) & .TextMatrix(.Rows - 1, 1) & Chr(9) & .TextMatrix(.Rows - 1, 2) & Chr(9) & .TextMatrix(.Rows - 1, 3) & Chr(9) & _
            .TextMatrix(.Rows - 1, 4) & Chr(9) & Format(nCodCompl + 1, "00") & Chr(9) & Format(mskDataParc.Text, "dd/mm/yyyy") & Chr(9) & "N/A" & Chr(9) & _
              FormatNumber(CDbl(lblValorNaoPago.Caption), 2) & Chr(9) & "03-NÃO PAGO"
        End If
    Else
        If CDbl(lblValorTotal.Caption) > CDbl(lblVlNComp.Caption) Then
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

Private Sub Corrige()

Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, qd As New rdoQuery
Dim nCodReduzido As Long, sNumProc As String, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim nPrincipal As Double, nJuros As Double, nMulta As Double, nCorrecao As Double, nJurosApl As Double, nTotal As Double, nCodTributo As Integer, nHonorario As Double
Dim nNumproc As Long, nAnoproc As Integer

nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
nAnoproc = Right$(txtNumProc.Text, 4)
sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)

Sql = "SELECT destinoreparc.numprocesso, destinoreparc.codreduzido, destinoreparc.anoexercicio, destinoreparc.codlancamento, destinoreparc.numsequencia, "
Sql = Sql & "destinoreparc.NumParcela , destinoreparc.CODCOMPLEMENTO FROM destinoreparc INNER JOIN processoreparc ON destinoreparc.numprocesso = processoreparc.numprocesso INNER JOIN "
Sql = Sql & "debitoparcela ON destinoreparc.codreduzido = debitoparcela.codreduzido AND destinoreparc.anoexercicio = debitoparcela.anoexercicio AND destinoreparc.codlancamento = debitoparcela.codlancamento AND "
Sql = Sql & "destinoreparc.numsequencia = debitoparcela.seqlancamento AND destinoreparc.NumParcela = debitoparcela.NumParcela Where (processoreparc.Cancelado = 0)  AND (destinoreparc.codreduzido=" & Val(txtCod.Text) & ") AND (destinoreparc.numprocesso='" & sNumProc & "')  "
Sql = Sql & "ORDER BY destinoreparc.codreduzido, destinoreparc.numprocesso, destinoreparc.anoexercicio, destinoreparc.codlancamento, destinoreparc.numsequencia,destinoreparc.NumParcela , destinoreparc.CODCOMPLEMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNumProc = !numprocesso
        nCodReduzido = !CODREDUZIDO
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !numsequencia
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        
        nPrincipal = 0: nJuros = 0: nMulta = 0: nCorrecao = 0: nTotal = 0: nJurosApl = 0: nHonorario = 0
        'CARREGA O EXTRATO
        Set qd.ActiveConnection = cn
        qd.QueryTimeout = 0
        On Error Resume Next
        RdoAux2.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduzido
        qd(1) = nCodReduzido
        qd(2) = nAno: qd(3) = nAno
        qd(4) = nLanc: qd(5) = nLanc
        qd(6) = nSeq: qd(7) = nSeq
        qd(8) = nParc: qd(9) = nParc
        qd(10) = nCompl: qd(11) = nCompl
        qd(12) = 0: qd(13) = 99
        qd(14) = Format("01/01/1990", "mm/dd/yyyy")
        qd(15) = NomeDoUsuario
        Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
        
        With RdoAux2
            If RdoAux2.RowCount > 0 Then
                Do Until .EOF
                    nCodTributo = !CodTributo
                    If nCodTributo = 26 Then
                        nCorrecao = !ValorTributo
                    ElseIf nCodTributo = 112 Then
                        nMulta = !ValorTributo
                    ElseIf nCodTributo = 113 Then
                        nJuros = !ValorTributo
                    ElseIf nCodTributo = 585 Then
                        nJurosApl = !ValorTributo
                    ElseIf nCodTributo = 90 Then
                        nHonorario = !ValorTributo
                    ElseIf nCodTributo = 587 Then
                        'CORREC.APL.
                    Else
                        nPrincipal = nPrincipal + !ValorTributo
                    End If
proximo:
                    .MoveNext
                Loop
              End If
           .Close
        End With
        Sql = "UPDATE DESTINOREPARC SET VALORLIQUIDO=" & Virg2Ponto(CStr(nPrincipal)) & ",JUROS=" & Virg2Ponto(CStr(nJuros)) & ",MULTA="
        Sql = Sql & Virg2Ponto(CStr(nMulta)) & ",CORRECAO=" & Virg2Ponto(CStr(nCorrecao)) & ",VALORPRINCIPAL=" & Virg2Ponto(CStr(nPrincipal + nJurosApl))
        Sql = Sql & ",JUROSAPL=" & Virg2Ponto(CStr(nJurosApl)) & ",HONORARIO=" & Virg2Ponto(CStr(nHonorario)) & " WHERE NUMPROCESSO='" & sNumProc & "' AND CODREDUZIDO=" & nCodReduzido & " AND ANOEXERCICIO=" & nAno & " AND "
        Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND NUMSEQUENCIA=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

End Sub
