VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRelBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório Bancário"
   ClientHeight    =   2220
   ClientLeft      =   1530
   ClientTop       =   3300
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2220
   ScaleWidth      =   5835
   Begin VB.CheckBox chk1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Todos"
      Height          =   210
      Left            =   1065
      TabIndex        =   1
      Top             =   705
      Value           =   1  'Checked
      Width           =   840
   End
   Begin VB.ComboBox cmbBanco 
      Height          =   315
      Left            =   2070
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   660
      Width           =   3525
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1245
      TabIndex        =   2
      Top             =   150
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelBanco.frx":0000
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
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   4005
      TabIndex        =   3
      Top             =   165
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelBanco.frx":001C
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
      Left            =   4485
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   1245
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmRelBanco.frx":0038
      PICN            =   "frmRelBanco.frx":0054
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3270
      TabIndex        =   5
      ToolTipText     =   "Cancelar Edição"
      Top             =   1260
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
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
      MICON           =   "frmRelBanco.frx":00C2
      PICN            =   "frmRelBanco.frx":00DE
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
      Height          =   225
      Left            =   465
      TabIndex        =   6
      Top             =   1305
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   195
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2985
      TabIndex        =   11
      Top             =   210
      Width           =   1455
   End
   Begin VB.Label lblMsg 
      Caption         =   "Selecione as Datas de Inicio e Término"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   1830
      Width           =   5535
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2700
      TabIndex        =   9
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   135
      TabIndex        =   8
      Top             =   1305
      Width           =   270
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco..:"
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   705
      Width           =   765
   End
End
Attribute VB_Name = "frmRelBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset

Private Sub chk1_Click()
If chk1.Value = 1 Then
    cmbBanco.Enabled = False
Else
    cmbBanco.Enabled = True
End If

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.Value = (nPosF * 100) / nTotal
Else
   Pb.Value = 100
End If
lblPB.Caption = FormatNumber(Pb.Value, 2)

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub cmdCalculo_Click()
Dim sComputer As String, sArquivo As String, sFuncionario As String, sBanco As String
Dim nNumReg As Integer, nValorTotal As Double, sCodRemessa As String, sCodConvenio As String
Dim dDataGeracao As Date, nNumSequencia As Integer, nLayout As Integer, sTr As String
Dim sContaPref As String, dDataPagto As Date, dDataCredito As Date, sNumDoc As String
Dim sCodReduz As String, nAnoExercicio As Integer, nCodLancamento As Integer, nSeqLancamento As Integer
Dim nNumParcela As Integer, nCodComplemento As Integer, dDataVencto As Date, nValorLancado As Double
Dim nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorCalculado As Double
Dim nValorDif As Double, nValorPago As Double, nValorTarifa As Double, nValorPagoReal As Double
Dim sSituacao As String, nSequencia As Integer, nValorBanco As Double, nRegBanco As Integer
Dim sRetorno As String, sAgencia As String, qd As New rdoQuery, xId As Integer, nNumRec As Long


If Not IsDate(mskDataIni.text) Then
    MsgBox "Data de Inicio inválido", vbExclamation, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataFim.text) Then
    MsgBox "Data de Fim inválido", vbExclamation, "atenção"
    Exit Sub
End If

If CDate(mskDataIni.text) > CDate(mskDataFim.text) Then
    MsgBox "Data de Inicio tem que ser maior que data de termino", vbExclamation, "atenção"
    Exit Sub
End If


'Valores Default
'sComputer = NomeDoUsuario: sArquivo = "": sFuncionario = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8)
sComputer = NomeDoUsuario: sArquivo = "": sFuncionario = NomeDeLogin
nNumReg = 0: nValorTotal = 0: sCodRemessa = "": sCodConvenio = "": dDataGeracao = Format(Now, "mm/dd/yyyy")
nNumSequencia = 0: nLayout = 0: sTr = "": sContaPref = "": nSequencia = 0: nValorBanco = 0
nRegBanco = 0: sRetorno = "": sAgencia = ""

Sql = "DELETE FROM BAIXATMP WHERE COMPUTADOR='" & sComputer & "'"
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn
xId = 0
Sql = "SELECT DEBITOPAGO.*,BANCO.NOMEBANCO , NUMDOCUMENTO.VALORTAXADOC "
Sql = Sql & "FROM DEBITOPAGO INNER JOIN BANCO ON DEBITOPAGO.CODBANCO = BANCO.CODBANCO INNER JOIN "
Sql = Sql & "NUMDOCUMENTO ON DEBITOPAGO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO "
Sql = Sql & "Where (DEBITOPAGO.NumDocumento Is Not Null) And (DEBITOPAGO.RESTITUIDO Is Null) "
Sql = Sql & "AND (DEBITOPAGO.DATARECEBIMENTO BETWEEN CONVERT(DATETIME, '" & Format(mskDataIni.text, "mm/dd/yyyy") & "', 102) AND CONVERT(DATETIME, '" & Format(mskDataFim.text, "mm/dd/yyyy") & "', 102)) "
If chk1.Value = 0 Then
    Sql = Sql & " AND DEBITOPAGO.CODBANCO=" & cmbBanco.ItemData(cmbBanco.ListIndex)
End If
Sql = Sql & " ORDER BY DEBITOPAGO.DATARECEBIMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        If xId Mod 10 = 0 Then
            CallPb CLng(xId), nNumRec
        End If
        
        sCodReduz = CStr(!CODREDUZIDO)
        nAnoExercicio = !AnoExercicio
        nCodLancamento = !CodLancamento
        nSeqLancamento = !SeqLancamento
        nNumParcela = !NumParcela
        nCodComplemento = !CODCOMPLEMENTO
        dDataVencto = Format(Now, "mm/dd/yyyy")
        sNumDoc = !NumDocumento
        'If Val(sNumDoc) = 2406038 Then MsgBox "AQUI"
        dDataPagto = Format(!DataPagamento, "mm/dd/yyyy")
        dDataCredito = Format(!DATARECEBIMENTO, "mm/dd/yyyy")
        sBanco = Format(!CodBanco, "000") & "-" & !NomeBanco
        
        'BUSCA VALOR LANÇADO
        Sql = "SELECT SUM(VALORTRIBUTO) AS VALORTRIBUTO,DATAVENCIMENTO,DATADEBASE FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON "
        Sql = Sql & "DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO "
        Sql = Sql & " AND DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NUMPARCELA = DEBITOPARCELA.NUMPARCELA AND DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
        Sql = Sql & "WHERE DEBITOPARCELA.CODREDUZIDO=" & !CODREDUZIDO & " AND DEBITOPARCELA.ANOEXERCICIO = " & !AnoExercicio
        Sql = Sql & " AND DEBITOPARCELA.CODLANCAMENTO=" & !CodLancamento & " AND DEBITOPARCELA.NUMPARCELA=" & !NumParcela & " AND DEBITOPARCELA.SEQLANCAMENTO=" & !SeqLancamento
        Sql = Sql & " AND DEBITOPARCELA.CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
        Sql = Sql & " GROUP BY DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE"
         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
         With RdoAux2
            If .RowCount > 0 Then
                dDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
                nValorCorrecao = FormatNumber(CalculaCorrecao2(!ValorTributo, dDataVencto, dDataPagto), 2)
                nValorJuros = FormatNumber(CalculaJuros2(!ValorTributo + nValorCorrecao, dDataVencto, dDataPagto), 2)
                nValorMulta = FormatNumber(CalculaMulta2(!ValorTributo + nValorCorrecao, dDataVencto, dDataPagto), 2)
                nValorLancado = Format(!ValorTributo, "#0.00")
            Else
                nValorLancado = 0
                nValorMulta = 0
                nValorJuros = 0
                nValorCorrecao = 0
            End If
           .Close
        End With
                
        nValorTarifa = RdoAux!VALORTAXADOC
        If nValorTarifa = 0 Then
            Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If .RowCount > 0 Then
                    nValorTarifa = Format(!ValorTributo, "#0.00")
                Else
                    nValorTarifa = 0
                End If
               .Close
            End With
        End If
        nValorPago = !ValorPago
        nValorPagoReal = !VALORPAGOREAL
        
        nValorCalculado = 0
        nValorDif = nValorPagoReal - (nValorLancado + nValorJuros + nValorCorrecao + nValorMulta + nValorTarifa)
        
        sSituacao = ""
        
        
        'GRAVA TABELA TEMP
        
        qd.Sql = "{ Call spGRAVABAIXATMP(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = sComputer
        qd(1) = sArquivo
        qd(2) = sFuncionario
        qd(3) = sBanco
        qd(4) = nNumReg
        qd(5) = nValorTotal
        qd(6) = sCodRemessa
        qd(7) = sCodConvenio
        qd(8) = Format(dDataGeracao, "mm/dd/yyyy")
        qd(9) = nNumSequencia
        qd(10) = nLayout
        qd(11) = sTr
        qd(12) = sContaPref
        qd(13) = Format(dDataPagto, "mm/dd/yyyy")
        qd(14) = Format(dDataCredito, "mm/dd/yyyy")
        qd(15) = sNumDoc
        qd(16) = sCodReduz
        qd(17) = nAnoExercicio
        qd(18) = nCodLancamento
        qd(19) = nSeqLancamento
        qd(20) = nNumParcela
        qd(21) = nCodComplemento
        qd(22) = Null 'DATA VENCIMENTO
        qd(23) = Virg2Ponto(CStr(nValorLancado))
        qd(24) = Virg2Ponto(CStr(nValorJuros))
        qd(25) = Virg2Ponto(CStr(nValorMulta))
        qd(26) = Virg2Ponto(CStr(nValorCorrecao))
        qd(27) = Virg2Ponto(CStr(nValorCalculado))
        qd(28) = Virg2Ponto(CStr(nValorDif))
        qd(29) = Virg2Ponto(CStr(nValorPago))
        qd(30) = Virg2Ponto(CStr(nValorTarifa))
        qd(31) = Virg2Ponto(CStr(nValorPagoReal))
        qd(32) = sSituacao
        qd(33) = 0
        qd(34) = nValorBanco
        qd(35) = nRegBanco
        qd(36) = sRetorno
        qd(37) = sAgencia
        Set RdoAux2 = qd.OpenResultset(rdOpenForwardOnly)
        xId = xId + 1
       .MoveNext
    Loop
   .Close
End With

Pb.Value = 100
lblPB.Caption = "100 %"

frmReport.ShowReport "RELBANCO", frmMdi.hwnd, Me.hwnd
Sql = "DELETE FROM BAIXATMP WHERE COMPUTADOR='" & sComputer & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me

Sql = "SELECT CODBANCO,NOMEBANCO FROM BANCO WHERE CODBANCO>0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbBanco.AddItem !NomeBanco
        cmbBanco.ItemData(cmbBanco.NewIndex) = !CodBanco
       .MoveNext
    Loop
End With
cmbBanco.ListIndex = 0
cmbBanco.Enabled = False

End Sub

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

If Year(dDataBase) > Year(dDataVencto) Then
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

