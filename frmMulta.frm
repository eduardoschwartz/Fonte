VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmMulta 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Multa de Infração"
   ClientHeight    =   4410
   ClientLeft      =   4155
   ClientTop       =   3345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   60
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2550
      Width           =   8925
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4860
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   4
      Top             =   3630
      Width           =   1035
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   7830
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   4005
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
      MICON           =   "frmMulta.frx":0000
      PICN            =   "frmMulta.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7830
      TabIndex        =   5
      ToolTipText     =   "Cancelar Edição"
      Top             =   3600
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
      MICON           =   "frmMulta.frx":008A
      PICN            =   "frmMulta.frx":00A6
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
      Left            =   1620
      TabIndex        =   3
      Top             =   3990
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      MouseIcon       =   "frmMulta.frx":0145
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
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   2
      Top             =   3630
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2160
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   3810
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmMulta.frx":0161
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Observação (Máximo 200 caractéres)"
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   10
      Top             =   2310
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercício....:"
      Height          =   255
      Index           =   2
      Left            =   3690
      TabIndex        =   9
      Top             =   3660
      Width           =   1065
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Vencimento...:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   8
      Top             =   3990
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor da Multa....:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   3660
      Width           =   1455
   End
End
Attribute VB_Name = "frmMulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodReduz As Long, TipoView As Integer

Public Property Let nTipo(nTipoView As Integer)
    TipoView = nTipoView
End Property

Private Sub cmdGerar_Click()
Dim nSeq As Integer, RdoAux As rdoResultset, x As Integer, nSeqO As Integer, RdoAux2 As rdoResultset

If Val(txtValor.Text) = 0 Then
    MsgBox "Digite o valor da multa.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskVenc.Text) Then
    MsgBox "Data inválida.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtAno.Text) < 1990 Or Val(txtAno.Text) > Year(Now) Then
    MsgBox "Exercício fora do intervalo.", vbCritical, "Atenção"
    Exit Sub
End If

'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & 69 & " AND ANOEXERCICIO=" & Val(txtAno.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!SEQMAXIMA) Then
    nSeq = 0
Else
    nSeq = RdoAux!SEQMAXIMA + 1
End If

'GRAVA DÉBITO
'Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USUARIO) VALUES("
'Sql = Sql & nCodReduz & "," & Val(txtAno.Text) & "," & 69 & "," & nSeq & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(mskVenc.Text, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','" & Left$(NomeDeLogin, 25) & "')"
Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
Sql = Sql & nCodReduz & "," & Val(txtAno.Text) & "," & 69 & "," & nSeq & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(mskVenc.Text, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect

Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
Sql = Sql & nCodReduz & "," & Val(txtAno.Text) & "," & 69 & "," & nSeq & "," & 1 & "," & 0 & "," & 20 & "," & Virg2Ponto(txtValor.Text) & ")"
cn.Execute Sql, rdExecDirect

If Trim(txtObs.Text) <> "" Then
'    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES("
'    Sql = Sql & nCodReduz & "," & Val(txtAno.Text) & "," & 69 & "," & nSeq & "," & 1 & "," & 0 & "," & 0 & ",'" & Mask(txtObs.Text) & "','"
'    Sql = Sql & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES("
    Sql = Sql & nCodReduz & "," & Val(txtAno.Text) & "," & 69 & "," & nSeq & "," & 1 & "," & 0 & "," & 0 & ",'" & Mask(txtObs.Text) & "',"
    Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
End If

'GRAVA LINK
With grdTemp
    For x = 1 To grdTemp.Rows - 1
        Sql = "INSERT multainfracao (CODIGO,ANO,LANCAMENTO,SEQUENCIA,PARCELA,COMPLEMENTO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
        Sql = Sql & "CODCOMPLEMENTO) VALUES(" & nCodReduz & "," & Val(txtAno.Text) & "," & 69 & "," & nSeq & "," & 1 & "," & 0 & "," & nCodReduz & ","
        Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & ")"
        cn.Execute Sql, rdExecDirect
    
    
        'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & .TextMatrix(x, 0) & " AND CODLANCAMENTO=" & .TextMatrix(x, 1) & " AND SEQLANCAMENTO=" & .TextMatrix(x, 2) & " AND NUMPARCELA=" & .TextMatrix(x, 3) & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 4)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
            nSeqO = 0
        Else
            nSeqO = RdoAux2!SEQMAXIMA + 1
        End If
        
'        Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES("
'        Sql = Sql & nCodReduz & "," & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & nSeqO & ",'" & Mask(txtObs.Text) & "','"
'        Sql = Sql & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
        Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES("
        Sql = Sql & nCodReduz & "," & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & nSeqO & ",'" & Mask(txtObs.Text) & "',"
        Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With

MsgBox "Multa gerada.", vbInformation, "Informação"
Unload Me
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub Form_Load()

If TipoView = 1 Then
    cmdGerar.Enabled = True
    CarregaLista
    txtAno.Text = Year(Now)
ElseIf TipoView > 1 Then
    cmdGerar.Enabled = False
    CarregaBloco
End If

End Sub

Private Sub CarregaLista()
Dim sAno As String, sLanc As String, sSeq As String, sParc As String, sComp As String, sSit As String, sVencto As String
Dim nValorPrincipal As Double, nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double, nValorTotal As Double
Dim qd As New rdoQuery, RdoAux As rdoResultset, sDA As String, sAj As String

With frmDebitoImob.grdExtrato
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
    For x = 1 To .Rows
        If .CellText(x, 12) = "S" Then
            sAno = .CellText(x, 1)
            sLanc = Left$(.CellText(x, 2), 3)
            sSeq = .CellText(x, 3)
            sParc = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
            sComp = .CellText(x, 5)
            sSit = Left$(.CellText(x, 6), 2)
            sVencto = .CellText(x, 7)
            sDA = .CellText(x, 8)
            sAj = .CellText(x, 9)
                       
            '***********************
            'CARREGA O EXTRATO
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = Val(sAno)
            qd(3) = Val(sAno)
            qd(4) = Val(sLanc)
            qd(5) = Val(sLanc)
            qd(6) = Val(sSeq)
            qd(7) = Val(sSeq)
            qd(8) = Val(sParc)
            qd(9) = Val(sParc)
            qd(10) = Val(sComp)
            qd(11) = Val(sComp)
            qd(12) = 1
            qd(13) = 99
            qd(14) = Format(Now, "mm/dd/yyyy")
            qd(15) = NomeDoUsuario
            Set RdoAux = qd.OpenResultset(rdOpenKeyset)
            With RdoAux
                Do Until .EOF
                    nValorPrincipal = nValorPrincipal + !ValorTributo
                    nValorJuros = nValorJuros + !ValorJuros
                    nValorMulta = nValorMulta + !ValorMulta
                    nValorCorrecao = nValorCorrecao + !ValorCorrecao
                    nValorTotal = nValorTotal + !ValorTotal
                   .MoveNext
                Loop
               .Close
            End With
                       
            grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
            sComp & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
            FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & _
            FormatNumber(nValorMulta, 2) & Chr(9) & FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorTotal, 2)
           nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorTotal = 0
        End If
    Next
End With

End Sub

Private Sub CarregaBloco()

Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nComp As Integer, sSit As String, sVencto As String
Dim nValorPrincipal As Double, nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double, nValorTotal As Double
Dim qd As New rdoQuery, RdoAux As rdoResultset, sDA As String, sAj As String, Sql As String, RdoAux2 As rdoResultset
Dim nSeqMulta As Integer
If frmDebitoImob.grdExtrato.SelectedRow = 0 Then Exit Sub
nCodReduz = Val(frmDebitoImob.txtCod.Text)
With frmDebitoImob.grdExtrato
    nAno = Val(.CellText(.SelectedRow, 1))
    nLanc = Val(Left(.CellText(.SelectedRow, 2), 3))
    nSeq = Val(.CellText(.SelectedRow, 3))
    nParc = Val(.CellText(.SelectedRow, 4))
    nComp = Val(.CellText(.SelectedRow, 5))
End With

If TipoView = 3 Then
    Sql = "SELECT * FROM MULTAINFRACAO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nComp
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nAno = !Ano
        nLanc = !lancamento
        nSeq = !Sequencia
        nParc = !Parcela
        nComp = !Complemento
       .Close
    End With
End If

Sql = "SELECT * FROM MULTAINFRACAO WHERE CODIGO=" & nCodReduz & " AND ANO=" & nAno & " AND LANCAMENTO=" & nLanc & " AND SEQUENCIA=" & nSeq & " AND PARCELA=" & nParc & " AND COMPLEMENTO=" & nComp
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount = 0 Then
        MsgBox "Não localizado vínculo para esta multa.", vbExclamation, "Atenção"
        Exit Sub
    End If
    txtAno.Text = !Ano
    nSeqMulta = !Sequencia
    Sql = "SELECT OBS FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento
    Sql = Sql & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtObs.Text = !obs
        End If
    End With
    
    Do Until .EOF
        'CARREGA O EXTRATO
        Set qd.ActiveConnection = cn
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = !AnoExercicio
        qd(3) = !AnoExercicio
        qd(4) = !CodLancamento
        qd(5) = !CodLancamento
        qd(6) = !SeqLancamento
        qd(7) = !SeqLancamento
        qd(8) = !NumParcela
        qd(9) = !NumParcela
        qd(10) = !CODCOMPLEMENTO
        qd(11) = !CODCOMPLEMENTO
        qd(12) = 1
        qd(13) = 99
        qd(14) = Format(Now, "mm/dd/yyyy")
        qd(15) = NomeDoUsuario
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        With RdoAux
            sDA = IIf(IsNull(!datainscricao), "N", "S")
            sAj = IIf(IsNull(!dataajuiza), "N", "S")
            sVencto = Format(!DataVencimento, "dd/mm/yyyy")
            Do Until .EOF
                nValorPrincipal = nValorPrincipal + !ValorTributo
                nValorJuros = nValorJuros + !ValorJuros
                nValorMulta = nValorMulta + !ValorMulta
                nValorCorrecao = nValorCorrecao + !ValorCorrecao
                nValorTotal = nValorTotal + !ValorTotal
               .MoveNext
            Loop
           .Close
        End With
                   
        grdTemp.AddItem !AnoExercicio & Chr(9) & !CodLancamento & Chr(9) & !SeqLancamento & Chr(9) & !NumParcela & Chr(9) & _
        !CODCOMPLEMENTO & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
        FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & _
        FormatNumber(nValorMulta, 2) & Chr(9) & FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorTotal, 2)
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.NumParcela,"
Sql = Sql & "debitoparcela.CODCOMPLEMENTO, debitoparcela.DataVencimento, Sum(debitotributo.valortributo) AS Total FROM debitoparcela INNER JOIN "
Sql = Sql & "debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO Where debitotributo.CodTributo <> 3 "
Sql = Sql & "GROUP BY debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento,debitoparcela.NumParcela,"
Sql = Sql & "debitoparcela.CODCOMPLEMENTO, debitoparcela.DataVencimento HAVING debitoparcela.codreduzido = " & nCodReduz & " AND debitoparcela.anoexercicio = " & Val(txtAno.Text) & " AND "
Sql = Sql & "debitoparcela.codlancamento = 69 AND debitoparcela.seqlancamento = " & nSeqMulta
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    mskVenc.Text = Format(!DataVencimento, "dd/mm/yyyy")
    txtValor.Text = FormatNumber(!Total, 2)
   .Close
End With

End Sub

