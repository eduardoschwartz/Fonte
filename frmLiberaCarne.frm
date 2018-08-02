VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLiberaCarne 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberação de Carnê de Parcelamento"
   ClientHeight    =   2550
   ClientLeft      =   6225
   ClientTop       =   5985
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4905
   Begin VB.OptionButton optGuia 
      Caption         =   "Boleto"
      Height          =   195
      Index           =   1
      Left            =   3915
      TabIndex        =   17
      Top             =   1125
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.OptionButton optGuia 
      Caption         =   "Normal"
      Enabled         =   0   'False
      Height          =   195
      Index           =   0
      Left            =   3015
      TabIndex        =   16
      Top             =   1125
      Width           =   825
   End
   Begin VB.CheckBox chkTxExp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00EEEEEE&
      Caption         =   "Emitir com Taxa de Expediente..:"
      Enabled         =   0   'False
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
      Left            =   60
      TabIndex        =   15
      Top             =   1890
      Width           =   3255
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   0
      Top             =   330
      Width           =   1275
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      TabIndex        =   4
      Top             =   1080
      Width           =   1275
   End
   Begin VB.TextBox txtAno 
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
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   3840
      MaxLength       =   6
      TabIndex        =   3
      Top             =   210
      Width           =   765
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3540
      TabIndex        =   5
      ToolTipText     =   "Imprime o Carnê de Parcelamento"
      Top             =   1710
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmLiberaCarne.frx":0000
      PICN            =   "frmLiberaCarne.frx":001C
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
      Height          =   345
      Left            =   3540
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   2130
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   609
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
      MICON           =   "frmLiberaCarne.frx":0176
      PICN            =   "frmLiberaCarne.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNome 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   150
      TabIndex        =   14
      Top             =   690
      Width           =   4485
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido...:"
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   13
      Top             =   390
      Width           =   1485
   End
   Begin VB.Label lblDataParc 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1830
      TabIndex        =   12
      Top             =   1530
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data do Parcelamento:"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   1530
      Width           =   1665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo.....:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1140
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
      Left            =   0
      TabIndex        =   9
      Top             =   30
      Width           =   2910
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   495
   End
   Begin VB.Label lblCancel 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CANCELADO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   2190
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.Label lblNumProc 
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   2580
      Width           =   1575
   End
   Begin VB.Label lblAnoProc 
      Height          =   315
      Left            =   2010
      TabIndex        =   1
      Top             =   2580
      Width           =   1635
   End
End
Attribute VB_Name = "frmLiberaCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim dDataBase As Date, nQtdeParc As Integer
Dim xImovel As clsImovel

Private Sub cmdPrint_Click()
EmiteBoleto
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
dDataBase = CDate(Mid$(frmMdi.Sbar.Panels(6).Text, 12, 2) & "/" & Mid$(frmMdi.Sbar.Panels(6).Text, 15, 2) & "/" & Right$(frmMdi.Sbar.Panels(6).Text, 4))
txtAno.Text = Year(Now)
Set xImovel = New clsImovel
End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_Change()
lblCancel.Visible = False
lblNome.Caption = ""
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
            lblNome.Caption = !nomecidadao
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
lblCancel.Visible = False
End Sub

Private Sub txtNumProc_LostFocus()
Dim nNumproc As Long, nAnoproc As Integer
On Error Resume Next
If Trim$(txtNumProc.Text) <> "" Then
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
        nAnoproc = Right$(txtNumProc.Text, 4)
        lblNumProc.Caption = nNumproc
        lblAnoProc.Caption = nAnoproc
        Sql = "SELECT NUMPROC,ANOPROC,DATAREPARC,QTDEPARCELA,CANCELADO FROM PROCESSOREPARC  WHERE CODIGORESP=" & Val(txtCod.Text) & " AND NUMPROC=" & nNumproc & " AND ANOPROC=" & nAnoproc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
                txtNumProc.SetFocus
                Exit Sub
            Else
                lblDataParc.Caption = Format(!datareparc, "dd/mm/yyyy")
                nQtdeParc = !qtdeparcela
                lblCancel.Visible = !Cancelado
            End If
           .Close
        End With
    Else
        MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
        txtNumProc.SetFocus
    End If
End If

End Sub

Private Sub EmiteBoleto()

Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, Y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, sDataDam As String, bBoleto As Boolean, sValor As String, dDataVencto As Date
bBoleto = False

If lblCancel.Visible = True Then
    MsgBox "Parcelamento Cancelado.", vbExclamation, "Atenção"
    Exit Sub
End If

If lblNome.Caption = "" Or lblDataParc.Caption = "" Then
    MsgBox "Selecione o proprietário e o processo de parcelamento.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno.Text) < Year(Now) Or Val(txtAno.Text) > Year(Now) + 6 Then
    MsgBox "Ano inválido.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Emitir as parcelas do parcelamento de " & txtAno.Text, vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
Ocupado
If chkTxExp.value = vbChecked Then
    'BUSCA O VALOR DA TAXA DE EXPEDIENTE
    Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 1"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nValorExp = FormatNumber(!VALORPARCELA, 2)
       .Close
    End With
Else
    nValorExp = 0
End If

'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

sLib = "LIBERACAO"

nCodReduz = Val(txtCod.Text)
'ENDEREÇO DO CONTRIBUINTE
Select Case Val(txtCod.Text)
    Case 1 To 99999
        sTipoImposto = "REPARCEL."
        sSetor = "IMOBILIÁRIO"
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
        sSetor = "MOBILIÁRIO"
        sTipoImposto = "REPARCEL."
        sNomeResp = lblNome.Caption
        sNumInsc = txtCod.Text
        sCodReduz = txtCod.Text
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
        sSetor = "TAXAS DIVERSAS"
        sTipoImposto = "REPARCEL."
        sNomeResp = lblNome.Caption
        sNumInsc = txtCod.Text
        sCodReduz = txtCod.Text
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
        
        Sql = "SELECT codcidadao,nomecidadao,cpf,cnpj,rg from cidadao WHERE CODCIDADAO=" & Val(txtCod.Text)
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

sNumProc = lblNumProc.Caption & "/" & lblAnoProc.Caption
dDataProc = lblDataParc.Caption
Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, statuslanc, datavencimento, datadebase, codmoeda, "
Sql = Sql & "numerolivro , paginalivro, numcertidao, datainscricao, dataajuiza, valorjuros, numprocesso, intacto From debitoparcela "
Sql = Sql & "WHERE debitoparcela.codreduzido = " & Val(txtCod.Text) & " AND debitoparcela.codlancamento = 20 AND DEBITOPARCELA.NUMPARCELA > 1 AND "
Sql = Sql & "YEAR(debitoparcela.datavencimento) = " & txtAno.Text & " AND debitoparcela.numprocesso = '" & sNumProc & "' AND STATUSLANC=3 order by anoexercicio,numparcela"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Não existem parcelas a serem impressas." & vbCrLf & "Verifique se estas parcelas não estão bloqueadas.", vbExclamation, "Atenção"
        Liberado
        Exit Sub
    End If
    x = 1
    
    Do Until .EOF
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        sDataDam = Format(!DataVencimento, "dd/mm/yyyy")
        
        Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO <> 3"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nValorParc = FormatNumber(!soma, 2)
           .Close
        End With
        
     '   sql = "SELECT MIN(NUMDOCUMENTO) AS MINIMO FROM PARCELADOCUMENTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
     '   sql = sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
     '   Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
     '   If Not IsNull(RdoAux2!MINIMO) Then
     '       nNumDoc = RdoAux2!MINIMO
     '       RdoAux2.Close
     '   Else
     '       RdoAux2.Close
            Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nNumDoc = !maximo + 1
            End With
            'GRAVA NA TABELA NUMDOCUMENTO
            Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor,valorguia) VALUES("
            Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0,0,'" & NomeDeLogin & " (LIBERAÇÃO CARNÊ)" & "'," & Virg2Ponto(RemovePonto(CStr(nValorParc))) & ")"
            cn.Execute Sql, rdExecDirect
            'GRAVA NA TABELA PARCELADOCUMENTO
            Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & ")"
            cn.Execute Sql, rdExecDirect
      '  End If
        
        

        nNumGuia = nNumDoc

        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)


        Sql = "SELECT VALORALIQ FROM TRIBUTOALIQUOTA WHERE CODTRIBUTO=3 AND ANO=" & Year(Now)
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nValorTaxa = RdoAux2!valoraliq
        RdoAux2.Close
        
        sValorParc = Format(nValorParc, "#0.00")
        nValorGuia = sValorParc + CDbl(nValorExp)
        nValorDoc = nValorGuia
    '**** GERADOR DE CÓDIGO DE BARRAS ********
    'sNossoNumero = "0760369"
    
    If bBoleto Then
        sNossoNumero = "2678478"
        sDigitavel = "001900000"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
        sDigitavel = sDigitavel & sDv
        
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
        sValor = nValorDoc
        dDataVencto = CDate(sDataDam)
      '  nNumDoc = Val(sNumDoc2)
        sDadosLanc = sTipoImposto
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
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numproc,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & Val(txtCod.Text) & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 1, nParc) & "," & nQtdeParc & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & "Parcelamento: " & Left$(txtNumProc.Text, 25) & "','" & Left$(txtNumProc.Text, 25) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
        cn.Execute Sql, rdExecDirect
        x = x + 1
       .MoveNext
    Loop
   .Close
End With


'******* INTEGRATIVA *****
'ConectaIntegrativa 'ABRE CONEXÃO COM A INTEGRATIVA
'
'sql = "SELECT * FROM debitoparcela WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND "
'sql = sql & "(debitoparcela.numparcela = 1) AND (debitoparcela.numprocesso = '" & lblNumProc.Caption & "/" & lblAnoProc.Caption & "')"
'Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'With RdoAux2
'    If .RowCount > 0 Then
'        dDataPrimeiraParc = !DataVencimento
'    End If
'    sql = "SELECT valortributo FROM debitotributo WHERE codreduzido = " & nCodReduz & " and anoexercicio=" & !AnoExercicio & " and "
'    sql = sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and "
'    sql = sql & "codcomplemento=" & !CODCOMPLEMENTO & " and codtributo=90"
'    Set RdoAux3 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux3
'        If .RowCount > 0 Then
'            nValorTotalHon = !ValorTributo * nQtdeParc
'        Else
'            nValorTotalHon = 0
'        End If
'       .Close
'    End With
'   .Close
'End With
'
'
''*** VERIFICA SE O PARCELAMENTO JÁ EXISTE NA TABELA ACORDOS **
'sql = "select * from acordos where idacordo=" & Val(lblNumProc.Caption) & " and anoacordo=" & Val(lblAnoProc.Caption)
'Set RdoAux = cnInt.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'If RdoAux.RowCount = 0 Then
'    'GRAVA O ACORDO
'    sql = "insert acordos(idacordo,anoacordo,dtparcelamento,setordevedor,iddevedor,nroprocessoadm,crcacordante,nomeacordante,cpfcnpj,rginscrestadual,"
'    sql = sql & "cep,endereco,numero,complemento,bairro, cidade,estado,vlrtotal,qtdparcelas,primeirovencimento,vlrtotalhonorarios,qtdparcelashonorarios,"
'    sql = sql & "vlrparcelahonorarios,dtvenctohonorarios,VlrTotalDespesas, QtdParcelasDespesas, VlrParcelaDespesas, DtVenctoDespesas, DtGeracao) values ("
'    sql = sql & Val(lblNumProc.Caption) & "," & Val(lblAnoProc.Caption) & ",'" & Format(lblDataParc.Caption, "mm/dd/yyyy") & "','" & sSetor & "',"
'    sql = sql & nCodReduz & ",'" & txtNumProc.Text & "'," & nCodReduz & ",'" & Mask(Left(sNomeResp, 30)) & "','" & sCPF & "','" & Left(sRG, 20) & "','" & sCep & "','" & sEndImovel & "',"
'    sql = sql & nNumImovel & ",'" & Left(sComplImovel, 40) & "','" & sBairroImovel & "','" & sCidadeEntrega & "','" & sUFEntrega & "'," & Virg2Ponto(Round((nValorParc * nQtdeParc), 2)) & "," & nQtdeParc & ",'"
'    sql = sql & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'," & Virg2Ponto(Round(nValorTotalHon, 2)) & "," & IIf(nValorTotalHon = 0, 0, nQtdeParc) & "," & Virg2Ponto(Round((nValorTotalHon / nQtdeParc), 2)) & ","
'    sql = sql & IIf(nValorTotalHon = 0, "Null", "'" & Format(dDataPrimeiraParc, "mm/dd/yyyy") & "'") & "," & "0,0,0," & "Null" & ",'" & Format(Now, "mm/dd/yyyy") & "')"
'    cnInt.Execute sql, rdExecDirect
'
'   'GRAVA NA TABELA ACORDOSTATUS
'    sql = "insert acordostatus(idacordo,anoacordo,dtocorrencia,ocorrencia,dtgeracao) values("
'    sql = sql & Val(lblNumProc.Caption) & "," & Val(lblAnoProc.Caption) & ",'" & Format(Now, "mm/dd/yyyy") & "','PARCELAMENTO EM DIA','" & Format(Now, "mm/dd/yyyy") & "')"
'    cnInt.Execute sql, rdExecDirect
'
'   'GRAVA OS DÉBITOS DO ACORDO
'    sql = "SELECT DISTINCT origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
'    sql = sql & "origemreparc.numparcela, origemreparc.codcomplemento, SUM(debitotributo.valortributo) AS Total, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza "
'    sql = sql & "FROM origemreparc INNER JOIN debitotributo ON origemreparc.codreduzido = debitotributo.codreduzido AND origemreparc.anoexercicio = debitotributo.anoexercicio AND "
'    sql = sql & "origemreparc.codlancamento = debitotributo.codlancamento AND origemreparc.numsequencia = debitotributo.seqlancamento AND "
'    sql = sql & "origemreparc.numparcela = debitotributo.numparcela AND origemreparc.codcomplemento = debitotributo.codcomplemento INNER JOIN debitoparcela ON origemreparc.codreduzido = debitoparcela.codreduzido AND "
'    sql = sql & "origemreparc.anoexercicio = debitoparcela.anoexercicio AND origemreparc.codlancamento = debitoparcela.codlancamento AND origemreparc.numsequencia = debitoparcela.seqlancamento AND "
'    sql = sql & "origemreparc.NumParcela = debitoparcela.NumParcela And origemreparc.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO GROUP BY origemreparc.numprocesso, origemreparc.codreduzido, origemreparc.anoexercicio, origemreparc.codlancamento, origemreparc.numsequencia,"
'    sql = sql & "origemreparc.NumParcela , origemreparc.CODCOMPLEMENTO, debitoparcela.numerolivro, debitoparcela.paginalivro,debitoparcela.dataajuiza "
'    sql = sql & "HAVING origemreparc.numprocesso = '" & lblNumProc.Caption & "/" & lblAnoProc.Caption & "' AND origemreparc.codreduzido =" & nCodReduz
'    sql = sql & "ORDER BY origemreparc.anoexercicio, origemreparc.numparcela"
'    Set RdoAux2 = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux2
'        Do Until .EOF
'            nPagina = Val(SubNull(!paginalivro))
'            nLivro = Val(SubNull(!numerolivro))
'           'GRAVA NA TABELA ACORDODEBITO
'            sql = "insert acordodebitos(idacordo,anoacordo,nrolivro,nrofolha,seq,lancamento,exercicio,vlroriginal,vlrcorrecao,vlrjuros,vlrmulta,vlrtotal,nroparcela,complparcela,ajuizado,dtgeracao) values("
'            sql = sql & Val(lblNumProc.Caption) & "," & Val(lblAnoProc.Caption) & "," & nLivro & "," & nPagina & ","
'            sql = sql & RdoAux2!numSequencia & "," & RdoAux2!CodLancamento & "," & RdoAux2!AnoExercicio & "," & Virg2Ponto(Format(RdoAux2!Total, "#0.##")) & ",0,0,0," & Virg2Ponto(Format(RdoAux2!Total, "#0.##")) & ","
'            sql = sql & RdoAux2!NumParcela & "," & RdoAux2!CODCOMPLEMENTO & "," & IIf(IsNull(!DATAAJUIZA), 0, 1) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
'            cnInt.Execute sql, rdExecDirect
'           .MoveNext
'        Loop
'       .Close
'    End With
'
'End If
'RdoAux.Close
'
'cnInt.Close 'FECHA CONEXÃO COM A INTEGRATIVA
'*************************

sObs = "Liberação de Carnê Código: " & txtCod.Text & " - " & lblNome.Caption & " Processo: " & txtNumProc.Text & " pelo usuário: " & RetornaUsuarioFullName2(NomeDeLogin)
Sql = "SELECT MAX(SEQ) AS MAXIMO FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & Val(txtCod.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With
'Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USUARIO,DATAOBS,OBS) VALUES(" & Val(txtCod.Text) & "," & nSeq & ",'GTI','" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USERID,DATAOBS,OBS) VALUES(" & Val(txtCod.Text) & "," & nSeq & ",236,'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
cn.Execute Sql, rdExecDirect

If bBoleto Then
    frmReport.ShowReport2 "boletoguia2", frmMdi.hwnd, Me.hwnd, nSid
Else
    frmReport.ShowReport2 "boletoguia_v4", frmMdi.hwnd, Me.hwnd, nSid
End If


'EXIBE RELATORIO
'If optGuia(0).value = True Then
'    If frmMdi.frTeste.Visible = False Then
'        frmReport.ShowReport "Carne2", frmMdi.hwnd, Me.hwnd
'    Else
'        frmReport.ShowReport "CarneTmp", frmMdi.hwnd, Me.hwnd
'    End If
'Else
'    If frmMdi.frTeste.Visible = False Then
 '       frmReport.ShowReport2 "boletoguia2", frmMdi.hwnd, Me.hwnd, nSid
'    Else
'        frmReport.ShowReport2 "boletoguia2Tmp", frmMdi.hwnd, Me.hwnd, nSid'
    'End If
'End If
Liberado

End Sub
