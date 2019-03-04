VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLiberaCarne 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberação de Carnê de Parcelamento"
   ClientHeight    =   2535
   ClientLeft      =   14790
   ClientTop       =   5130
   ClientWidth     =   4905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2535
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
   Begin VB.Label lblQtdeParc 
      Caption         =   "0"
      Height          =   345
      Left            =   270
      TabIndex        =   18
      Top             =   3270
      Width           =   885
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
Dim bRegistrado As Boolean, nCodReduz As Long, nSeq As Integer, nNumdoc As Long, nNumproc As Long, nAnoproc As Integer

nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
nAnoproc = Right$(txtNumProc.Text, 4)
nCodReduz = Val(txtCod.Text)

Sql = "select * from destinoreparc where anoproc=" & nAnoproc & " and numproc=" & nNumproc & " and codreduzido=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    nSeq = RdoAux!numsequencia
Else
    RdoAux.Close
    
    Sql = "select * from destinoreparc where numprocesso='" & nNumproc & "/" & nAnoproc & "' and codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nSeq = RdoAux!numsequencia
End If
RdoAux.Close

Sql = "select * from parceladocumento where codreduzido=" & nCodReduz & " and codlancamento=20 and seqlancamento=" & nSeq
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumdoc = RdoAux!NumDocumento
RdoAux.Close

Sql = "select * from ficha_compensacao_documento where numero_documento=" & nNumdoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    bRegistrado = True
Else
    bRegistrado = False
End If
RdoAux.Close


'If bFichaCompensacao Then
    'If bRegistrado Then
'If optGuia(0).value = False Then
    EmiteBoleto2
'Else
 '   EmiteBoleto
'End If
'    Else
'        EmiteBoleto
'    End If
'Else
'    EmiteBoleto
'End If

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
            lblNome.Caption = !razaosocial
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
                lblQtdeParc.Caption = "0"
                MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
                txtNumProc.SetFocus
                Exit Sub
            Else
                lblDataParc.Caption = Format(!datareparc, "dd/mm/yyyy")
                nQtdeParc = !qtdeparcela
                lblQtdeParc.Caption = nQtdeParc
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
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumdoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
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
        
        Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nNumdoc = !maximo + 1
        End With
        'GRAVA NA TABELA NUMDOCUMENTO
        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor,valorguia) VALUES("
        Sql = Sql & nNumdoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0,0,'" & NomeDeLogin & " (LIBERAÇÃO CARNÊ)" & "'," & Virg2Ponto(RemovePonto(CStr(nValorParc))) & ")"
        cn.Execute Sql, rdExecDirect
        'GRAVA NA TABELA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
        Sql = Sql & Val(txtCod.Text) & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumdoc & ")"
        cn.Execute Sql, rdExecDirect

        nNumGuia = nNumdoc
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
    
        sValor = nValorDoc
        dDataVencto = CDate(sDataDam)
      '  nNumDoc = Val(sNumDoc2)
        sDadosLanc = sTipoImposto
        NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumdoc, nCodReduz)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
    
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        sBarra = StrBarra2
        
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
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USERID,DATAOBS,OBS) VALUES(" & Val(txtCod.Text) & "," & nSeq & ",236,'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
cn.Execute Sql, rdExecDirect

If bBoleto Then
    frmReport.ShowReport2 "boletoguia2", frmMdi.HWND, Me.HWND, nSid
Else
    frmReport.ShowReport2 "boletoguia_v4", frmMdi.HWND, Me.HWND, nSid
End If


Liberado

End Sub

Private Sub EmiteBoleto2()
Dim nSid As Long, Sql As String, RdoAux As rdoResultset, nCodReduz As Long, sNome As String, sDoc As String, sEndereco As String, sBairro As String, sCidade As String
Dim sUF As String, sCep As String, sDataVencimento As String, nValorGuia As Double, nNumdoc As Long, sObs As String, sDigitavel As String, sBarra As String
Dim sNumInsc As String, sLote As String, sQuadra As String, sEnd As String, nNum As Integer, sCompl As String, RdoAux2 As rdoResultset, sTipoEnd As String, sInsc As String
Dim sNumProc As String, dDataProc As String, dDataBase As String, sNossoNumero As String, nFatorVencto As Long, sQuintoGrupo As String, sNumDoc As String
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim sCampo1 As String, sCampo2 As String, sCampo3 As String, sCampo4 As String, sCampo5 As String, sDigitavel2 As String, bRegistrado As Boolean


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

'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

Sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute Sql, rdExecDirect

nCodReduz = Val(txtCod.Text)


'ENDEREÇO DO CONTRIBUINTE
Select Case nCodReduz
    Case 1 To 99999
        Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !Inscricao
            sNome = !nomecidadao
            sDoc = Format(SubNull(!CPF), "00000000000")
            If sDoc = "" Then
                sDoc = Format(SubNull(!Cnpj), "00000000000000")
            End If
            sEnd = SubNull(!Logradouro)
            nNum = Val(SubNull(!Li_Num))
            sCompl = Left(SubNull(!Li_Compl), 30)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!li_uf)
            sQuadras = Left(SubNull(!Li_Quadras), 15)
            sLotes = Left(SubNull(!Li_Lotes), 10)
            sCep = CStr(RetornaCEP(!CodLogr, !Li_Num))
           .Close
        End With
    Case 100000 To 300000
        Sql = "SELECT * FROM vwFULLEMPRESA3 WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = SubNull(!INSCESTADUAL)
            sNome = !razaosocial
            sDoc = SubNull(!CPF)
            If Val(sDoc) = 0 Then
                sDoc = SubNull(!Cnpj)
            End If
            sEnd = !Logradouro
            nNum = !Numero
            sCompl = SubNull(!Complemento)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!SiglaUF)
            sQuadras = ""
            sLotes = ""
            If !CodCidade = 413 Then
                sCep = CStr(RetornaCEP(!CodLogradouro, !Numero))
            Else
                sCep = SubNull(!Cep)
            End If
         End With
     Case 500000 To 800000
        sTipoEnd = "R"
        Sql = "select * from cidadao where codcidadao=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If SubNull(RdoAux2!etiqueta) = "N" And SubNull(RdoAux2!etiqueta2) = "S" Then
                sTipoEnd = "C"
            End If
            RdoAux2.Close
        End If
        
        If sTipoEnd = "R" Then
            Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
            Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
            Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
            Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        Else
            Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
            Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
            Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
            Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        On Error Resume Next
        With RdoAux2
            If .RowCount > 0 Then
                 sNome = !nomecidadao
                 If Val(SubNull(!FCodLogradouro)) > 0 Then
                     Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                     Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                     Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                     Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                     With RdoS
                         If .RowCount > 0 Then
                            sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                         Else
                            sEnd = ""
                         End If
                        .Close
                     End With
                 Else
                    sEnd = SubNull(!FNomeLogradouro)
                 End If
                 nNum = Val(SubNull(RdoAux2!fNUMIMOVEL))
                  
                 Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                 Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                 If RdoS.RowCount > 0 Then
                     sCidade = RdoS!descCidade
                     sUF = SubNull(!fsiglauf)
                 Else
                      sCidade = ""
                      sUF = ""
                 End If
                 If Not IsNull(!CodBairro) Then
                     Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                     If .RowCount > 0 Then
                         sBairro = RdoS!DescBairro
                     Else
                         sBairro = ""
                     End If
                 Else
                     sBairro = ""
                 End If

                 sCompl = SubNull(!fcomplemento)
            Else
                sEnd = ""
                sBairro = ""
                sCidade = ""
                sUF = ""
                sCompl = ""
            End If
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
            End If
            
            If !fCodCidade = 413 Then
                sCep = CStr(RetornaCEP(!FCodLogradouro, !fNUMIMOVEL))
            Else
                sCep = SubNull(!FCEP)
            End If
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
        sDataVencimento = Format(!DataVencimento, "dd/mm/yyyy")
        
        Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO <> 3"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nValorGuia = FormatNumber(!soma, 2)
           .Close
        End With
        
        Sql = "select min(numdocumento) as MINIMO from parceladocumento where CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND "
        Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If Not IsNull(RdoAux2!minimo) Then
            nNumdoc = RdoAux2!minimo
        Else
            RdoAux2.Close
            Sql = "select max(numdocumento) as maximo from numdocumento"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nNumdoc = RdoAux2!maximo + 1
            Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & !CODREDUZIDO & "," & nAno & "," & nLanc & "," & nSeq & ","
            Sql = Sql & nParc & "," & nCompl & "," & nNumdoc & ")"
            cn.Execute Sql, rdExecDirect
        End If
        RdoAux2.Close
        
        
        Sql = "select * from ficha_compensacao_documento where numero_documento=" & nNumdoc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            bRegistrado = True
        Else
            bRegistrado = False
        End If
        RdoAux2.Close
        sEndereco = Left(sEnd, 80) & ", " & nNum & " " & Left(sCompl, 30)
        If Not bRegistrado Then
            Sql = "insert ficha_compensacao_documento(numero_documento,data_vencimento,valor_documento,nome,cpf,endereco,bairro,cep,cidade,uf) values(" & nNumdoc & ",'"
            Sql = Sql & Format(sDataVencimento, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorGuia)) & ",'" & Mask(Left(sNome, 40)) & "','" & RetornaNumero(sDoc) & "','"
            Sql = Sql & Mask(Left(sEndereco, 40)) & "','" & Mask(Left(sBairro, 15)) & "','" & RetornaNumero(sCep) & "','" & Mask(Left(sCidade, 30)) & "','" & sUF & "')"
            cn.Execute Sql, rdExecDirect
        End If
        
        sNossoNumero = "2873532"
        dDataBase = "07/10/1997"
        nFatorVencto = CDate(sDataVencimento) - CDate(dDataBase)
        sQuintoGrupo = Format(nFatorVencto, "0000")
        sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorGuia, 2)), "0000000000")
        sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorGuia, 2)), "0000000000") & "000000287353200"
        sBarra = sBarra & CStr(nNumdoc) & "17"

        sCampo1 = "0019" & Mid(sBarra, 20, 5)
        sDigitavel = sCampo1 & Val(Calculo_DV10(sCampo1))
        sCampo2 = Mid(sBarra, 24, 10)
        sDigitavel = sDigitavel & sCampo2 & Val(Calculo_DV10(sCampo2))
        sCampo3 = Mid(sBarra, 34, 10)
        sDigitavel = sDigitavel & sCampo3 & Val(Calculo_DV10(sCampo3))
        sCampo5 = Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorGuia, 2)), "0000000000")
        sCampo4 = Val(Calculo_DV11(sBarra))
        sDigitavel = sDigitavel & sCampo4 & sCampo5
        sBarra = Left(sBarra, 4) & sCampo4 & Mid(sBarra, 5, Len(sBarra) - 4)
        sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
        sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
        sBarra = Gera2of5Str(sBarra)

        
        sObs = "Referente ao Parcelamento de nº: " & sNumProc & " - Código do contribuinte: " & Format(nCodReduz, "000000")
        sNumDoc = "287353200" & Format(nNumdoc, "00000000")
        Sql = "Insert FICHA_COMPENSACAO(SID,SEQ,CODIGO,NOME,CPF,ENDERECO,BAIRRO,CIDADE,CEP,DOCUMENTO,VALOR,VENCIMENTO,PARCELA,DIGITAVEL,CODBARRA,OBS,INSCRICAO,QUADRA,LOTE,UF) VALUES("
        Sql = Sql & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & sDoc & "','" & sEndereco & "','" & Left(Mask(sBairro), 25) & "','"
        Sql = Sql & Mask(sCidade) & "','" & sCep & "'," & sNumDoc & "," & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & Format(sDataVencimento, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(nParc, "00") & "/" & Format(Val(lblQtdeParc.Caption), "00") & "','" & sDigitavel2 & "','" & Mask(sBarra) & "','" & sObs & "','" & sInsc & "','"
        Sql = Sql & Mask(sQuadra) & "','" & Mask(sLote) & "','" & sUF & "')"
        cn.Execute Sql, rdExecDirect

        x = x + 1
       .MoveNext
    Loop
   .Close
End With


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
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USERID,DATAOBS,OBS) VALUES(" & Val(txtCod.Text) & "," & nSeq & ",236,'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sObs) & "')"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport3 "FICHACOMPENSACAO", frmMdi.HWND, Me.HWND, nSid

Sql = "delete from ficha_compensacao where sid=" & nSid
cn.Execute Sql, rdExecDirect

Liberado

End Sub

