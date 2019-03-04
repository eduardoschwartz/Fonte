VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCalculoCIP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Contribuição de Iluminação Pública - CIP"
   ClientHeight    =   3000
   ClientLeft      =   5940
   ClientTop       =   4035
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton btGuias 
      Height          =   345
      Left            =   360
      TabIndex        =   18
      ToolTipText     =   "Gerar Boletos"
      Top             =   2580
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Guias sem registro"
      ENAB            =   0   'False
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmCalculoCIP.frx":0000
      PICN            =   "frmCalculoCIP.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Entrega"
      Height          =   825
      Left            =   3870
      TabIndex        =   15
      Top             =   570
      Width           =   1245
      Begin VB.OptionButton OptEntrega 
         Caption         =   "Balcão"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   510
         Width           =   1035
      End
      Begin VB.OptionButton OptEntrega 
         Caption         =   "Correio"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmCalculoCIP.frx":0176
      Left            =   1590
      List            =   "frmCalculoCIP.frx":0178
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   90
      TabIndex        =   0
      Top             =   540
      Width           =   3705
      Begin VB.Label lblQtdeParcela 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   1950
         TabIndex        =   14
         Top             =   735
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Qtde parcelas..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   13
         Top             =   735
         Width           =   1725
      End
      Begin VB.Label lblValorParcela 
         Caption         =   "R$0.000,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1950
         TabIndex        =   12
         Top             =   450
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Valor parcela..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   450
         Width           =   1725
      End
      Begin VB.Label lblValorTotal 
         Caption         =   "R$0.000.000,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1950
         TabIndex        =   10
         Top             =   1290
         Width           =   1605
      End
      Begin VB.Label Label2 
         Caption         =   "Valor total....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   9
         Top             =   1290
         Width           =   1725
      End
      Begin VB.Label lblQtdeLamina 
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   1950
         TabIndex        =   8
         Top             =   1005
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Qtde lâminas...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   1005
         Width           =   1725
      End
      Begin VB.Label Label2 
         Caption         =   "Qtde terrenos..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label lblQtdeTerreno 
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   1950
         TabIndex        =   1
         Top             =   180
         Width           =   615
      End
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   615
      Left            =   3870
      TabIndex        =   4
      ToolTipText     =   "Gerar Boletos"
      Top             =   1560
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Gerar Boletos"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmCalculoCIP.frx":017A
      PICN            =   "frmCalculoCIP.frx":0196
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   2250
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de Cálculo....:"
      Height          =   225
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   1425
   End
End
Attribute VB_Name = "frmCalculoCIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xImovel As clsImovel
Dim aVencto(12) As String


Private Sub btGuias_Click()
Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, Y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, sDataDam As String, xImovel As clsImovel, sLote As String, sQuadra As String


btGuias.Enabled = False
'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


Sql = "delete from boletoguia_cip where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

sLib = "CIP"


'sNumProc = lblNumProc.Caption & "/" & lblAnoProc.Caption
'dDataProc = lblDataParc.Caption
Sql = "SELECT cadimob.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento, parceladocumento.numparcela, "
Sql = Sql & "parceladocumento.CODCOMPLEMENTO , parceladocumento.NumDocumento, debitoparcela.DataVencimento, debitotributo.ValorTributo FROM cadimob INNER JOIN "
Sql = Sql & "parceladocumento ON cadimob.codreduzido = parceladocumento.codreduzido INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
Sql = Sql & "debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
Sql = Sql & "Where (cadimob.codreduzido in (select codigo from cip_semregistro where ano=2019)) "
'Sql = Sql & "Where (cadimob.li_codbairro =1069) "
Sql = Sql & " And (parceladocumento.AnoExercicio = 2019) And (parceladocumento.CodLancamento =79) And (parceladocumento.SeqLancamento = 0) "
Sql = Sql & "AND  statuslanc=18 ORDER BY cadimob.codreduzido, parceladocumento.numparcela"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    x = 1
   Set xImovel = New clsImovel
    Do Until .EOF
        DoEvents
        nCodReduz = !CODREDUZIDO
        sTipoImposto = "Cont.Ilum.Pub."
        sSetor = "IMOBILIÁRIO"
        xImovel.CarregaImovel nCodReduz
        sNumInsc = xImovel.Inscricao
        sCodReduz = nCodReduz
        sNomeResp = xImovel.NomePropPrincipal
        sQuadra = xImovel.Li_Quadras
        sLote = xImovel.Li_Lotes
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
        sEndImovel = xImovel.Endereco
        nNumImovel = xImovel.Numero
        sComplImovel = xImovel.Complemento
        sBairroImovel = xImovel.Bairro
        If xImovel.Ee_TipoEnd = 0 Then
            xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
        ElseIf xImovel.Ee_TipoEnd = 1 Then
            xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
        ElseIf xImovel.Ee_TipoEnd = 2 Then
            xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
        End If
        sEndEntrega = xImovel.Endereco
        nNumEntrega = Val(xImovel.Numero)
        sComplEntrega = xImovel.Complemento
        sBairroEntrega = xImovel.Bairro
        sCidadeEntrega = xImovel.Cidade
        sUFEntrega = xImovel.UF
        sCepEntrega = xImovel.Cep
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
        
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        sDataDam = Format(!DataVencimento, "dd/mm/yyyy")
        nNumDoc = !NumDocumento
        nValorParc = !ValorTributo

        nNumGuia = nNumDoc

        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)


        
        sValorParc = Format(nValorParc, "#0.00")
        nValorGuia = sValorParc
        nValorDoc = nValorGuia

    sValor = nValorDoc
    dDataVencto = CDate(sDataDam)
    nNumDoc = nNumGuia
    sDadosLanc = "CONTRIBUIÇÃO DE ILUMINAÇÃO PÚBLICA 2019"
    NumBarra2 = Gera2of5Cod(CStr(sValor), CDate(dDataVencto), CLng(nNumDoc), CLng(nCodReduz))
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)

    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
    sBarra = StrBarra2

    If nParc = 1 Then
        sDescImposto = "CIP 2019 - Referente aos meses de Jan, Fev, Mar e Abr."
    ElseIf nParc = 2 Then
        sDescImposto = "CIP 2019 - Referente aos meses de Mai, Jun, Jul e Ago."
    ElseIf nParc = 3 Then
        sDescImposto = "CIP 2019 - Referente aos meses de Set, Out, Nov e Dez."
    End If

    '*******************************************

        Sql = "insert boletoguia_cip(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numbarra2a,numbarra2b,numbarra2c,numbarra2d,endereco2,numimovel2,complemento2,bairro2,cidade2,uf2,cep2) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & "JABOTICABAL" & "','" & "SP" & "','" & Mask(sDescImposto) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 0, nParc) & "," & 3 & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & "Quadra: " & Left(Trim(sQuadra), 15) & " Lote: " & Left(Trim(sLote), 15) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & Left(Mask(sEndEntrega), 80) & "',"
        Sql = Sql & nNumEntrega & ",'" & Left(Mask(sComplEntrega), 30) & "','" & Left(Mask(sBairroEntrega), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & sCepEntrega & "')"
        cn.Execute Sql, rdExecDirect
        
        If nParc = 1 Then
            Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
            Sql = Sql & NomeDeLogin & "'," & x & ",'" & Format(nCodReduz, "000000") & " - CIP 2019','" & Mask(sNomeResp) & "','"
            Sql = Sql & Left(sEndEntrega & " " & nNumEntrega & " " & sComplEntrega, 60) & "','" & sBairroEntrega & " - " & sCidadeEntrega & "','" & sUFEntrega & " - " & sCepEntrega & "')"
            cn.Execute Sql, rdExecDirect
        End If
        
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport2 "BOLETOGUIA_CIP", frmMdi.HWND, Me.HWND, nSid, nNumGuia
If cGetInputState() <> 0 Then DoEvents
frmReport.ShowReport "ETIQUETACIP", frmMdi.HWND, Me.HWND



Liberado


Sql = "delete from boletoguia_cip where sid=" & nSid
 cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
btGuias.Enabled = True
End Sub

Private Sub cmbAno_Click()

Dim Sql As String, RdoAux As rdoResultset, nValorParcela As Double, nAno As Integer, nQtdeParc As Integer, nQtdeTerreno As Integer, nValorTotal As Double, nQtdeLamina As Integer

nAno = Val(cmbAno.Text)
nValorParcela = 0
nQtdeParc = 0
nQtdeTerreno = 0
Ocupado
Sql = "SELECT  paramparcela.codtipo, paramparcela.ano, paramparcela.qtdeparcela, paramparcela.parcelaunica, paramparcela.descontounica, paramparcela.vencunica,"
Sql = Sql & "paramparcela.venc01, paramparcela.venc02, paramparcela.venc03, paramparcela.venc04, paramparcela.venc05, paramparcela.venc06, paramparcela.venc07,"
Sql = Sql & "paramparcela.venc08 , paramparcela.venc09, paramparcela.venc10, paramparcela.venc11, paramparcela.venc12, cip_valor.Valor "
Sql = Sql & "FROM paramparcela INNER JOIN cip_valor ON paramparcela.ano = cip_valor.ano WHERE CODTIPO=6 AND paramparcela.ANO=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        cmdExec.Enabled = False
        lblValorParcela.Caption = FormatNumber(nValorParcela, 2)
        lblQtdeParcela.Caption = Format(nQtdeParc, "00")
        lblQtdeLamina.Caption = Format(nQtdeTerreno, "000000")
        lblValorTotal.Caption = "R$0.000.000,00"
        Liberado
        Exit Sub
    Else
        aVencto(0) = !vencunica
        aVencto(1) = !venc01
        aVencto(2) = SubNull(!venc02)
        aVencto(3) = SubNull(!venc03)
        aVencto(4) = SubNull(!venc04)
        aVencto(5) = SubNull(!venc05)
        aVencto(6) = SubNull(!venc06)
        aVencto(7) = SubNull(!venc07)
        aVencto(8) = SubNull(!venc08)
        aVencto(9) = SubNull(!venc09)
        aVencto(10) = SubNull(!venc10)
        aVencto(11) = SubNull(!venc11)
        aVencto(12) = SubNull(!venc12)
        nValorParcela = !Valor
        nQtdeParc = !qtdeparcela
        cmdExec.Enabled = True
    End If
   .Close
End With

Sql = "select count(*) as contador from vwfullimovel where ativo='S' and codreduzido not in (select codreduzido from areas) AND (imune = 0 OR imune IS NULL)  "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nQtdeTerreno = Val(SubNull(RdoAux!contador))

nQtdeLamina = nQtdeParc * nQtdeTerreno
lblQtdeLamina.Caption = Format(nQtdeParc * nQtdeTerreno, "000000")
lblValorParcela.Caption = FormatNumber(nValorParcela, 2)
lblQtdeParcela.Caption = Format(nQtdeParc, "00")
lblQtdeTerreno.Caption = Format(nQtdeTerreno, "00000")
nValorTotal = nValorParcela * nQtdeLamina
lblValorTotal.Caption = FormatNumber(nValorTotal, 2)

Liberado
End Sub

Private Sub cmdExec_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, nPos As Long, nTot As Long, sNome As String, nTipoEnd As Integer
Dim sNomeLogr As String, sComplemento As String, nNumero As Integer, sBairro As String, sCidade As String, sUF As String, sCep As String
Dim sCPF As String, RdoAux2 As rdoResultset, nNumDoc As Long, Y As Integer, nAno As Integer, nValorParcela As Double, nSid As Long
Dim sTipoImposto As String, nNumGuia As Long, sNumDoc As String, sNumDoc2 As String, sNumDoc3 As String, nNumParcela As Integer, sDataVencto As String
Dim sBarra As String, sDigitavel2 As String
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim nFatorVencto As Long, sDigitavel As String, sNossoNumero As String
nAno = Val(cmbAno.Text)
nValorParcela = CDbl(Replace(lblValorParcela.Caption, "R$", "0"))
nPos = 1
sTipoImposto = "CIP"

'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect


'GoTo continua
Sql = "select * from debitoparcela where anoexercicio=" & nAno & " and codlancamento=79 "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    nNumDoc = 17110332
    Sql = "select * from vwfullimovel where ativo='S' and codreduzido not in (select codreduzido from areas) AND (imune = 0 OR imune IS NULL) order by cpf,cnpj,logradouro,li_num,nomecidadao"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nTot = .RowCount
        Do Until .EOF
            If nPos Mod 20 = 0 Then
                CallPb nPos, nTot
            End If
            nCodReduz = !CODREDUZIDO
            For Y = 1 To Val(lblQtdeParcela.Caption)
                'GRAVA NA TABELA DEBITOPARCELA
                 Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
                 Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodReduz & "," & nAno & ",79," & 0 & "," & Y & ",0,18,'"
                 Sql = Sql & Format(aVencto(Y), "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',1," & 236 & ")"
                 cn.Execute Sql, rdExecDirect
                'GRAVA NA TABELA DEBITO TRIBUTO
                 Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                 Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAno & ",79," & 0 & "," & Y & ",0,669," & Virg2Ponto(CStr(nValorParcela)) & ")"
                 cn.Execute Sql, rdExecDirect
                'GRAVA NA TABELA NUMDOCUMENTO
                 Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor,valorguia) VALUES("
                 Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "',0,0,0," & 0 & ",'" & "GTI/CIP" & "'," & Virg2Ponto(CStr(nValorParcela)) & ")"
                 cn.Execute Sql, rdExecDirect
                'GRAVA NA TABELA PARCELADOCUMENTO
                 Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES("
                 Sql = Sql & nCodReduz & "," & nAno & ",79," & 0 & "," & Y & ",0," & nNumDoc & ")"
                 cn.Execute Sql, rdExecDirect
                 nNumDoc = nNumDoc + 1
            Next
            
            
            nPos = nPos + 1
           .MoveNext
        Loop
       .Close
    End With
End If
Exit Sub



continua:
nPos = 1
Sql = "SELECT debitoparcela.codreduzido, debitoparcela.numparcela, debitoparcela.datavencimento, parceladocumento.numdocumento, vwFULLIMOVEL.nomecidadao,"
Sql = Sql & "vwFULLIMOVEL.CPF , vwFULLIMOVEL.Cnpj, vwFULLIMOVEL.Ee_TipoEnd FROM debitoparcela INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND "
Sql = Sql & "debitoparcela.anoexercicio = parceladocumento.anoexercicio AND debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN vwFULLIMOVEL ON debitoparcela.codreduzido = vwFULLIMOVEL.codreduzido "
Sql = Sql & "Where (debitoparcela.AnoExercicio = 2019) And (debitoparcela.CodLancamento = 79) ORDER BY debitoparcela.codreduzido, debitoparcela.numparcela"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 20 = 0 Then
            CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        sNome = !nomecidadao
        nTipoEnd = !Ee_TipoEnd
        'If nPos > 50 Then GoTo fim
'        If nTipoEnd = 0 And OptEntrega(0).value = True Then
'            GoTo proximo
'        ElseIf nTipoEnd > 0 And OptEntrega(1).value = True Then
'            GoTo proximo
'        End If
        
        nNumParcela = !NumParcela
        sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
        nNumDoc = !NumDocumento
        
        With xImovel
            If nTipoEnd = 0 Then
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
            ElseIf nTipoEnd = 1 Then
                xImovel.RetornaEndereco nCodReduz, Imobiliario, cadastrocidadao
            Else
                xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
            End If

            sNomeLogr = .Endereco
            sComplemento = .Complemento
            nNumero = Val(.Numero)
            sBairro = .Bairro
            sCidade = .Cidade
            sUF = .UF
            sCep = .Cep
        End With

        sCPF = Trim(SubNull(!CPF))
        If Trim(sCPF) = "" Then
           sCPF = Trim(SubNull(!Cnpj))
        End If
        
        '**** GERADOR DE CÓDIGO DE BARRAS ********
        
'        nNumGuia = nNumDoc
'        sNumDoc = CStr(nNumGuia) & "-" & RetornaDVNumDoc(nNumGuia)
'        sNumDoc2 = CStr(nNumGuia) & RetornaDVNumDoc(nNumGuia)
'        sNumDoc3 = CStr(nNumGuia) & Modulo11(nNumGuia)
'        sNossoNumero = "2678478"
'
'        sDigitavel = "001900000"
'        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'        sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
'        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'        sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
'        sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
'        sDigitavel = sDigitavel & sDv
'
'        dDataBase = "07/10/1997"
'        nFatorVencto = CDate(sDataVencto) - CDate(dDataBase)
'        sQuintoGrupo = Format(nFatorVencto, "0000")
'        sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000")
'        sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorParcela, 2)), "0000000000") & "00000026784780"
'        sBarra = sBarra & sNumDoc3 & "18"
'        sDv = Trim(Calculo_DV11(sBarra))
'        sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
'
'        sDigitavel = sDigitavel & sDv & sQuintoGrupo
'
'        sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
'        sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
'        sBarra = Gera2of5Str(sBarra)
    
        '*******************************************

 '       Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
 '       Sql = Sql & "digitavel,codbarra,valorguia,obs,numproc,cep) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nPos & "," & nCodReduz & ",'" & Left(Mask(sNome), 80) & "','" & sCPF & "','"
 '       Sql = Sql & Left(Mask(sNomeLogr), 80) & "'," & nNumero & ",'" & Left(Mask(sComplemento), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Mask(sCidade) & "','" & sUF & "','" & Mask(sTipoImposto) & "','"
 '       Sql = Sql & CStr(nNumGuia) & "'," & nNumParcela & "," & 2 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
 '       Sql = Sql & Virg2Ponto(Format(nValorParcela, "#0.00")) & "," & "'','','" & sCep & "')"
'        cn.Execute Sql, rdExecDirect
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

fim:
'frmReport.ShowReport2 "boletoCIP", frmMdi.hwnd, Me.hwnd, nSid
Liberado
Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Form_Load()
Dim x As Integer

Centraliza Me
Set xImovel = New clsImovel
For x = 2016 To Year(Now) + 1
    cmbAno.AddItem x
Next
cmbAno.Text = Year(Now)

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub


