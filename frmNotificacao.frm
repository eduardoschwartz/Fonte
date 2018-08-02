VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmNotificacao 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificação de imposto devido"
   ClientHeight    =   1725
   ClientLeft      =   4680
   ClientTop       =   3675
   ClientWidth     =   5100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   5100
   Begin VB.CheckBox chkCalc 
      Caption         =   "Calculo"
      Enabled         =   0   'False
      Height          =   195
      Left            =   225
      TabIndex        =   10
      Top             =   1845
      Width           =   1905
   End
   Begin VB.CheckBox chkEtiq 
      Caption         =   "Gerar etiquetas"
      Height          =   195
      Left            =   3015
      TabIndex        =   9
      Top             =   765
      Width           =   1905
   End
   Begin VB.OptionButton Opt 
      Caption         =   "ISS"
      Height          =   240
      Index           =   1
      Left            =   1620
      TabIndex        =   8
      Top             =   720
      Width           =   960
   End
   Begin VB.OptionButton Opt 
      Caption         =   "IPTU"
      Height          =   240
      Index           =   0
      Left            =   495
      TabIndex        =   7
      Top             =   720
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.TextBox txtSeq 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3870
      TabIndex        =   1
      Top             =   180
      Width           =   1035
   End
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   240
      Left            =   225
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2895
      _ExtentX        =   5106
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
      Color           =   12500670
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1530
      TabIndex        =   0
      Top             =   165
      Width           =   945
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Index           =   1
      Left            =   495
      TabIndex        =   3
      ToolTipText     =   "Imprimir Detalhe"
      Top             =   2430
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Etiquetas"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmNotificacao.frx":0000
      PICN            =   "frmNotificacao.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Index           =   0
      Left            =   3510
      TabIndex        =   2
      ToolTipText     =   "Imprimir as cartas"
      Top             =   1305
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Cartas"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmNotificacao.frx":00CD
      PICN            =   "frmNotificacao.frx":00E9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seq Inicial..:"
      Height          =   225
      Index           =   1
      Left            =   2790
      TabIndex        =   6
      Top             =   225
      Width           =   990
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercício............:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   1395
   End
End
Attribute VB_Name = "frmNotificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TIPOS
Private Type PROFUNDIDADE
    Distrito As Integer
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORPROFUN
    Distrito As Integer
    Codigo As Integer
    Fator As Double
End Type
Private Type GLEBA
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORCATEG
    Uso As Integer
    Tipo As Integer
    Categoria As Integer
    Fator As Double
End Type
'MATRIZES
Dim aFatorD() As Double
Dim aFatorP() As Double
Dim aFatorT() As Double
Dim aFatorS() As Double
Dim aFatorG() As Double
Dim aFatorR() As Double
Dim aProf() As PROFUNDIDADE
Dim aFatorF() As FATORPROFUN
Dim aFatorC() As FATORCATEG
Dim aGleba() As GLEBA

Dim nAreaTerreno As Double, nAreaConstruida As Double, nTestadaPrincipal As Double, nVVT As Double, nVVC As Double, nVVI As Double
Dim nValorIptu As Double, nValorITU As Double, nAnoCalculo As Integer, nValorFinal As Double

Private Sub Form_Load()
If NomeDeLogin <> "SCHWARTZ" Then
    chkCalc.Enabled = False
End If
Centraliza Me
End Sub

Private Sub cmdPrint_Click(Index As Integer)
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, nAreaOld As Double, nCodReduz As Integer, nDif As Integer, nArea As Double
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, sExtenso As String, nCodCidadao As Long
Dim sBairro As String, sCidade As String, sEndereco2 As String, sCep As String, nSeq As Integer, sProcesso As String, sInsc As String, nPerc As Integer
Dim nTipoEnd As Integer, sLogradouro As String, sBAIRRO2 As String, sCEP2 As String, nNumero2 As Integer, nValorIss As Double, nSeq2 As Integer
Dim sTipo As String, nUso As Integer, nCateg As Integer, nCodTributo As Integer, sDataBase As String, sDataVencto As String, sHist As String
Dim nValorPago As Double

sDataBase = Format(Now, "dd/mm/yyyy")
sDataVencto = "15/07/2011"

If Val(txtAno.Text) < 2004 Or Val(txtAno.Text) > Year(Now) Then
    MsgBox "Ano inválido.", vbExclamation, "Atenção"
    Exit Sub
End If
If Val(txtSeq.Text) = 0 Then
    MsgBox "Digite a sequencia inicial.", vbExclamation, "Atenção"
    Exit Sub
End If

nAnoCalculo = Val(txtAno.Text)
LoadMatrix

If Opt(0).value = True Then
    Sql = "DELETE FROM NOTIFICACAO WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect

    Ocupado
    If cGetInputState() <> 0 Then DoEvents
    
    nSeq = Val(txtSeq.Text)
    
    Sql = "select * from vwfullimovel2 inner join laseriptu on (vwfullimovel2.codreduzido=laseriptu.codreduzido) INNER JOIN "
    Sql = Sql & "notificacaoproc ON vwFULLIMOVEL2.codreduzido = notificacaoproc.codigo where ano=" & Val(txtAno.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            If .AbsolutePosition Mod 10 = 0 Then
                CallPb .AbsolutePosition, .RowCount
            End If
            
            nCodReduz = !CODREDUZIDO
            
            Sql = "select sum(valorpagoreal) as somapago from debitopago where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & Val(txtAno.Text)
            Sql = Sql & " and codlancamento=1"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux2!somapago) Then
                nValorPago = 0
            Else
                nValorPago = RdoAux2!somapago
            End If
            RdoAux2.Close
            
            sProcesso = !Processo
            nPerc = !Perc
            
            CalculoIndividual (nCodReduz)
            nDif = nValorFinal * nPerc / 100
            
'            If nCodReduz = 4310 Then MsgBox "teste"
            
            sEndereco = !Logradouro
            sLogradouro = sEndereco
            nNumero = !Li_Num
            nNumero2 = nNumero
            sBairro = !DescBairro
            sCep = RetornaCEP(!CodLogr, !Li_Num)
            sEndereco = sEndereco & " Nº " & nNumero & ", " & sBairro & " " & sCep
            sEndereco2 = ""
            sInsc = !Inscricao
'            nDif = Format(nValorFinal - (!VALORTOTALPARC * 12), "#0.00")
            sExtenso = Extenso(nDif)
            
                    'ENDEREÇO DE ENTREGA
            nTipoEnd = !Ee_TipoEnd
            If nTipoEnd = 0 Then 'Endereço do imóvel
                sEndereco2 = sEndereco
                sBAIRRO2 = sBairro
                sCidade = "JABOTICABAL"
                sUF = "SP"
                sCEP2 = RetornaCEP(!CodLogr, nNumero)
            ElseIf nTipoEnd = 1 Then 'Endereço do prop
                If Not IsNull(!NomeLogradouro) And SubNull(!NomeLogradouro) <> "" Then
                    sLogradouro = !NomeLogradouro
                Else
                    sLogradouro = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2
                End If
                nNumero2 = Val(SubNull(!NUMIMOVEL))
                sBAIRRO2 = SubNull(!DESCBAIRROP)
                sCidade = SubNull(!descCidade)
                sUF = SubNull(!SiglaUF)
                sCEP2 = RetornaCEP(Val(SubNull(!CodLogradouro)), nNumero)
                sEndereco2 = sLogradouro & " Nº " & nNumero2 & ", " & sBAIRRO2 & " " & sCEP2
            ElseIf nTipoEnd = 2 Then 'Endereço de entrega
                If (!Ee_CodLog) = 0 Then
                    sLogradouro = SubNull(!Ee_NomeLog)
                Else
                    sLogradouro = Trim$(SubNull(!AbrevTipoLogEE)) & " " & Trim$(SubNull(!AbrevTitLogEE)) & " " & !Ee_NomeLog
                End If
                nNumero2 = Val(SubNull(!Ee_NumImovel))
                sBAIRRO2 = SubNull(!BairroEE)
                sCidade = IIf(IsNull(!CidadeEE), "JABOTICABAL", !CidadeEE)
                sUF = IIf(IsNull(!Ee_Uf), "SP", !Ee_Uf)
                If Not IsNull(!Ee_Cep) Then
                    sCEP2 = !Ee_Cep
                Else
                    sCEP2 = RetornaCEP(Val(SubNull(!Ee_CodLog)), Val(SubNull(!Ee_NumImovel)))
                End If
                sEndereco2 = sLogradouro & " Nº " & nNumero2 & ", " & sBAIRRO2 & " " & sCEP2
            End If
    
            Sql = "insert notificacao (usuario,codreduz,processo,numseq,nome,end1,end2,inscricao,ano,areat,areac,testadap,vvt,vvc,vvi,iptunovo,iptupago,dif,valorext,perc) "
            Sql = Sql & "values('"
            Sql = Sql & NomeDeLogin & "'," & nCodReduz & ",'" & sProcesso & "'," & nSeq & ",'" & Mask(!nomecidadao) & "','" & Left(Mask(sEndereco), 70) & "','" & Left(Mask(sEndereco2), 70) & "','"
            Sql = Sql & sInsc & "'," & Val(txtAno.Text) & "," & Virg2Ponto(!AreaTerreno) & "," & Virg2Ponto(CStr(nAreaConstruida)) & "," & Virg2Ponto(!TESTADAPRINC) & ","
            'Sql = Sql & Virg2Ponto(CStr(nVVT)) & "," & Virg2Ponto(CStr(nVVC)) & "," & Virg2Ponto(CStr(nVVI)) & "," & Virg2Ponto(Format(nValorFinal, "#0.00")) & "," & Virg2Ponto(!VALORTOTALPARC * 12) & ","
            Sql = Sql & Virg2Ponto(CStr(nVVT)) & "," & Virg2Ponto(CStr(nVVC)) & "," & Virg2Ponto(CStr(nVVI)) & "," & Virg2Ponto(Format(nValorFinal, "#0.00")) & "," & Virg2Ponto(Format(nValorPago, "#0.00")) & ","
            Sql = Sql & nDif & ",'" & sExtenso & "'," & nPerc & ")"
            cn.Execute Sql, rdExecDirect
            
            If chkEtiq.value = vbChecked Then
                Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
                Sql = Sql & NomeDeLogin & "'," & nSeq & ",'" & "" & "','" & Mask(!nomecidadao) & "','"
                Sql = Sql & Left(sLogradouro, 55) & "," & nNumero2 & "','" & sCEP2 & "   " & sBAIRRO2 & "','" & sCidade & "   " & sUF & "')"
                cn.Execute Sql, rdExecDirect
            End If
            
            nSeq = nSeq + 1
           .MoveNext
        Loop
       .Close
    End With
    
    Liberado
    
    frmReport.ShowReport2 "NOTIFICACAO2", frmMdi.hwnd, Me.hwnd
    
    If chkEtiq.value = vbChecked Then
        frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd
    End If
    
    Sql = "DELETE FROM NOTIFICACAO WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
Else
    Sql = "DELETE FROM NOTIFICACAOISS WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
    
    Ocupado
    If cGetInputState() <> 0 Then DoEvents
    
    nSeq = Val(txtSeq.Text)
    
    Sql = "select * from vwfullimovel2 INNER JOIN notificacaoproc2 ON vwFULLIMOVEL2.codreduzido = notificacaoproc2.codigo order by nomecidadao"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            If .AbsolutePosition Mod 10 = 0 Then
                CallPb .AbsolutePosition, .RowCount
            End If
            sNome = !nomecidadao
            nCodReduz = !CODREDUZIDO
'            If nCodReduz = 1576 Then MsgBox "teste"
            nCodCidadao = !CodCidadao
            sProcesso = !Processo
            sEndereco = !Logradouro
            sEndereco2 = sEndereco
            nNumero = !Li_Num
            sBairro = !DescBairro
            sCep = RetornaCEP(!CodLogr, !Li_Num)
            sEndereco = sEndereco & " Nº " & nNumero & ", " & sBairro & " " & sCep
            nSeq = !NUMNOT
            nDif = Format(123.67, "#0.00")
            sExtenso = Extenso(nDif)
            nArea = !Area
            nUso = !Uso
            nCateg = !CATEG
            If nUso = 1 Then
                sTipo = "Residencial "
            ElseIf nUso = 2 Then
                sTipo = "Industrial "
            ElseIf nUso = 3 Then
                sTipo = "Comercial "
            End If
                       
            If nUso <> 2 Then
                If nCateg = 1 Then
                    sTipo = sTipo & " Médio"
                ElseIf nCateg = 3 Then
                    sTipo = sTipo & " Baixo"
                End If
            End If
            
            If Not IsNull(!CODCID) Then
                nCodCidadao = !CODCID
                Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                sNome = RdoAux2!nomecidadao
                RdoAux2.Close
            End If
            
'            Sql = "SELECT areas.codreduzido,areas.areaconstr, usoconstr.descusoconstr,usoconstr.codusoconstr,categconstr.codcategconstr, categconstr.desccategconstr FROM areas INNER JOIN "
'            Sql = Sql & "usoconstr ON areas.usoconstr = usoconstr.codusoconstr INNER JOIN categconstr ON areas.catconstr = categconstr.codcategconstr "
'            Sql = Sql & "WHERE (areas.codreduzido = " & nCodReduz & ") AND (areas.tipoarea = 'P')"
'            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'            If RdoAux2.RowCount = 0 Then GoTo PROXIMO
'            sTipo = RdoAux2!descusoconstr & " " & RdoAux2!desccategconstr
'            nArea = RdoAux2!areaconstr
'            nUso = RdoAux2!codusoconstr
'            nCateg = RdoAux2!codcategconstr
'            RdoAux2.Close
            
            If nUso = 1 Then 'residencial
                If nCateg = 3 Or nCateg = 4 Or nCateg = 7 Or nCateg = 8 Then  'baixo
                    nValorIss = 3.93
                    nCodTributo = 179
                ElseIf nCateg = 1 Or nCateg = 2 Then 'medio
                    nValorIss = 5.44
                    nCodTributo = 180
                ElseIf nCateg = 5 Or nCateg = 6 Then 'alto
                    nCodTributo = 181
                    nValorIss = 6.97
                End If
            ElseIf nUso = 2 Then 'comercial
                If nCateg = 3 Or nCateg = 4 Or nCateg = 7 Or nCateg = 8 Then  'baixo
                    nCodTributo = 182
                    nValorIss = 4.52
                ElseIf nCateg = 1 Or nCateg = 2 Then 'medio
                    nCodTributo = 183
                    nValorIss = 6.04
                ElseIf nCateg = 5 Or nCateg = 6 Then 'alto
                    nCodTributo = 184
                    nValorIss = 6.66
                End If
            ElseIf nUso = 3 Then 'industrial
                nCodTributo = 185
                nValorIss = 4.52
            End If
            
            nValorFinal = nArea * nValorIss
                
            Sql = "insert notificacaoiss(usuario,codigo,razao,processo,seq,ano,endereco,tipo,area,valoriss,valorpago,codcidadao) "
            Sql = Sql & "values('"
            Sql = Sql & NomeDeLogin & "'," & nCodReduz & ",'" & Mask(sNome) & "','" & sProcesso & "'," & nSeq & "," & Val(txtAno.Text) & ",'" & Left(Mask(sEndereco), 70) & "','" & sTipo & "',"
            Sql = Sql & Virg2Ponto(CStr(nArea)) & "," & Virg2Ponto(CStr(nValorIss)) & "," & Virg2Ponto(CStr(nValorFinal)) & "," & nCodCidadao & ")"
            cn.Execute Sql, rdExecDirect

            If chkEtiq.value = vbChecked Then
                Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
                Sql = Sql & NomeDeLogin & "'," & nSeq & ",'" & "Notificação nº " & nSeq & "/2011" & "','" & Mask(sNome) & "','"
                Sql = Sql & sEndereco2 & "," & nNumero & "','" & sCep & "   " & sBairro & "','" & "JABOTICABAL" & "   " & "SP" & "')"
                cn.Execute Sql, rdExecDirect
            End If

            'insere parcelas de calculo
            If chkCalc.value = vbChecked Then
                GoTo proximo
                Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodCidadao & " AND ANOEXERCICIO=2011 AND CODLANCAMENTO=65"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If IsNull(!maximo) Then
                        nSeq2 = 1
                    Else
                        nSeq2 = !maximo + 1
                    End If
                   .Close
                End With
                
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USUARIO) VALUES("
'                Sql = Sql & nCodCidadao & "," & 2011 & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "','" & Left$("GTI", 25) & "')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
                Sql = Sql & nCodCidadao & "," & 2011 & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "'," & 236 & ")"
                cn.Execute Sql, rdExecDirect
                
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & nCodCidadao & "," & 2011 & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & nCodTributo & "," & Virg2Ponto(Format(nValorFinal, "#0.00")) & ")"
                cn.Execute Sql, rdExecDirect
                
                sHist = "Iss construção civil processo nº " & sProcesso & " notificação nº " & nSeq
                
'                Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES("
'                Sql = Sql & nCodCidadao & "," & 2011 & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 0 & ",'" & Mask(sHist) & "','"
'                Sql = Sql & "GTI" & "','" & Format(Now, "mm/dd/yyyy") & "')"
                Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES("
                Sql = Sql & nCodCidadao & "," & 2011 & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 0 & ",'" & Mask(sHist) & "',"
                Sql = Sql & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                cn.Execute Sql, rdExecDirect
                
                sHist = "Iss construção civil lançado no código " & nCodCidadao & " processo nº " & sProcesso & " notificação nº " & nSeq
                
                Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If IsNull(!maximo) Then
                        nSeq2 = 1
                    Else
                        nSeq2 = !maximo + 1
                    End If
                   .Close
                End With

'                Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'                Sql = Sql & nCodReduz & "," & nSeq2 & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "','" & "GTI" & "','" & Format(Now, "mm/dd/yyyy") & "')"
                Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
                Sql = Sql & nCodReduz & "," & nSeq2 & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
                cn.Execute Sql, rdExecDirect
            End If

proximo:
           .MoveNext
        Loop
       .Close
    End With
    
    Liberado
    
    frmReport.ShowReport2 "NOTIFICACAO3", frmMdi.hwnd, Me.hwnd
    If chkEtiq.value = vbChecked Then
        frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd
    End If
    
    Sql = "DELETE FROM NOTIFICACAOISS WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
    Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
End If

Me.MousePointer = vbDefault
Pb.value = 0
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub CalculoIndividual(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double, bTemPredial As Boolean
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer, nValorVenalTerritorial As Double, nValorVenalPredial As Double

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where (CADIMOB.CODREDUZIDO = " & nCodReduz & ") GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "
nValorIptu = 0: nValorITU = 0
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
    
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    nAreaConstruida = !soma
                Else
                    nAreaConstruida = 0
                End If
               '.Close
            End With
        End If
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
        End If
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
    For x = 1 To UBound(aGleba)
        If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
             Exit For
        ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
             Exit For
        End If
    Next
    nCodGleba = aGleba(x).Codigo
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
    nFatorGleba = aFatorG(nCodGleba)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    On Error Resume Next
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= aProf(x).Min And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
     Else
        nFatorProfundidade = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = FormatNumber(nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba, 2)
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        nValorFatores = nValorFatores * nFatorDistrito
        nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
        nVVT = nValorVenalTerritorial
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        nFatorCategoria = 0
        For x = 1 To UBound(aFatorC)
            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
               nFatorCategoria = aFatorC(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
        
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
    Else
        nFatorDistrito = 0
        nFatorCategoria = 0
    End If
    nVVC = nValorVenalPredial
    nVVI = nVVC + nVVT
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorIptu = nValorVenalImovel * (1.5 / 100) 'reajuste 2004-2005 (TIRADO)
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorITU = nValorVenalImovel * (3 / 100)  'reajuste 2004-2005 (TIRADO)
    End If
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    If bTemPredial Then
        nValorFinal = nValorIptu
    Else
        nValorFinal = nValorITU
    End If
End With

End Sub


Private Sub LoadMatrix()

ReDim aFatorD(3)
ReDim aFatorD98(3)
ReDim aFatorP(6)
ReDim aFatorP98(6)
ReDim aFatorT(6)
ReDim aFatorT98(6)
ReDim aFatorS(6)
ReDim aFatorS98(6)
ReDim aFatorG(23)
ReDim aFatorG98(23)
ReDim aFatorR(8)
ReDim aFatorR98(8)

Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .Close
End With

ReDim aProf(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aProf(UBound(aProf) + 1)
        aProf(UBound(aProf)).Distrito = !CODDISTRITO
        aProf(UBound(aProf)).Codigo = !CODPROFUN
        aProf(UBound(aProf)).Min = !MINPROFUN
        aProf(UBound(aProf)).Max = !MAXPROFUN
       .MoveNext
     Loop
    .Close
End With


ReDim aFatorF(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aGleba(0)
Sql = "SELECT CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY CODGLEBA "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aGleba(UBound(aGleba) + 1)
        aGleba(UBound(aGleba)).Codigo = !CODGLEBA
        aGleba(UBound(aGleba)).Min = !MINGLEBA
        aGleba(UBound(aGleba)).Max = !MAXGLEBA
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorC(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !coduso
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !codcateg
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub


Private Sub lblEsc_Click()

End Sub
