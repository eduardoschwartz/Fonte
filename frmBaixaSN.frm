VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmBaixaSN 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa por CNPJ - Simples Nacional"
   ClientHeight    =   2625
   ClientLeft      =   2865
   ClientTop       =   2700
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4875
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   2070
      Width           =   4515
      _ExtentX        =   7964
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
   Begin VB.ListBox lstCNPJ 
      Height          =   1425
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   390
      Width           =   2715
   End
   Begin VB.DirListBox Dir2 
      Height          =   1215
      Left            =   5520
      TabIndex        =   3
      Top             =   7770
      Width           =   1425
   End
   Begin VB.FileListBox File 
      Height          =   1065
      Left            =   7050
      TabIndex        =   2
      Top             =   7770
      Width           =   1485
   End
   Begin VB.DirListBox Dir 
      Height          =   1215
      Left            =   3540
      TabIndex        =   1
      Top             =   7740
      Width           =   1875
   End
   Begin VB.ListBox lstArq 
      Height          =   1425
      Left            =   3510
      TabIndex        =   0
      Top             =   9150
      Width           =   13455
   End
   Begin prjChameleon.chameleonButton cmdBusca 
      Height          =   375
      Left            =   3030
      TabIndex        =   5
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   990
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Carregar Lista"
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
      MICON           =   "frmBaixaSN.frx":0000
      PICN            =   "frmBaixaSN.frx":001C
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
      Height          =   375
      Left            =   3030
      TabIndex        =   8
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   1440
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Efetuar Baixa"
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
      MICON           =   "frmBaixaSN.frx":0176
      PICN            =   "frmBaixaSN.frx":0192
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
      Caption         =   "Selecione o CNPJ:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lblMsg 
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
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   2310
      Width           =   4665
   End
End
Attribute VB_Name = "frmBaixaSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Registro
    sCNPJ As String
    sCompetencia As String
    sVencimento As String
    sArquivo As String
    sDataCredito As String
    nValor As String
    nCodBanco As Integer
End Type

Dim aReg() As Registro

Private Sub cmdBaixa_Click()
Dim Sql As String, RdoAux As rdoResultset

If lstCNPJ.ListIndex = -1 Then
    MsgBox "Selecione um CNPJ.", vbCritical, "Atenção"
    Exit Sub
End If

EfetuaBaixa lstCNPJ.Text

End Sub

Private Sub cmdBusca_Click()
lstCNPJ.Clear
cmdBusca.Enabled = False
cmdBaixa.Enabled = False
CarregaCNPJ
cmdBusca.Enabled = True
cmdBaixa.Enabled = True
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CarregaCNPJ()
Dim sPath As String, sAno As String, x As Integer, m As Integer, f As Integer, d As Integer
Dim sDia As String, sMes As String, nMesDe As Integer, nMesAte As Integer, sArquivo As String, nCodBanco As Integer
Dim bDA As Boolean, sRetorno As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String
Dim Posicao  As Long, nTot As Long, nCount As Long, sReg As String, nCodReduz As Long, sRazao As String, sPago As String
Dim sCNPJ As String, bAchou As Boolean, k As Long, sCNPJ2 As String, nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer

lblMsg.Caption = "Preparando Arquivos..."
lblMsg.Refresh

ReDim aReg(0): lstArq.Clear: PBar.value = 0
For x = 2007 To Year(Now)
    sAno = CStr(x)
    sPath = "\\192.168.200.130\AtualizaGTI\" & sAno
    'sPath = "C:\Trabalho\GTI.NET\GTI.NET\BancoTmp\" & sAno
    Dir.Path = sPath
    For m = 0 To Dir.ListCount - 1
        
        sMes = Dir.List(m)
        sPath = Dir.List(m)
        Dir2.Path = sPath
        For d = 0 To Dir2.ListCount - 1
            sPath = Dir2.List(d)
            File.Path = sPath
            For f = 0 To File.ListCount - 1
                If Left(File.List(f), 4) = "DAF6" And UCase(Right(File.List(f), 3)) = "RET" Then
                    lstArq.AddItem sPath & "\" & File.List(f)
                End If
            Next
        Next
proximomes:
    Next
Next

On Error GoTo 0
lblMsg.Caption = "Lendo arquivos bancários..."
lblMsg.Refresh
Ocupado
bAchou = False
nTot = lstArq.ListCount - 1
For x = 0 To lstArq.ListCount - 1
    CallPb CLng(x), CLng(lstArq.ListCount - 1)
    FF1 = FreeFile()
    Open lstArq.List(x) For Binary Access Read Write As FF1
        While Not EOF(FF1)
            If FileLen(lstArq.List(x)) = 0 Then GoTo CloseFile2
            Input #FF1, sReg
            If Left(sReg, 1) = "" Then GoTo CloseFile2
            bExec = False
            
            If sReg = "" Then GoTo CloseFile2
            If Left(sReg, 1) = "9" Then GoTo CloseFile2
            If Left(sReg, 1) = "1" Then
                nCodBanco = Mid(sReg, 76, 3)
                '** TROCA OS BANCOS PELOS BANCOS VIRTUAIS **
                If nCodBanco = 1 Then
                    nCodBanco = 91
                ElseIf nCodBanco = 33 Then
                    nCodBanco = 92
                ElseIf nCodBanco = 237 Then
                    nCodBanco = 93
                ElseIf nCodBanco = 341 Then
                    nCodBanco = 94
                ElseIf nCodBanco = 409 Then
                    nCodBanco = 95
                ElseIf nCodBanco = 151 Then
                    nCodBanco = 96
                ElseIf nCodBanco = 104 Then
                    nCodBanco = 97
                ElseIf nCodBanco = 399 Then
                    nCodBanco = 98
                Else
                    nCodBanco = 91
                End If
            ElseIf Left(sReg, 1) = "2" Then
                ReDim Preserve aReg(UBound(aReg) + 1)
                aReg(UBound(aReg)).sCNPJ = Mid(sReg, 75, 14)
                aReg(UBound(aReg)).sCompetencia = Mid(sReg, 101, 6)
                aReg(UBound(aReg)).sVencimento = ConvDataSerial(Mid(sReg, 18, 8))
                aReg(UBound(aReg)).sArquivo = lstArq.List(x)
                aReg(UBound(aReg)).sDataCredito = ConvDataSerial(Mid(sReg, 10, 8))
                nValorPrincipal = CDbl(Mid(sReg, 107, 17)) / 100
                nValorMulta = CDbl(Mid(sReg, 124, 17)) / 100
                nValorJuros = CDbl(Mid(sReg, 141, 17)) / 100
                aReg(UBound(aReg)).nValor = nValorPrincipal + nValorMulta + nValorJuros
                aReg(UBound(aReg)).nCodBanco = nCodBanco
                sReg = ""
            End If
        Wend
CloseFile2:
    sReg = ""
    Close #FF1
Next x

lblMsg.Caption = "Carregando CNPJ..."
lblMsg.Refresh

For x = 1 To UBound(aReg)
    bAchou = False
    For m = 0 To lstCNPJ.ListCount - 1
        If lstCNPJ.List(m) = aReg(x).sCNPJ Then
            bAchou = True: Exit For
        End If
    Next
    If Not bAchou Then
        lstCNPJ.AddItem aReg(x).sCNPJ
    End If
Next

Liberado
lblMsg.Caption = "Finalizado..."
lblMsg.Refresh

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)

If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Function RetornaArquivo(sArq As String) As String
Dim k As Integer, sChar As String
For k = Len(sArq) To 1 Step -1
    sChar = Mid(sArq, k, 1)
    If sChar = "\" Then
        RetornaArquivo = Right(sArq, Len(sArq) - k)
        Exit Function
    End If
Next

End Function

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Private Sub EfetuaBaixa(sCNPJ As String)
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim Sql As String, RdoAux As rdoResultset, x As Integer, nCount As Integer, nNumDoc As Long

'BUSCA EMPRESA
Sql = "SELECT CODIGOMOB,RAZAOSOCIAL FROM MOBILIARIO WHERE CNPJ='" & sCNPJ & "' ORDER BY DATAABERTURA DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
    nCodReduz = RdoAux!codigomob
Else
    nCodReduz = 0
End If
RdoAux.Close

If nCodReduz = 0 Then
    'BUSCA CIDADAO
    Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & sCNPJ & "' ORDER BY CODCIDADAO DESC"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
        nCodReduz = RdoAux!CodCidadao
    Else
        nCodReduz = 0
    End If
    RdoAux.Close
End If


If nCodReduz = 0 Then
    MsgBox "CNPJ " & sCNPJ & " não cadastrado.", vbCritical, "Atenção"
    Exit Sub
End If

nCount = 0
For x = 1 To UBound(aReg)
    If aReg(x).sCNPJ = sCNPJ Then
        Sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE CODREDUZIDO=" & nCodReduz & " AND ARQUIVOBANCO='" & RetornaArquivo(aReg(x).sArquivo) & "' AND VALOR=" & Virg2Ponto(CStr(aReg(x).nValor))
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        With RdoAux
            If .RowCount = 0 Then
                nAno = Val(Left(aReg(x).sCompetencia, 4))
               'O NÚMERO DA PARCELA A SER CRIADA SERÁ O ÚLTIMO NÚMERO DE PARCELA DO ANO
                Sql = "SELECT MAX(NUMPARCELA) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAno & ")"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                  If IsNull(!maximo) Then
                      nParc = 1
                  Else
                      nParc = !maximo + 1
                  End If
                 .Close
                End With
                nCompl = 0
                nSeq = 0
                'CRIAR PARCELA DE ISS VARIAVEL NESTE MES E ANO COM O VENCIMENTO QUE VEIO DO BANCO
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'                Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
'                Sql = Sql & nParc & "," & nCompl & ",2,'" & Format(aReg(x).sVencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,'" & Left$(NomeDeLogin, 25) & "')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
                Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
                Sql = Sql & nParc & "," & nCompl & ",2,'" & Format(aReg(x).sVencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0," & RetornaUsuarioID(NomeDeLogin) & ")"
                cn.Execute Sql, rdExecDirect
               'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
                Sql = Sql & nParc & "," & nCompl & "," & 13 & "," & Virg2Ponto(aReg(x).nValor) & ")"
                cn.Execute Sql, rdExecDirect
               'CRIAR O DOCUMENTO PARA ELA
                Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                     nNumDoc = !maximo + 1
                    .Close
                End With
                Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,emissor) VALUES(" & nNumDoc & ",'"
                Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & aReg(x).nCodBanco & "," & 0 & "," & Virg2Ponto(aReg(x).nValor) & ",'" & NomeDeLogin & " (BAIXA SN)" & "')"
                cn.Execute Sql, rdExecDirect
                'CRIAR A PARCELADOCUMENTO
                Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
                Sql = Sql & nParc & "," & nCompl & "," & nNumDoc & ")"
                cn.Execute Sql, rdExecDirect
                'ULTIMA SEQ DE PAGTO
                Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND CODCOMPLEMENTO=" & nCompl
                Sql = Sql & " AND NUMPARCELA=" & nParc
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                     If IsNull(!maximo) Then
                        nSeqAdd = 0
                     Else
                        If .RowCount = 0 Then
                           nSeqAdd = 0
                       Else
                          nSeqAdd = !maximo + 1
                       End If
                    End If
                   .Close
                End With
               'CRIAR DEBITOPAGO
                Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQPAG,"
                Sql = Sql & "DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL,INTACTO,VALORTARIFA,ARQUIVOBANCO,VALORDIF) VALUES("
                Sql = Sql & nCodReduz & "," & nAno & ",5," & nSeq & "," & nParc & "," & nCompl & "," & nSeqAdd & ",'" & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "','"
                Sql = Sql & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(aReg(x).nValor) & "," & aReg(x).nCodBanco & "," & 0 & "," & nNumDoc & ","
                Sql = Sql & Virg2Ponto(aReg(x).nValor) & ",0,0" & ",'" & RetornaArquivo(aReg(x).sArquivo) & "'," & 0 & ")"
                cn.Execute Sql, rdExecDirect
               'CRIAR COMPLEMENTOSIMPLES
                Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR) VALUES(" & nCodReduz & ","
                Sql = Sql & nAno & ",5," & nSeq & "," & nParc & "," & nCompl & ",'" & RetornaArquivo(aReg(x).sArquivo) & "','" & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(aReg(x).nValor)) & ")"
                cn.Execute Sql, rdExecDirect
                nCount = nCount + 1
            End If
           .Close
        End With
    
    
        
    End If
Next

If nCount = 0 Then
    MsgBox "Esta empresa não possue lançamentos do Simples Nacional não baixados.", vbExclamation, "Atenção"
Else
    MsgBox "Efetuado baixa em " & CStr(nCount) & " lançamentos.", vbInformation, "Informação"
End If


End Sub





