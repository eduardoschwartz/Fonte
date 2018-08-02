VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmArrecadaSN 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrecadação do Simples Nacional"
   ClientHeight    =   915
   ClientLeft      =   6660
   ClientTop       =   4455
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   7725
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   285
      Left            =   135
      TabIndex        =   8
      Top             =   135
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   503
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
   Begin VB.CheckBox chkNaoPago 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Somente não pagos"
      Height          =   195
      Left            =   5730
      TabIndex        =   7
      Top             =   570
      Width           =   1905
   End
   Begin VB.ListBox lstArq 
      Height          =   1425
      Left            =   150
      TabIndex        =   3
      Top             =   4800
      Width           =   13455
   End
   Begin VB.DirListBox Dir 
      Height          =   1215
      Left            =   180
      TabIndex        =   2
      Top             =   3390
      Width           =   1875
   End
   Begin VB.FileListBox File 
      Height          =   1065
      Left            =   3690
      TabIndex        =   1
      Top             =   3420
      Width           =   1485
   End
   Begin VB.DirListBox Dir2 
      Height          =   1215
      Left            =   2160
      TabIndex        =   0
      Top             =   3420
      Width           =   1425
   End
   Begin prjChameleon.chameleonButton cmdBusca 
      Height          =   315
      Left            =   6570
      TabIndex        =   4
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   120
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Gerar"
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
      MICON           =   "frmArrecadaSN.frx":0000
      PICN            =   "frmArrecadaSN.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCorrige 
      Height          =   315
      Left            =   1260
      TabIndex        =   6
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   1410
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Corrige"
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
      MICON           =   "frmArrecadaSN.frx":00BB
      PICN            =   "frmArrecadaSN.frx":00D7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   450
      Width           =   6345
   End
End
Attribute VB_Name = "frmArrecadaSN"
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

Private Sub cmdBusca_Click()
Dim sPath As String, sAno As String, aAno() As String, x As Integer, m As Integer, f As Integer, d As Integer
Dim sDia As String, sMes As String, nMesDe As Integer, nMesAte As Integer, sArquivo As String, nCodBanco As Integer
Dim bDA As Boolean, sRetorno As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String
Dim Posicao  As Long, nTot As Long, nCount As Long, sReg As String, aReg() As Registro, nCodReduz As Long, sRazao As String, sPago As String
Dim sCNPJ As String, bAchou As Boolean, k As Long, sCNPJ2 As String, nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer

ReDim aAno(0): ReDim aReg(0): lstArq.Clear: PBar.value = 0

If MsgBox("Gerar relatório ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

cmdBusca.Enabled = False
Sql = "DELETE FROM ARRECADACAOSN WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

lblMsg.Caption = "Aguarde...Preparando lista de arquivos..."
lblMsg.Refresh
For x = 2007 To Year(Now) + 1
    ReDim Preserve aAno(UBound(aAno) + 1)
    aAno(UBound(aAno)) = x
Next
TriQuickSortString aAno, SortAscending
On Error GoTo proximomes
For x = 1 To UBound(aAno)
    sAno = aAno(x)
    sPath = "\\192.168.200.130\AtualizaGTI\" & sAno
'    sPath = "C:\Trabalho\GTI.NET\GTI.NET\BancoTmp\" & sAno
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
'Exit Sub
lblMsg.Caption = "Carregando Lista..."
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

PBar.value = 0
lblMsg.Caption = "Aguarde...Gravando os dados..."
lblMsg.Refresh

For x = 1 To UBound(aReg)
    If x Mod 50 = 0 Then
        CallPb CLng(x), CLng(UBound(aReg))
    End If
    
    'BUSCA EMPRESA
    Sql = "SELECT CODIGOMOB,RAZAOSOCIAL FROM MOBILIARIO WHERE CNPJ='" & aReg(x).sCNPJ & "' ORDER BY DATAABERTURA DESC"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
        nCodReduz = RdoAux!codigomob
        sRazao = RdoAux!RazaoSocial
    Else
        'BUSCA CIDADAO
        Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & aReg(x).sCNPJ & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux.RowCount > 0 Then
            nCodReduz = RdoAux!CodCidadao
            sRazao = RdoAux!nomecidadao
        Else
            nCodReduz = 0
            sRazao = "CNPJ NÃO LOCALIZADO"
        End If
        RdoAux.Close
    End If
    'RdoAux.Close
        
        
        
        
'    If RetornaArquivo(aReg(x).sArquivo) = "DAF607P9.RET" Then MsgBox "TESTE"
    Sql = "SELECT * from COMPLEMENTOSIMPLES WHERE CODREDUZIDO=" & nCodReduz & " AND ARQUIVOBANCO='" & RetornaArquivo(aReg(x).sArquivo) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
        sPago = "SIM"
    Else
        sPago = "NÃO"
    End If
    RdoAux.Close
        
    With aReg(x)
        Sql = "INSERT ARRECADACAOSN (COMPUTER,CNPJ,COMPETENCIA,VENCIMENTO,ARQUIVO,DATACREDITO,VALOR,CODREDUZIDO,RAZAOSOCIAL,PAGO,CODBANCO) VALUES('"
        Sql = Sql & NomeDeLogin & "','" & .sCNPJ & "','" & .sCompetencia & "','" & Format(.sVencimento, "mm/dd/yyyy") & "','"
        Sql = Sql & RetornaArquivo(.sArquivo) & "','" & Format(.sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(RemovePonto(FormatNumber(.nValor, 2))) & ","
        Sql = Sql & nCodReduz & ",'" & Mask(Left(sRazao, 50)) & "','" & sPago & "'," & .nCodBanco & ")"
        cn.Execute Sql, rdExecDirect
    End With
Next

If chkNaoPago.value = vbChecked Then
    Sql = "DELETE FROM ARRECADACAOSN WHERE COMPUTER='" & NomeDeLogin & "' AND PAGO='SIM'"
    cn.Execute Sql, rdExecDirect
End If

frmReport.ShowReport "ARRECADACAOSN", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ARRECADACAOSN WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

PBar.value = 0
lblMsg.Caption = "Finalizado !!"
lblMsg.Refresh
cmdBusca.Enabled = True
Liberado
Exit Sub
Erro:
 MsgBox Err.Description
Close #1
Liberado

End Sub

Private Sub Form_Load()
Centraliza Me
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

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

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


Private Sub cmdCorrige_Click()
Dim sPath As String, sAno As String, aAno() As String, x As Integer, m As Integer, f As Integer, d As Integer
Dim sDia As String, sMes As String, nMesDe As Integer, nMesAte As Integer, sArquivo As String, nCodBanco As Integer
Dim bDA As Boolean, sRetorno As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String
Dim Posicao  As Long, nTot As Long, nCount As Long, sReg As String, aReg() As Registro, nCodReduz As Long, sRazao As String, sPago As String
Dim sCNPJ As String, bAchou As Boolean, k As Long, sCNPJ2 As String, nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, bValido As Boolean
Dim nAnoCompetencia As Integer, nMesCompetencia As Integer, nNumDoc As Long, nSeqAdd As Integer

Exit Sub
ReDim aAno(0): ReDim aReg(0): lstArq.Clear: PBar.value = 0

If MsgBox("CORRIGIR ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

cmdBusca.Enabled = False

lblMsg.Caption = "Aguarde...Preparando lista de arquivos..."
lblMsg.Refresh

Sql = "DELETE FROM COMPLEMENTOSIMPLES"
cn.Execute Sql, rdExecDirect

For x = 2007 To Year(Now) + 1
    ReDim Preserve aAno(UBound(aAno) + 1)
    aAno(UBound(aAno)) = x
Next
TriQuickSortString aAno, SortAscending
On Error GoTo proximomes
For x = 1 To UBound(aAno)
    sAno = aAno(x)
    sPath = "\\192.168.200.130\AtualizaGTI\" & sAno
'    sPath = "C:\Trabalho\GTI.NET\GTI.NET\BancoTmp\" & sAno
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
lblMsg.Caption = "Carregando Lista..."
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


'GoTo EfetuaBaixa

PBar.value = 0
lblMsg.Caption = "Aguarde...Gravando os dados..."
lblMsg.Refresh

For x = 1 To UBound(aReg)
    If x Mod 50 = 0 Then
        CallPb CLng(x), CLng(UBound(aReg))
    End If
    
    'BUSCA EMPRESA
    Sql = "SELECT CODIGOMOB,RAZAOSOCIAL FROM MOBILIARIO WHERE CNPJ='" & aReg(x).sCNPJ & "' ORDER BY DATAABERTURA DESC"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
        nCodReduz = RdoAux!codigomob
        sRazao = RdoAux!RazaoSocial
    Else
        nCodReduz = 0
        sRazao = "CNPJ NÃO LOCALIZADO"
    End If
    RdoAux.Close
        
    Sql = "SELECT     debitopago.*, debitoparcela.statuslanc FROM  debitopago INNER JOIN  debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND "
    Sql = Sql & "debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
    Sql = Sql & "debitopago.NumParcela = debitoparcela.NumParcela And debitopago.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
    Sql = Sql & "WHERE DEBITOPAGO.CODREDUZIDO=" & nCodReduz & " AND DATARECEBIMENTO='" & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "' and statuslanc<>6"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
        nCodReduz = RdoAux!CODREDUZIDO
        nAno = RdoAux!AnoExercicio
        nLanc = RdoAux!CodLancamento
        nSeq = RdoAux!SeqLancamento
        nParc = RdoAux!NumParcela
        nCompl = RdoAux!CODCOMPLEMENTO
        
        bValido = False
        Sql = "SELECT CODREDUZIDO,NUMDOCUMENTO FROM DEBITOPAGO WHERE CODREDUZIDO=" & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO=" & RdoAux!AnoExercicio & " AND "
        Sql = Sql & "CODLANCAMENTO=" & RdoAux!CodLancamento & " AND SEQLANCAMENTO=" & RdoAux!SeqLancamento & " AND NUMPARCELA=" & RdoAux!NumParcela & " AND CODCOMPLEMENTO=" & RdoAux!CODCOMPLEMENTO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!NumDocumento > 3100000 Then
                bValido = True
            End If
            RdoAux2.Close
        End If
        
        If bValido Then
            Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR,CNPJ,ANO,MES) VALUES(" & nCodReduz & ","
            Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",'" & RetornaArquivo(aReg(x).sArquivo) & "','" & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "',"
            Sql = Sql & Virg2Ponto(CStr(aReg(x).nValor)) & ",'" & aReg(x).sCNPJ & "'," & Val(Left(aReg(x).sCompetencia, 4)) & "," & Val(Right(aReg(x).sCompetencia, 2)) & ")"
            cn.Execute Sql, rdExecDirect
        End If
    End If
    RdoAux.Close
Next

GoTo fim

COMPENSA:
PBar.value = 0
lblMsg.Caption = "Compensando Parcelas..."
lblMsg.Refresh

Sql = "SELECT * FROM COMPLEMENTOSIMPLES ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        CallPb CLng(RdoAux.AbsolutePosition), CLng(RdoAux.RowCount)
        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=6 WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
        Sql = Sql & "NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
        Sql = Sql & "NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
'        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
'        Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & 0 & ",'PARCELA COMPENSADA EM FUNÇÃO DA REATIVAÇÃO DOS DÉBITOS DO SIMPLES NACIONAL','" & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
        Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & 0 & ",'PARCELA COMPENSADA EM FUNÇÃO DA REATIVAÇÃO DOS DÉBITOS DO SIMPLES NACIONAL'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
        
       .MoveNext
    Loop
   .Close
End With

GoTo fim

EfetuaBaixa:
Sql = "DELETE FROM COMPLEMENTOSIMPLES"
cn.Execute Sql, rdExecDirect

lblMsg.Caption = "Efetuando Baixa..."
lblMsg.Refresh

For x = 1 To UBound(aReg)
    CallPb CLng(x), CLng(UBound(aReg))
    nAnoCompetencia = Val(Left(aReg(x).sCompetencia, 4))
    'BUSCA CÓDIGO
    Sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE CONVERT(BIGINT, cnpj) = " & aReg(x).sCNPJ & " ORDER BY DATAABERTURA DESC"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            nCodReduz = !codigomob
        Else
          nCodReduz = 0
'          MsgBox "O CNPJ " & aReg(x).sCNPJ & " NÃO FOI LOCALIZADO.", vbCritical, "Atenção"
        End If
       .Close
    End With
    If nCodReduz > 0 Then
       'O NÚMERO DA PARCELA A SER CRIADA SERÁ O ÚLTIMO NÚMERO DE PARCELA DO ANO
        Sql = "SELECT MAX(NUMPARCELA) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAnoCompetencia & ")"
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
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
'        Sql = Sql & nParc & "," & nCompl & ",2,'" & Format(aReg(x).sVencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,'" & Left$(NomeDeLogin, 25) & "')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
        Sql = Sql & nParc & "," & nCompl & ",2,'" & Format(aReg(x).sVencimento, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect
       'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
        Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
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
        Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & aReg(x).nCodBanco & "," & 0 & "," & Virg2Ponto(aReg(x).nValor) & ",'" & NomeDeLogin & " (ARRECADA SN)" & "')"
        cn.Execute Sql, rdExecDirect
        'CRIAR A PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & nAnoCompetencia & "," & 5 & "," & nSeq & ","
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
        Sql = Sql & nCodReduz & "," & nAnoCompetencia & ",5," & nSeq & "," & nParc & "," & nCompl & "," & nSeqAdd & ",'" & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(aReg(x).nValor) & "," & aReg(x).nCodBanco & "," & 0 & "," & nNumDoc & ","
        Sql = Sql & Virg2Ponto(aReg(x).nValor) & ",0,0" & ",'" & RetornaArquivo(aReg(x).sArquivo) & "'," & 0 & ")"
        cn.Execute Sql, rdExecDirect
       'CRIAR COMPLEMENTOSIMPLES
        Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR) VALUES(" & nCodReduz & ","
        Sql = Sql & nAnoCompetencia & ",5," & nSeq & "," & nParc & "," & nCompl & ",'" & RetornaArquivo(aReg(x).sArquivo) & "','" & Format(aReg(x).sDataCredito, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(aReg(x).nValor)) & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

fim:
lblMsg.Caption = "Finalizado !!"
lblMsg.Refresh
cmdBusca.Enabled = True
Liberado
Exit Sub
Erro:
 MsgBox Err.Description
Close #1
Liberado

End Sub

