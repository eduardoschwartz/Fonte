VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmArqBanco 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Arquivos dos Bancos"
   ClientHeight    =   4665
   ClientLeft      =   2625
   ClientTop       =   2385
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstLog 
      Height          =   1230
      Left            =   45
      TabIndex        =   6
      Top             =   2955
      Width           =   6285
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   2715
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   465
      Width           =   3630
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   30
      TabIndex        =   2
      Top             =   780
      Width           =   2625
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   2640
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5265
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   4260
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
      MICON           =   "frmArqBanco.frx":0000
      PICN            =   "frmArqBanco.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   315
      Left            =   4185
      TabIndex        =   5
      ToolTipText     =   "Incluir Arquivos"
      Top             =   4275
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Incluir"
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
      MICON           =   "frmArqBanco.frx":008A
      PICN            =   "frmArqBanco.frx":00A6
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
      Caption         =   "Selecione o(s) Arquivo(s) Bancários a serem incluidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   5820
   End
End
Attribute VB_Name = "frmArqBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FebrabanA
    CodigoRegistro As String * 1
    CodigoRemessa As String * 1
    CodigoConvenio As String * 20
    NomeEmpresa As String * 20
    CodigoBanco As String * 3
    NomeBanco As String * 20
    DataGeracao As String * 8
    NumeroSeq As String * 6
    VersaoLayout As String * 2
    Filler As String * 69
End Type
Private Type FebrabanA2
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    Filler1 As String * 8
    TipoInscricao As String * 1
    NumeroInscricao As String * 15
    Agencia As String * 5
    NumeroConta As String * 10
    Filler2 As String * 5
    CodigoCedente As String * 9
    Filler3 As String * 11
    NomeEmpresa As String * 30
    NomeBanco As String * 30
    Filler4 As String * 10
    CodigoRetorno As String * 1
    DataGeracao As String * 8 'DDMMAAAA
    Filler5 As String * 6
    NumeroSeq As String * 6
    NumeroVersao As String * 3
    Filler6 As String * 74
End Type

'Private Type FebrabanA2
'    CodigoRegistro As String * 2
'    CodigoRetorno As String * 7
'    CodigoCobrança As String * 2
'    NomeCobrança As String * 8
'    CodigoPrefeitura As String * 11
'    NomePrefeitura As String * 26
'    CodigoBanco As String * 3
'    NomeBanco As String * 7
'    DataCredito As String * 6
'    Filler1 As String * 4
'    Filler2 As String * 38
'    NumRegistros As String * 6
'End Type

Private Type FebrabanB 'CADASTRAMENTO DE DEBITO AUTOMATICO
   CodigoRegistro As String * 1
   CODREDUZIDO As String * 15
'   Distrito As String * 2
'   Setor As String * 2
'   Quadra As String * 4
 '  Lote As String * 5
  ' Seq As String * 2
   FillerID As String * 10
   CodAgencia As String * 4
   ContaCliente As String * 14
   DataOpcao As String * 8
   Filler As String * 97
   CodMovimento As String * 1
End Type

Private Type FebrabanF
    CodigoRegistro As String * 1
    Distrito As String * 2
    Setor As String * 2
    Quadra As String * 4
    Lote As String * 5
    Seq As String * 2
    FillerID As String * 10
    CodAgencia As String * 4
    ContaCliente As String * 14
    DataVencto As String * 8
    ValorDebito As String * 15
    CodRetorno As String * 2
    NumDoc As String * 9
    Filler1 As String * 51
    Filler2 As String * 20
    CodMovimento As String * 1
End Type
Private Type FebrabanG
   CodigoRegistro As String * 1
   ContaPrefeitura As String * 20
   DataPagamento As String * 8
   DataCredito As String * 8
   PreCodBarra As String * 4
   ValorRecebido As String * 11
   CodigoMunic As String * 4
   DataVencto As String * 8
   NumDocumento As String * 9
   NumParcela As String * 2
   SituacaoRetorno As String * 2
   FillerSmar As String * 4
   ValorRetornado As String * 12
   ValorTarifa As String * 7
   NumSeq As String * 8
   CodAgencia As String * 8
   FormaPagamento As String * 1
   NumAutentica As String * 23
   Filler As String * 10
End Type

Private Type FebrabanG2U
    CodigoBanco As String * 3
    NumeroLote As String * 4
    TipoRegistro As String * 1
    NumeroSeq As String * 5
    CodSegmento As String * 1
    Filler1 As String * 1
    CodigoMov As String * 2
    JurosMulta As String * 15
    ValorDesconto As String * 15
    ValorAbatimento As String * 15
    ValorIOF As String * 15
    ValorPago As String * 15
    ValorCreditado As String * 15
    ValorOutrasDespesas As String * 15
    ValorOutrosCreditos As String * 15
    DataOcorrencia As String * 8 'DDMMAAAA
    DataCredito As String * 8 'DDMMAAAA
    Outros As String * 87
End Type

Private Type SimplesNacionalHeader
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    CodigoConvenio As String * 20
    DataGeracao As String * 8
    NumeroRemessa As String * 6
    NumeroVersao As String * 2
    Filler1 As String * 22
    Filler2 As String * 8
    CodigoBanco As String * 3
    Filler3 As String * 422
End Type
Private Type SimplesNacionalDetalhe
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    DataArrecada As String * 8
    DataVencimento As String * 8
    Filler1 As String * 12
    Filler2 As String * 37
    Cnpj As String * 14
    Filler3 As String * 11
    Esfera As String * 1
    Competencia As String * 6
    ValorPrincipal As String * 17
    ValorMulta As String * 17
    ValorJuros As String * 17
    Filler4 As String * 87
    CodigoBanco As String * 3
    CodigoAgencia As String * 4
    Filler5 As String * 249
End Type
Private Type SimplesNacionalTrailer
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    TotalRegistro As String * 6
    ValorRegistro As String * 17
    Filler As String * 468
End Type

Dim aFebrabanA() As FebrabanA
Dim aFebrabanB() As FebrabanB
Dim aFebrabang2U() As FebrabanG2U
Dim aFebrabanA2() As FebrabanA2, nCodBanco As Integer
Dim aSimplesH() As SimplesNacionalHeader
Dim aSimplesD() As SimplesNacionalDetalhe
Dim aSimplesT() As SimplesNacionalTrailer

Dim RegistroG2U As FebrabanG2U
Dim Registro As FebrabanG, Registro2 As FebrabanF, RegistroB As FebrabanB, sDataCredito As String, sBancoTexto As String, RegistroSN As SimplesNacionalDetalhe
Dim IsValidFile As Boolean, sFullPath As String, DA As Boolean

Private Sub cmdOK_Click()
Dim x As Integer, Achou As Boolean, Sql As String, RdoAux As rdoResultset
Dim nSim As Integer, nNao As Integer, nSeq As Integer
lstLog.Clear
Achou = False: nSim = 0: nNao = 0
'If frmMdi.frTeste.Visible = True Then
 '   sPathArqBanco = "C:\Trabalho\GTI\Bancos"
'End If

'MsgBox "spatharqbanco=" & sPathArqBanco

For x = 0 To File1.ListCount - 1
'    MsgBox File1.List(x)
    If File1.Selected(x) Then
       Achou = True
       IsValidFile = False
       sFullPath = File1.Path & "\" & File1.List(x)
       
 '      MsgBox "sfullpath=" & sFullPath
       
       LeArquivo
       If IsValidFile Then
          nSim = nSim + 1
          'cria os diretorios
          If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4), vbDirectory) = "" Then
             'cria o ano
             MkDir sPathArqBanco & "\" & Right$(sDataCredito, 4)
          End If
          If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2), vbDirectory) = "" Then
            'cria o mes
             MkDir sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2)
          End If
          If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2), vbDirectory) = "" Then
            'cria o dia
             MkDir sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2)
          End If
          If Dir$(sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2) & "\" & File1.List(x)) = "" Then
             FileCopy sFullPath, sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2) & "\" & File1.List(x)
             Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'"
             Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             If RdoAux.RowCount = 0 Then
                nSeq = 1
                RdoAux.Close
             Else
                Sql = "SELECT MAX(SEQ) AS MAXIMO FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                nSeq = RdoAux!maximo + 1
                RdoAux.Close
             End If
             Sql = "INSERT ARQUIVOBANCO(DATACREDITO,SEQ,CODBANCO,CODAGENCIA,DATAINCLUSAO,NOMEARQ,DA) VALUES('"
             Sql = Sql & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'," & nSeq & "," & nCodBanco & "," & "888" & ",'"
             Sql = Sql & Format(Now, "mm/dd/yyyy") & "','" & File1.List(x) & "'," & IIf(DA, 1, 0) & ")"
             cn.Execute Sql, rdExecDirect
             lstLog.AddItem File1.List(x) & " (" & sBancoTexto & ")" & " foi incluido no diretório."
          Else
             If MsgBox("O arquivo " & UCase(File1.List(x)) & " ja existe deseja substitui-lo ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
                'terminar a baixa no sqlL
                Sql = "DELETE FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "' AND "
                Sql = Sql & "NOMEARQ='" & File1.List(x) & "'"
                cn.Execute Sql, rdExecDirect
                Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount = 0 Then
                   nSeq = 1
                   RdoAux.Close
                Else
                   Sql = "SELECT MAX(SEQ) AS MAXIMO FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'"
                   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                   nSeq = RdoAux!maximo + 1
                   RdoAux.Close
                End If
                Sql = "INSERT ARQUIVOBANCO(DATACREDITO,SEQ,CODBANCO,CODAGENCIA,DATAINCLUSAO,NOMEARQ,DA) VALUES('"
                Sql = Sql & Format(CDate(sDataCredito), "mm/dd/yyyy") & "'," & nSeq & "," & nCodBanco & "," & "888" & ",'"
                Sql = Sql & Format(Now, "mm/dd/yyyy") & "','" & File1.List(x) & "'," & IIf(DA, 1, 0) & ")"
                cn.Execute Sql, rdExecDirect
                FileCopy sFullPath, sPathArqBanco & "\" & Right$(sDataCredito, 4) & "\" & Mid$(sDataCredito, 4, 2) & "\" & Left$(sDataCredito, 2) & "\" & File1.List(x)
                lstLog.AddItem File1.List(x) & " subtituiu o arquivo existente."
             Else
                nSim = nSim - 1
                lstLog.AddItem File1.List(x) & " não subtituiu o arquivo existente."
             End If
          End If
       Else
          nNao = nNao + 1
          lstLog.AddItem File1.List(x) & " não é um arquivo de banco válido."
       End If
    End If
Next

If Not Achou Then
   MsgBox "Nenhum arquivo foi selecionado.", vbCritical, "Atenção"
   Exit Sub
End If

If nNao > 0 Then
   MsgBox CStr(nSim) & " arquivo(s) foram incluido(s) no diretório e " & CStr(nNao) & " arquivo(s) não são arquivo(s) bancário(s)."
Else
   If nSim > 0 Then
      MsgBox CStr(nSim) & " arquivo(s) foram incluido(s) no diretório."
   End If
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
lstLog.Clear
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo Erro
Dir1.Path = Drive1.Drive
Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub LeArquivo()
On Error GoTo Erro

Dim Header As FebrabanA, HeaderSN As SimplesNacionalHeader
Dim Header2 As FebrabanA2, nContP As Integer, nContN As Integer, nContI As Integer
Dim Posicao  As Long, RdoAux As rdoResultset, Sql As String
Dim nCountB As Integer, nCodReduz As Long, sDataOpcao As String, nCodMov As Integer
Dim dDataOpcao As Date

ReDim aFebrabanA(0)
ReDim aFebrabanA2(0)
ReDim aSimplesH(0)
nContP = 0: nContN = 0: nContI = 0

If InStr(1, sFullPath, "DAF607", vbBinaryCompare) = 0 Then 'verifica se é arquivo do simples nacional
    Open sFullPath For Binary Access Read Write As #1
        Get #1, 1, Header
        aFebrabanA(0).CodigoRegistro = Trim$(Header.CodigoRegistro)
        aFebrabanA(0).CodigoRemessa = Trim$(Header.CodigoRemessa)
        aFebrabanA(0).CodigoConvenio = Trim$(Header.CodigoConvenio)
        aFebrabanA(0).NomeEmpresa = Trim$(Header.NomeEmpresa)
        aFebrabanA(0).CodigoBanco = Trim$(Header.CodigoBanco)
        aFebrabanA(0).NomeBanco = Trim$(Header.NomeBanco)
        aFebrabanA(0).DataGeracao = Trim$(Header.DataGeracao)
        aFebrabanA(0).NumeroSeq = Trim$(Header.NumeroSeq)
        aFebrabanA(0).VersaoLayout = Trim$(Header.VersaoLayout)
        aFebrabanA(0).Filler = Trim$(Header.Filler)
        nCodBanco = Val(aFebrabanA(0).CodigoBanco)
        Posicao = Len(Header) + 3
        Get #1, Posicao, Registro
        If Left$(aFebrabanA(0).Filler, 10) <> "DEBITO AUT" And Left$(aFebrabanA(0).Filler, 10) <> "DÉBITO AUT" Then
           sDataCredito = Right$(Registro.DataCredito, 2) & "/" & Mid$(Registro.DataCredito, 5, 2) & "/" & Left(Registro.DataCredito, 4)
           sBancoTexto = Trim$(aFebrabanA(0).NomeBanco)
           DA = False
        Else
           sBancoTexto = Trim$(aFebrabanA(0).NomeBanco) & " (DÉBITO AUT.)"
           Do While Not EOF(1)
              Get #1, Posicao, Registro2
              If Registro2.CodigoRegistro = "F" Or Registro2.CodigoRegistro = "B" Then
                 sDataCredito = Right$(Registro2.DataVencto, 2) & "/" & Mid$(Registro2.DataVencto, 5, 2) & "/" & Left(Registro2.DataVencto, 4)
                 Exit Do
              End If
              Posicao = Posicao + Len(Registro2) + 2
           Loop
           DA = True
        End If
    Close #1
Else
    Open sFullPath For Binary Access Read Write As #1
        Get #1, 1, HeaderSN
        aSimplesH(0).CodigoRegistro = Trim$(HeaderSN.CodigoRegistro)
        aSimplesH(0).SeqRegistro = Trim$(HeaderSN.SeqRegistro)
        aSimplesH(0).CodigoConvenio = Trim$(HeaderSN.CodigoConvenio)
        aSimplesH(0).DataGeracao = Trim$(HeaderSN.DataGeracao)
        aSimplesH(0).NumeroRemessa = Trim$(HeaderSN.NumeroRemessa)
        aSimplesH(0).NumeroVersao = Trim$(HeaderSN.NumeroVersao)
        aSimplesH(0).Filler1 = Trim$(HeaderSN.Filler1)
        aSimplesH(0).Filler2 = Trim$(HeaderSN.Filler2)
        aSimplesH(0).CodigoBanco = Trim$(HeaderSN.CodigoBanco)
        aSimplesH(0).Filler3 = Trim$(HeaderSN.Filler3)
        
        nCodBanco = Val(aSimplesH(0).CodigoBanco)
        Posicao = Len(HeaderSN) + 3
        Get #1, Posicao, RegistroSN
        sDataCredito = Right$(RegistroSN.DataArrecada, 2) & "/" & Mid$(RegistroSN.DataArrecada, 5, 2) & "/" & Left(RegistroSN.DataArrecada, 4)
        sBancoTexto = "Arquivo do Simples Nacional banco (" & nCodBanco & ")"
        DA = False
        IsValidFile = True
    Close #1
    Exit Sub
End If
 
 
If aFebrabanA(0).CodigoRegistro = "A" And IsNumeric(aFebrabanA(0).CodigoRemessa) Then
    IsValidFile = True
Else
    Open sFullPath For Binary Access Read Write As #1
        Get #1, 1, Header2
        aFebrabanA2(0).DataGeracao = Trim$(Header2.DataGeracao)
        aFebrabanA2(0).CodigoBanco = Trim$(Header2.CodigoBanco)
        aFebrabanA2(0).NomeBanco = Trim$(Header2.NomeBanco)
        nCodBanco = Val(aFebrabanA2(0).CodigoBanco)
        
        Posicao = Len(Header2) + 3
        Do While Not EOF(1)
             Get #1, Posicao, RegistroG2U
             If RegistroG2U.CodSegmento = "U" Then
                If RegistroG2U.CodigoMov = "06" Or RegistroG2U.CodigoMov = "17" Then
                  sDataCredito = Left(RegistroG2U.DataCredito, 2) & "/" & Mid(RegistroG2U.DataCredito, 3, 2) & "/" & Right(RegistroG2U.DataCredito, 4)
                  Exit Do
                End If
             End If
             Posicao = Posicao + Len(RegistroG2U) + 2
        Loop
        
    Close #1
    If (Trim$(aFebrabanA2(0).NomeBanco) = "BANCO SANTANDER BANESPA" Or Trim$(aFebrabanA2(0).NomeBanco) = "BANCO SANTANDER (BRASIL) S/A") And IsDate(sDataCredito) Then
        sBancoTexto = "BANESPA 2"
        IsValidFile = True
    End If
    If (Trim$(aFebrabanA2(0).NomeBanco)) = "BANCO DO BRASIL" And IsDate(sDataCredito) Then
        sBancoTexto = "RETORNO BB"
        IsValidFile = True
    End If
    
End If

If DA Then
    nCountB = 0: ReDim aFebrabanB(0)
    Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header
    Posicao = Len(Header) + 3
    Do While Not EOF(1)
         Get #1, Posicao, RegistroB
         If RegistroB.CodigoRegistro = "B" Then
              Get #1, Posicao, RegistroB
              aFebrabanB(nCountB).CodigoRegistro = RegistroB.CodigoRegistro
              aFebrabanB(nCountB).CODREDUZIDO = RegistroB.CODREDUZIDO
              If Left(RegistroB.CODREDUZIDO, 2) <> "00" Then
                GoTo Proximo
              End If
              
              'aFebrabanB(nCountB).Distrito = RegistroB.Distrito
'              aFebrabanB(nCountB).Setor = RegistroB.Setor
 '             aFebrabanB(nCountB).Quadra = RegistroB.Quadra
  '            aFebrabanB(nCountB).Lote = RegistroB.Lote
   '           aFebrabanB(nCountB).Seq = RegistroB.Seq
              aFebrabanB(nCountB).FillerID = RegistroB.FillerID
              aFebrabanB(nCountB).CodAgencia = RegistroB.CodAgencia
              aFebrabanB(nCountB).ContaCliente = RegistroB.ContaCliente
              aFebrabanB(nCountB).DataOpcao = RegistroB.DataOpcao
              aFebrabanB(nCountB).Filler = RegistroB.Filler
              aFebrabanB(nCountB).CodMovimento = RegistroB.CodMovimento
              With aFebrabanB(nCountB)
                  sDataOpcao = ConvDataSerial(.DataOpcao)
                  nCodMov = .CodMovimento
'                  Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE DISTRITO=" & .Distrito & " AND SETOR=" & .Setor
 '                 Sql = Sql & " AND QUADRA=" & .Quadra & " AND LOTE=" & .Lote & " AND SEQ=" & .Seq
 '                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
 '                 If RdoAux.RowCount > 0 Then
 '                    nCodReduz = RdoAux!CodReduzido
 '                 Else
 '                   GoTo PROXIMO
 ''                 End If
  '                RdoAux.Close
                  nCodReduz = .CODREDUZIDO
                                                      
                  Sql = "SELECT * FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz & " AND CODBANCO=" & nCodBanco
                  Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                  If RdoAux.RowCount > 0 Then
                     If nCodMov = 1 Then 'EXCLUSAO
                        If CDate(sDataOpcao) > Format(RdoAux!DataOpcao, "dd/mm/yyyy") Then
                            Sql = "DELETE FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz & " AND CODBANCO=" & nCodBanco
                            cn.Execute Sql, rdExecDirect
                            nContN = nContN + 1
                        Else
                            nContI = nContI + 1
                        End If
                     Else
                        nContI = nContI + 1
                     End If
                     dDataOpcao = Format(RdoAux!DataOpcao, "dd/mm/yyyy")
                  Else
                     If nCodMov = 2 Then 'INCLUSAO
                        Sql = "DELETE FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz
                        cn.Execute Sql, rdExecDirect
                         
                        Sql = "INSERT DEBITOAUTOMATICO(CODREDUZ,CODBANCO,CODAGENCIA,NUMEROCONTA,DATAOPCAO,CODIGOPREF) VALUES("
                        Sql = Sql & nCodReduz & "," & nCodBanco & "," & .CodAgencia & "," & .ContaCliente & ",'" & Format(sDataOpcao, "mm/dd/yyyy") & "'," & nCodReduz & ")"
                        cn.Execute Sql, rdExecDirect
                        nContP = nContP + 1
                     Else
                        nContI = nContI + 1
                     End If
                  End If
                  RdoAux.Close
              End With
              
              nCountB = nCountB + 1
              ReDim Preserve aFebrabanB(nCountB)
         End If
         
Proximo:
        Posicao = Posicao + Len(RegistroB) + 2
    Loop
    Close #1
End If



lstLog.AddItem "Debito Automático: Arquivo " & sFullPath
lstLog.AddItem "Registros Adicionados: " & CStr(nContP)
lstLog.AddItem "Registros Excluidos: " & CStr(nContN)
lstLog.AddItem "Registros Ignorados: " & CStr(nContI)
Exit Sub
Erro:
'MsgBox Err.Description
Resume Next
End Sub

Private Function ConvDataSerial(sData As String) As String
ConvDataSerial = Right$(sData, 2) & "/" & Mid(sData, 5, 2) & "/" & Left$(sData, 4)
End Function

