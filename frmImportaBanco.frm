VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImportaBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de Arquivos Bancários"
   ClientHeight    =   8535
   ClientLeft      =   11985
   ClientTop       =   5325
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   8325
   Begin VB.CommandButton Command3 
      Caption         =   "Importar"
      Height          =   345
      Left            =   6600
      TabIndex        =   12
      Top             =   7110
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   345
      Left            =   6600
      TabIndex        =   11
      Top             =   6720
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   345
      Left            =   360
      TabIndex        =   9
      Top             =   4710
      Width           =   915
   End
   Begin Tributacao.jcFrames frProgress 
      Height          =   1155
      Left            =   1950
      Top             =   1740
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2037
      FrameColor      =   255
      FillColor       =   4210688
      TextBoxColor    =   8454016
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   3
      ColorFrom       =   192
      ColorTo         =   8438015
      Begin Tributacao.XP_ProgressBar pBar 
         Height          =   165
         Left            =   150
         TabIndex        =   8
         Top             =   780
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   291
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
         Color           =   255
         Scrolling       =   1
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "ARRECADA08.ret"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   4155
      End
      Begin VB.Label lblFileNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Importando Arquivo 0 de 0"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   90
         Width           =   4305
      End
   End
   Begin VB.TextBox txtResult 
      Appearance      =   0  'Flat
      Height          =   2025
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2460
      Width           =   8145
   End
   Begin VB.ListBox lstNome 
      Height          =   645
      Left            =   20490
      TabIndex        =   2
      Top             =   12750
      Width           =   6105
   End
   Begin VB.ListBox lstArq 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   90
      TabIndex        =   1
      Top             =   330
      Width           =   8145
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   360
      Left            =   6210
      TabIndex        =   0
      ToolTipText     =   "Selecione o arquivo a importar"
      Top             =   4650
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Selecionar &Arquivos"
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
      MICON           =   "frmImportaBanco.frx":0000
      PICN            =   "frmImportaBanco.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog cDialog 
      Left            =   17820
      Top             =   12840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Color           =   8388608
      DialogTitle     =   "Selecione o arquivo de GIA"
   End
   Begin VB.ListBox List1 
      Height          =   2010
      ItemData        =   "frmImportaBanco.frx":00BA
      Left            =   570
      List            =   "frmImportaBanco.frx":00BC
      TabIndex        =   10
      Top             =   5910
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado da importação"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   2250
      Width           =   3345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos Selecionados"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3345
   End
End
Attribute VB_Name = "frmImportaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Banco
    Codigo As Integer
    Nome As String
End Type

Dim aBanco() As Banco, nQtdeArquivos As Integer, nFilePos As Integer, nFileTot As Integer, nCodigoSimples As Long

Private Sub cmdArq_Click()
On Error GoTo Erro:

Dim vFiles As Variant
Dim lFile As Long

txtResult.Text = ""
lstArq.Clear

With cDialog
    .FileName = "" 'Clear the filename
    .CancelError = True
    .MaxFileSize = 30000
    .DialogTitle = "Select File(s)..."
    .flags = cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNHideReadOnly 'Flags, allows Multi select, Explorer style and hide the Read only tag
    .Filter = "All files (*.*)|*.*"
    .ShowOpen
    
    vFiles = Split(.FileName, Chr(0)) 'Splits the filename up in segments
    nQtdeArquivos = UBound(vFiles)
    txtResult.Text = IIf(nQtdeArquivos = 0, 1, nQtdeArquivos) & " arquivo(s) selecionado(s)."
    frProgress.Visible = True
    cmdArq.Enabled = False
    If UBound(vFiles) = 0 Then ' If there is only 1 file then do this
        lstArq.AddItem .FileName
        lstNome.AddItem .FileTitle
        nFileTot = 1
        nFilePos = 1
        ImportaArquivo .FileName, .FileTitle
    Else
        For lFile = 1 To UBound(vFiles) ' More than 1 file then do this until there are no more files
            lstArq.AddItem vFiles(0) + "\" & vFiles(lFile)
            lstNome.AddItem vFiles(lFile)
            nFileTot = UBound(vFiles)
            nFilePos = lFile
            ImportaArquivo vFiles(0) + "\" & vFiles(lFile), CStr(vFiles(lFile))
        Next
    End If
    frProgress.Visible = False
    cmdArq.Enabled = True

End With

MsgBox "Importação concluída!", vbInformation, "Atenção"

Exit Sub
Erro:
If Err.Number = 32755 Then
Else
    MsgBox Err.Description
End If

End Sub

Private Sub Command1_Click()
frmBaixaBancariaNovo.show
End Sub

Private Sub Command2_Click()
List1.Clear
Dim strStartPath As String
    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\01"
    ListFolder strStartPath
    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\02"
    ListFolder strStartPath
    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\03"
    ListFolder strStartPath
    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\04"
    ListFolder strStartPath
    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\05"
    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\06"
'    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\07"
'    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\08"
'    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\09"
'    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\10"
'    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\11"
'    ListFolder strStartPath
'    strStartPath = "\\192.168.200.130\AtualizaGTI\2018\12"
'    ListFolder strStartPath

End Sub

Private Sub ListFolder(sFolderPath As String)
    Dim FS As New FileSystemObject
    Dim FSfolder As Folder
    Dim subfolder As Folder
    Dim i As Integer
    
    Set FSfolder = FS.GetFolder(sFolderPath)
 
    For Each subfolder In FSfolder.SubFolders
        DoEvents
        i = i + 1
        
        '***
        Dim fso As New FileSystemObject
Dim fld As Folder
Dim fil As File
Set fld = fso.GetFolder(subfolder)
For Each fil In fld.Files
  'Debug.Print fil.Name
  List1.AddItem subfolder & "\" & fil.Name
Next
Set fil = Nothing
Set fld = Nothing
Set fso = Nothing
        
        '*****
        
'        List1.AddItem subfolder
    Next subfolder
    Set FSfolder = Nothing
'    MsgBox "Total sub folders in " & sFolderPath & " : " & i
End Sub


Private Sub Command3_Click()
On Error GoTo Erro:

Dim vFiles As Variant
Dim lFile As Long

txtResult.Text = ""
lstArq.Clear


    
    nQtdeArquivos = List1.ListCount
    txtResult.Text = IIf(nQtdeArquivos = 0, 1, nQtdeArquivos) & " arquivo(s) selecionado(s)."
    frProgress.Visible = True
    
        For lFile = 0 To List1.ListCount - 1 ' More than 1 file then do this until there are no more files
            lstArq.AddItem List1.Text
            lstNome.AddItem List1.Text
            nFileTot = nQtdeArquivos
            nFilePos = lFile
            ImportaArquivo List1.List(lFile), GetFileNameFromPath(List1.List(lFile))
        Next
   
    frProgress.Visible = False
    


MsgBox "Importação concluída!", vbInformation, "Atenção"

Exit Sub
Erro:
If Err.Number = 32755 Then
Else
    MsgBox Err.Description
End If

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaBanco
End Sub

Private Sub CarregaBanco()
Dim Sql As String, RdoAux As rdoResultset

ReDim aBanco(0)
Sql = "select codbanco,nomebanco from banco order by codbanco"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aBanco(UBound(aBanco) + 1)
        aBanco(UBound(aBanco)).Codigo = !CodBanco
        aBanco(UBound(aBanco)).Nome = !NomeBanco
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Function RetornaNomeBanco(Codigo As Integer) As String
Dim x As Integer, bFind As Boolean

bFind = False
For x = 1 To UBound(aBanco) - 1
    If aBanco(x).Codigo = Codigo Then
        bFind = True
        Exit For
    End If
Next
If bFind Then
    RetornaNomeBanco = aBanco(x).Nome
Else
    RetornaNomeBanco = ""
End If

End Function

Private Sub ImportaArquivo(sFileName As String, sFileTitle As String)
Dim fso As FileSystemObject, TS As TextStream, Row As String, sTipoArq As String, x As Integer, nRegImportado As Integer, z As Long, nRetorno As Integer
Dim nCodBanco As Integer, sNomeBanco As String, sTipoReg As String, sDataPag As String, sDataCredito As String, sAgencia As String, sContaDeposito As String
Dim nNumDoc As Long, nValorPago As Double, sSitRetorno As String, Sql As String, aDataCredito() As String, RdoAux As rdoResultset, nCodReduz As Long, nCodMov As Integer
Dim sDataVencto As String, sCNPJ As String, nAno As Integer, nMes As Integer, nValorPrincipal As Double, nValorMulta As Double, nValorJuros As Double, sDataOpcao As String
Dim nValorGuia As Double, sConvenio As String, sCodMov As String, sConta As String, bFind As Boolean, nCodBancoReceptor As Integer, nPos As Long, nTot As Long
Dim nContN As Integer, nContI As Integer, nContP As Integer, nNumDocSimples As Long

On Error GoTo Erro
nContN = 0: nContI = 0: nContP = 0
lblFileNumber.Caption = "Importando arquivo " & nFilePos & " de " & nFileTot
lblFileName.Caption = sFileTitle
PBar.value = 0
nTot = FileRowCount(sFileName)

Sql = "SELECT MAX(numero_documento) AS MAXIMO FROM importacao_banco where numero_documento<2000000"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)

nNumDocSimples = Val(SubNull(RdoAux!maximo)) + 1
RdoAux.Close
    


ReDim aDataCredito(0)
Set fso = New FileSystemObject


Set TS = fso.OpenTextFile(sFileName, ForReading)
Row = TS.ReadLine
If Left(Row, 1) = "A" And InStr(1, Row, "DEBITO AUT") = 0 Then  'ARQUIVO NORMAL
    sTipoArq = "NORMAL"
ElseIf Left(Row, 15) = "100000001DAF607" Then 'ARQUIVO SIMPLES
    sTipoArq = "SIMPLES"
ElseIf Left(Row, 8) = "00100000" Or Left(Row, 8) = "03300000" Then  'ARQUIVO COBRANÇA BANCO DO BRASIL OU SANTANDER
    sTipoArq = "COBRANCA"
ElseIf InStr(1, Row, "DEBITO AUT") > 0 Then 'ARQUIVO DEBITO AUTOMATICO
    sTipoArq = "DEBAUT"
End If
TS.Close

If sTipoArq = "NORMAL" Then
    GoTo LEARQNORMAL
ElseIf sTipoArq = "SIMPLES" Then
    GoTo LEARQSIMPLES
ElseIf sTipoArq = "COBRANCA" Then
    nTot = nTot / 2
    GoTo LEARQCOBRANCA
ElseIf sTipoArq = "DEBAUT" Then
    GoTo LEDEBAUT
Else
    txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " é inválido!"
End If

Exit Sub

'******** ARRECADAÇÃO ********
LEARQNORMAL:
nPos = 1: nRegImportado = 0
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    Row = TS.ReadLine
    If nPos = 1 Then
        nCodBanco = Mid(Row, 43, 3)
        sNomeBanco = Trim(Mid(Row, 46, 20))
    Else
        sTipoReg = Left(Row, 1)
        If sTipoReg = "G" Then
            sDataCredito = ConvDataSerial(Mid(Row, 30, 8))
            z = BinarySearchString(aDataCredito, sDataCredito)
            If z = -1 Then
                ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                aDataCredito(UBound(aDataCredito)) = sDataCredito
                Sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then GoTo ARQUIVO_EXISTENTE
                RdoAux.Close
            End If
            sDataPag = ConvDataSerial(Mid(Row, 22, 8))
            sDataVencto = ConvDataSerial(Mid(Row, 57, 8))
            nNumDoc = CLng(Mid(Row, 65, 9))
            nValorPago = CDbl(Mid(Row, 82, 12)) / 100
            sAgencia = Mid(Row, 2, 4)
            sConta = Val(Trim(RetornaNumero(Mid(Row, 6, 16))))
            sConvenio = Trim(RetornaNumero(Mid(Row, 3, 20)))
            Sql = "select numdocumento from numdocumento where numdocumento=" & nNumDoc
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                sSitRetorno = "00-BAIXA NORMAL"
            Else
                sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
            End If
            RdoAux.Close
            If Not IsDate(sDataVencto) Then sDataVencto = sDataPag
SqlNormal:
            Sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,agencia,situacao_retorno,data_vencimento,convenio,conta_corrente) values("
            Sql = Sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ","
            Sql = Sql & nNumDoc & ",'" & sAgencia & "','" & sSitRetorno & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sConvenio & "','" & sConta & "')"
            cn.Execute Sql, rdExecDirect
            nRegImportado = nRegImportado + 1
        End If
    End If
    nPos = nPos + 1
Loop
TS.Close
txtResult.Text = txtResult.Text & vbCrLf & "Importado " & nRegImportado & " registros do arquivo " & sFileName & " - Arrecadação (" & sNomeBanco & ")"
Exit Sub

LEARQSIMPLES:
nPos = 1
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    Row = TS.ReadLine
    If nPos = 1 Then
        nCodBanco = Mid(Row, 76, 3)
        sNomeBanco = RetornaNomeBanco(nCodBanco)
        If sNomeBanco = "" Then
            sNomeBanco = "Sem convênio"
        End If
    Else
        sTipoReg = Left(Row, 1)
        If sTipoReg = "2" Then
            sDataCredito = ConvDataSerial(Mid(Row, 10, 8))
            z = BinarySearchString(aDataCredito, sDataCredito)
            If z = -1 Then
                ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                aDataCredito(UBound(aDataCredito)) = sDataCredito
                Sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then GoTo ARQUIVO_EXISTENTE
                RdoAux.Close
            End If
            sDataPag = ConvDataSerial(Mid(Row, 10, 8))
            sDataVencto = ConvDataSerial(Mid(Row, 18, 8))
            nValorPrincipal = CDbl(Mid(Row, 107, 17)) / 100
            nValorMulta = CDbl(Mid(Row, 141, 17)) / 100
            nValorJuros = CDbl(Mid(Row, 124, 17)) / 100
            nValorGuia = nValorPrincipal + nValorJuros + nValorMulta
            nAno = Val(Mid(Row, 101, 4))
            nMes = Val(Mid(Row, 105, 2))
            sCNPJ = Mid(Row, 75, 14)
            sAgencia = Mid(Row, 223, 4)
            'sSitRetorno = "CNPJ: " & Format(sCNPJ, "0#\.###\.###/####-##")
            
            nCodigoSimples = 0
            
'            nNumDoc = GravaSimples(sCNPJ, sDataVencto, nAno, nMes, nValorGuia, sFileTitle, sDataCredito, nCodBanco, sAgencia)
'            If nNumDoc > 0 Then
                Sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,agencia,data_vencimento,simples_nacional,cnpj,ano,mes,codigo_reduzido) values("
                Sql = Sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorGuia)) & ","
                Sql = Sql & nNumDocSimples & ",'" & sAgencia & "','" & Format(sDataVencto, "mm/dd/yyyy") & "'," & 1 & ",'" & sCNPJ & "'," & nAno & "," & nMes & "," & nCodigoSimples & ")"
                cn.Execute Sql, rdExecDirect
'            End If
            nNumDocSimples = nNumDocSimples + 1
            nRegImportado = nRegImportado + 1
        End If
    End If
    nPos = nPos + 1
Loop
TS.Close
txtResult.Text = txtResult.Text & vbCrLf & "Incluído o arquivo " & sFileName & " - S.Nacional (" & sNomeBanco & ")"
Exit Sub


LEARQCOBRANCA:
nPos = 1
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    Row = TS.ReadLine
    If nPos = 1 Then
       'Header de arquivo
        nCodBanco = Left(Row, 3)
        sNomeBanco = Trim(Mid(Row, 103, 30))
        sConvenio = Mid(Row, 35, 7)
    ElseIf nPos = 2 Then
       'Header de lote, nada a importar
    Else
        If Left(Row, 8) = "00199999" Or (Left(Row, 3) = "033" And Mid(Row, 14, 1) = " ") Then
           'Trailer de Lote
            Exit Do
        Else
            sTipoReg = Mid(Row, 14, 1)
            If sTipoReg = "T" Then
                sCodMov = Mid(Row, 16, 2)
                If sCodMov = "06" Or sCodMov = "17" Then
                    If Mid(Row, 45, 2) = "00" Then
                        nNumDoc = Val(Mid(Row, 47, 8))
                    ElseIf Mid(Row, 45, 1) <> "0" Then
                        nNumDoc = Val(Mid(Row, 45, 8))
                    Else
                        nNumDoc = Val(Mid(Row, 46, 8))
                    End If
                    If nNumDoc > 200000 And nNumDoc < 300000 Then
                        nNumDoc = Val(Mid(sReg, 45, 10))
                    End If
                    
                    Sql = "select numdocumento from numdocumento where numdocumento=" & nNumDoc
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux.RowCount > 0 Then
                        sSitRetorno = "00-BAIXA NORMAL"
                    Else
                        sSitRetorno = "01-DOCUMENTO NÃO ENCONTRADO"
                    End If
                    RdoAux.Close
                    nCodBancoReceptor = Val(Mid(Row, 97, 3))
                    sAgencia = Mid(Row, 100, 5)
                    sConta = Val(Mid(Row, 24, 13))
                    sDataVencto = ConvDataSerialBB(Mid(Row, 74, 8))
                    nValorPago = CDbl(Mid(Row, 84, 15) / 100)
                    'next line read U
                    Row = TS.ReadLine
                    sDataCredito = ConvDataSerialBB(Mid(Row, 146, 8))
                    If Not IsDate(sDataCredito) Then
                        sDataCredito = ConvDataSerialBB(Mid(Row, 138, 8))
                    End If
                    z = BinarySearchString(aDataCredito, sDataCredito)
                    If z = -1 Then
                        ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                        aDataCredito(UBound(aDataCredito)) = sDataCredito
                        Sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux.RowCount > 0 Then GoTo ARQUIVO_EXISTENTE
                        RdoAux.Close
                    End If
                    sDataPag = ConvDataSerialBB(Mid(Row, 138, 8))
                    If Val(Mid(Row, 78, 15)) > 0 Then
                        nValorPago = CDbl(Mid(Row, 78, 15) / 100)
                    End If
                    If sDataVencto = "00/00/0000" Or Not IsDate(sDataVencto) Then
                        sDataVencto = sDataPag
                    End If
                                        
                    bFind = False
                    If sCodMov = "17" Then
                        Sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and "
                        Sql = Sql & "nome_arquivo='" & sFileName & "' and numero_documento=" & nNumDoc
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If RdoAux.RowCount > 0 Then
                            bFind = True
                        End If
                        RdoAux.Close
                    End If
                    
                    If nNumDoc > 0 Then
                        If sCodMov = "06" Or (sCodMov = "17" And bFind = False) Then
                            Sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,agencia,situacao_retorno,data_vencimento,convenio,conta_corrente,codigo_banco_receptor) values("
                            Sql = Sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ","
                            Sql = Sql & nNumDoc & ",'" & sAgencia & "','" & sSitRetorno & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sConvenio & "','" & sConta & "'," & nCodBancoReceptor & ")"
                            cn.Execute Sql, rdExecDirect
                            nRegImportado = nRegImportado + 1
                        End If
                    Else
                        DoEvents
                    End If
                End If
            End If
        End If
    End If
    nPos = nPos + 1
Loop
TS.Close
txtResult.Text = txtResult.Text & vbCrLf & "Incluído o arquivo " & sFileName & " - Cobrança (" & sNomeBanco & ")"
Exit Sub

LEDEBAUT:
nPos = 1
Set TS = fso.OpenTextFile(sFileName, ForReading)
Do Until TS.AtEndOfStream
    If nPos Mod 10 = 0 Then CallPb nPos, nTot
    Row = TS.ReadLine
    If nPos = 1 Then
        nCodBanco = Mid(Row, 43, 3)
        sNomeBanco = Trim(Mid(Row, 46, 20))
    Else
        sTipoReg = Left(Row, 1)
        If sTipoReg = "F" Then
            sDataCredito = ConvDataSerial(Mid(Row, 45, 8))
            z = BinarySearchString(aDataCredito, sDataCredito)
            If z = -1 Then
                ReDim Preserve aDataCredito(UBound(aDataCredito) + 1)
                aDataCredito(UBound(aDataCredito)) = sDataCredito
                Sql = "select * from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & "' and nome_arquivo='" & sFileTitle & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then GoTo ARQUIVO_EXISTENTE
                RdoAux.Close
            End If
            sDataPag = sDataCredito
            sDataVencto = sDataCredito
            nCodReduz = Val(Mid(Row, 11, 6))
            nValorPago = CDbl(Mid(Row, 53, 15)) / 100
            sAgencia = Mid(Row, 27, 4)
            sContaDeposito = Trim(Mid(Row, 31, 9))
            nCodBancoReceptor = nCodBanco
            nRetorno = Val(Mid(Row, 68, 2))
            
            Select Case nRetorno
                Case "00"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Débito Efetuado"
                Case "01"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Insuficiência de Fundos"
                Case "02"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Conta Corrente não Cadastrada"
                Case "04"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Outras Restrições"
                Case "10"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Agência em Regime de Encerramento"
                Case "12"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Valor Inválido"
                Case "13"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Data de Lançamento inválida"
                Case "14"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Agência Inválida"
                Case "15"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "DAC da conta corrente inválido"
                Case "18"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Data do Débito anterior ao do processamento"
                Case "30"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Sem contrato de débito automático"
                Case "96"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Manutenção do Cadastro"
                Case "97"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - Não Encontrado"
                Case "98"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - não efetuado, fora de tempo habil"
                Case "99"
                        sSitRetorno = Format(nRetorno, "00") & " - " & "Cancelamento - cancelado conforme solicitado"
                Case Else
                       sSitRetorno = Format(nRetorno, "00") & " - " & "Erro Indefinido"
            End Select
            
            Sql = "SELECT lancamento.descreduz, debitoparcela.statuslanc, situacaolancamento.descsituacao, debitoparcela.datavencimento, debitoparcela.datadebase,"
            Sql = Sql & "debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
            Sql = Sql & "debitoparcela.CODCOMPLEMENTO , parceladocumento.NumDocumento FROM lancamento INNER JOIN debitoparcela ON lancamento.codlancamento = debitoparcela.codlancamento INNER JOIN "
            Sql = Sql & "situacaolancamento ON debitoparcela.statuslanc = situacaolancamento.codsituacao INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND "
            Sql = Sql & "debitoparcela.anoexercicio = parceladocumento.anoexercicio AND debitoparcela.codlancamento = parceladocumento.codlancamento AND "
            Sql = Sql & "debitoparcela.seqlancamento = parceladocumento.seqlancamento AND debitoparcela.numparcela = parceladocumento.numparcela AND debitoparcela.CODCOMPLEMENTO = parceladocumento.CODCOMPLEMENTO "
            Sql = Sql & "WHERE (DEBITOPARCELA.SEQLANCAMENTO=0) AND (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.CODLANCAMENTO = 1) AND "
            Sql = Sql & "(DEBITOPARCELA.NUMPARCELA > 0) AND (DEBITOPARCELA.DATAVENCIMENTO = '" & Format(sDataVencto, "mm/dd/yyyy") & "')"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount = 0 Then
                   ' MsgBox "Código não encontrado nos optantes de débito automático (" & nCodReduz & ")"
                    Sql = "delete from importacao_banco where codigo_banco=" & nCodBanco & " and data_credito='" & Format(sDataCredito, "mm/dd/yyyy") & " ' and nome_arquivo='" & sFileName & "' and data_controle is null"
                   ' cn.Execute Sql, rdExecDirect
                    txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " - Déb.Automático (" & sNomeBanco & ") não foi incluido pois contêm erros."
                    GoTo Close_DA
                Else
                    nNumDoc = !NumDocumento
                End If
               .Close
            End With
            If nNumDoc > 0 Then
                Sql = "insert importacao_banco(codigo_banco,data_credito,nome_arquivo,data_pagamento,valor_pago,numero_documento,debito_automatico,agencia,situacao_retorno,data_vencimento,conta_deposito,codigo_reduzido,codigo_banco_receptor) values("
                Sql = Sql & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','" & sFileTitle & "','" & Format(sDataPag, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValorPago)) & ","
                Sql = Sql & nNumDoc & ",1" & ",'" & sAgencia & "','" & sSitRetorno & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sContaDeposito & "'," & nCodReduz & "," & Val(nCodBancoReceptor) & ")"
                cn.Execute Sql, rdExecDirect
                nRegImportado = nRegImportado + 1
            End If
        ElseIf sTipoReg = "B" Then
            nCodReduz = Val(Mid(Row, 11, 6))
            sAgencia = Mid(Row, 27, 4)
            sConta = Mid(Row, 31, 14)
            sDataOpcao = ConvDataSerial(Mid(Row, 45, 8))
            nCodMov = Mid(Row, 150, 1)
           
            Sql = "SELECT * FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz & " AND CODBANCO=" & nCodBanco
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
               If nCodMov = 1 Then 'EXCLUSAO
                  If CDate(sDataOpcao) > Format(RdoAux!DataOpcao, "dd/mm/yyyy") Then
                      Sql = "DELETE FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodReduz & " AND CODBANCO=" & nCodBanco
                      'cn.Execute Sql, rdExecDirect
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
                  'cn.Execute Sql, rdExecDirect
                   
                  Sql = "INSERT DEBITOAUTOMATICO(CODREDUZ,CODBANCO,CODAGENCIA,NUMEROCONTA,DATAOPCAO,CODIGOPREF) VALUES("
                  Sql = Sql & nCodReduz & "," & nCodBanco & "," & sAgencia & "," & sConta & ",'" & Format(sDataOpcao, "mm/dd/yyyy") & "'," & nCodReduz & ")"
                  'cn.Execute Sql, rdExecDirect
                  nContP = nContP + 1
               Else
                  nContI = nContI + 1
               End If
            End If
            RdoAux.Close
        
            
        End If
    
    End If
    nPos = nPos + 1
Loop
txtResult.Text = txtResult.Text & vbCrLf & "Incluído o arquivo " & sFileName & " - Déb.Automático (" & sNomeBanco & ")"
If nContP > 0 Or nContI > 0 Or nContN > 0 Then
    txtResult.Text = txtResult.Text & vbCrLf & nContP & " Optantes incluidos, " & nContN & " Optantes removidos e " & nContI & " Optantes ignorados."
End If
Close_DA:
TS.Close
Exit Sub

ARQUIVO_EXISTENTE:
txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " " & sNomeBanco & " já está baixado, portanto foi ignorado."
Exit Sub

Erro:
If Err.Number = 62 Then
    txtResult.Text = txtResult.Text & vbCrLf & "O arquivo " & sFileName & " " & sNomeBanco & " é inválido, portanto foi ignorado."
    Resume Next
Else
    If rdoErrors.Count > 1 Then
        If rdoErrors(1).Number = 2627 Then 'duplicidade
            Resume Next
        End If
    Else
        MsgBox Err.Description
        Resume Next
    End If
End If
End Sub

Private Function GravaSimples(sCNPJ As String, sDataVencto As String, nAno As Integer, nMes As Integer, nValor As Double, sArquivo As String, sDataCredito As String, nCodBanco As Integer, sAgencia As String) As Long
Dim nNumDoc As Long, Sql As String, RdoAux As rdoResultset, nCodReduz As Long, nNumParc As Integer, bAchou As Boolean, nSeq As Integer, nCompl As Integer, dDataVencto As Date, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset
On Error GoTo Erro
DoEvents

'BUSCA CÓDIGO
Sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE DATAENCERRAMENTO IS NULL and CONVERT(BIGINT, cnpj) = " & Val(sCNPJ)
Sql = Sql & " OR CNPJ='" & Format(sCNPJ, "00\.000\.000/0000-00") & "' AND DATAENCERRAMENTO IS NULL ORDER BY CODIGOMOB DESC"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        nCodReduz = !codigomob
        .Close
    Else
        .Close
        Sql = "SELECT CODCIDADAO,CNPJ FROM CIDADAO WHERE CNPJ = '" & RetornaNumero(sCNPJ) & "' OR "
        Sql = Sql & "CNPJ='" & Format(sCNPJ, "00\.000\.000/0000-00") & "' ORDER BY CODCIDADAO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                nCodReduz = !CodCidadao
            Else
                'CNPJ NÃO LOCALIZADO
                Sql = "SELECT * FROM SIMPLESCNPJ WHERE CNPJ='" & sCNPJ & "' AND ANOCOMP=" & nAno & " AND MESCOMP=" & nMes
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux3.RowCount = 0 Then
                    Sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
                    Sql = Sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(sCNPJ) & "','" & sArquivo & "'," & nCodBanco & ",'" & Format(sDataCredito, "mm/dd/yyyy") & "','"
                    Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & nAno & "," & nMes & "," & Virg2Ponto(CStr(nValor)) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & ",'"
                    Sql = Sql & sAgencia & "'," & 0 & ")"
   '                 cn.Execute Sql, rdExecDirect
                End If
                RdoAux3.Close
                GoTo fim
            End If
        End With
    End If
End With

'BUSCA LANCAMENTO
Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,debitoparcela.CODCOMPLEMENTO, debitoparcela.DataVencimento, debitoparcela.statuslanc, debitotributo.ValorTributo "
Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento Where (debitoparcela.CODREDUZIDO = " & nCodReduz & ") And (debitoparcela.CodLancamento = 5) And (Month(debitoparcela.DataVencimento) = " & Month(CDate(sDataVencto)) & ") And "
Sql = Sql & "(YEAR(debitoparcela.datavencimento) = " & Year(CDate(sDataVencto)) & ") AND (debitotributo.codtributo = 13) and debitotributo.valortributo =" & Virg2Ponto(CStr(nValor)) & " AND statuslanc<>6 "
Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux3
   'EXISTE LANCAMENTO NESTE MÊS/ANO?
    If .RowCount > 0 Then 'SIM
'        nNumDoc = !NumDocumento
       
        nNumParc = !NumParcela 'CAPTURA A PARCELA
        bAchou = False
       'TEM ALGUMA QUE NÃO ESTA PAGA?
        Do Until .EOF
            If !statuslanc = 3 Then
                bAchou = True
                Exit Do
            End If
           .MoveNext
        Loop
        
       'SE ACHOU PEGA A PARCELA
        If bAchou Then
            nSeq = !SeqLancamento
            nCompl = !CODCOMPLEMENTO '---------------> PARCELA PRONTA PARA USO
            nNumDoc = 0
            GoTo fim
        Else
           'SE NÃO ACHAR
           .MoveFirst
            nCompl = 0
           'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
            Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND anoexercicio=" & nAno & " and (codlancamento = 5)"
            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux4
                If IsNull(!maximo) Then
                    nSeq = 0
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
        End If
    
    Else
       'NÃO ACHOU LANCAMENTOS NESTE MÊS/ANO
       'AUMENTA O LANCAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & nAno & ")"
        Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux4
            If IsNull(!maximo) Then
                nSeq = 1
            Else
                nSeq = !maximo + 1
            End If
           .Close
        End With
         
        nCompl = 0
        nNumParc = 1
    End If
 
   .Close
End With


Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & 5 & " AND "
Sql = Sql & "SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount = 0 Then
    'CRIA A PARCELA
    Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO where numdocumento<2000000"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    nNumDoc = RdoAux!maximo + 1
    RdoAux.Close
    
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,emissor) VALUES("
    Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",'" & NomeDeLogin & " (BAIXA BANCÁRIA/SN)" & "')"
 '   cn.Execute Sql, rdExecDirect
    
'    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) "
'    Sql = Sql & "VALUES(" & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & "," & nNumParc & "," & nCompl & ","
'    Sql = Sql & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & 1 & ",'GTI')"
 '   cn.Execute Sql, rdExecDirect
    
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
    Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
    Sql = Sql & nCodReduz & "," & nAno & "," & 5 & "," & nSeq & ","
    Sql = Sql & nNumParc & "," & nCompl & "," & 13 & "," & Virg2Ponto(CStr(nValor)) & ")"
 '   cn.Execute Sql, rdExecDirect
    
    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
    Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
    Sql = Sql & nAno & "," & 5 & "," & nSeq & "," & nNumParc & ","
    Sql = Sql & nCompl & "," & nNumDoc & ")"
 '   cn.Execute Sql, rdExecDirect


End If
'RdoAux.Close

'PROCURA SE O DEBITO JA FOI BAIXADO
Sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE ARQUIVOBANCO='" & sArquivo & "' AND DATACREDITO='" & Format(sDataCredito, "mm/dd/yyyy") & "' AND "
Sql = Sql & "CNPJ='" & sCNPJ & "' AND ANO=" & nAno & " AND MES=" & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    Sql = "INSERT COMPLEMENTOSIMPLES(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,ARQUIVOBANCO,DATACREDITO,VALOR,CNPJ,ANO,MES) VALUES(" & nCodReduz & ","
    Sql = Sql & nAno & "," & 5 & "," & nSeq & "," & nNumParc & "," & nCompl & ",'" & sArquivo & "','" & Format(sDataCredito, "mm/dd/yyyy") & "',"
    Sql = Sql & Virg2Ponto(CStr(nValor)) & ",'" & sCNPJ & "'," & nAno & "," & nMes & ")"
  '  cn.Execute Sql, rdExecDirect
    RdoAux.Close
End If
fim:

nCodigoSimples = nCodReduz
GravaSimples = nNumDoc

Exit Function
Erro:
MsgBox Err.Description

End Function

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Private Function ConvDataSerialBB(sData As String) As String
   ConvDataSerialBB = Left(sData, 2) & "/" & Mid$(sData, 3, 2) & "/" & Right(sData, 4)
End Function

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
DoEvents

End Sub

Private Function FileRowCount(sFile As String)
    Const BufSize As Long = 100000
    Dim T0 As Single
    Dim LfAnsi As String
    Dim f As Integer
    Dim FileBytes As Long
    Dim BytesLeft As Long
    Dim Buffer() As Byte
    Dim strBuffer As String
    Dim BufPos As Long
    Dim LineCount As Long

    T0 = Timer()
    LfAnsi = StrConv(vbLf, vbFromUnicode)
    f = FreeFile(0)
    Open sFile For Binary Access Read As #f
    FileBytes = LOF(f)
    ReDim Buffer(BufSize - 1)
    BytesLeft = FileBytes
    Do Until BytesLeft = 0
        If BufPos = 0 Then
            If BytesLeft < BufSize Then ReDim Buffer(BytesLeft - 1)
            Get #f, , Buffer
            strBuffer = Buffer 'Binary copy of bytes.
            BytesLeft = BytesLeft - LenB(strBuffer)
            BufPos = 1
        End If
        Do Until BufPos = 0
            BufPos = InStrB(BufPos, strBuffer, LfAnsi)
            If BufPos > 0 Then
                LineCount = LineCount + 1
                BufPos = BufPos + 1
            End If
        Loop
    Loop
    Close #f
    FileRowCount = LineCount
    'Add 1 to LineCount if last line of your files do not
    'have a trailing CrLf.
  '  MsgBox "Counted " & Format$(LineCount, "#,##0") & " lines in" & vbNewLine _
         & Format$(FileBytes, "#,##0") & " bytes of text." & vbNewLine _
         & Format$(Timer() - T0, "0.0#") & " seconds."
End Function

Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
