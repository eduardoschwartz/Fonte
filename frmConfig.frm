VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmConfig 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração"
   ClientHeight    =   5205
   ClientLeft      =   9030
   ClientTop       =   4305
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   5280
   Begin VB.CommandButton cmdPagos 
      Caption         =   "Pagos"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2610
      TabIndex        =   11
      Top             =   4770
      Width           =   960
   End
   Begin prjChameleon.chameleonButton btArquivos 
      Height          =   315
      Index           =   1
      Left            =   180
      TabIndex        =   9
      ToolTipText     =   "IP"
      Top             =   4770
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Le Arq"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfig.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btCorrige 
      Height          =   315
      Index           =   0
      Left            =   4140
      TabIndex        =   8
      ToolTipText     =   "IP"
      Top             =   4740
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Corrige"
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConfig.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdClearBD 
      Caption         =   "Limpa BD"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1740
      TabIndex        =   5
      Top             =   7800
      Width           =   1935
   End
   Begin VB.TextBox txtOld 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      TabIndex        =   4
      Top             =   4125
      Width           =   5085
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      TabIndex        =   3
      Top             =   3000
      Width           =   5085
   End
   Begin VB.ListBox lstParam 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "frmConfig.frx":0038
      Left            =   90
      List            =   "frmConfig.frx":003A
      TabIndex        =   2
      Top             =   90
      Width           =   5085
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4140
      TabIndex        =   0
      ToolTipText     =   "Cancelar Edição"
      Top             =   3420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmConfig.frx":003C
      PICN            =   "frmConfig.frx":0058
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   3060
      TabIndex        =   1
      ToolTipText     =   "Gravar os Dados"
      Top             =   3420
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmConfig.frx":01B2
      PICN            =   "frmConfig.frx":01CE
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
      Height          =   165
      Left            =   135
      TabIndex        =   6
      Top             =   3825
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2265
      Left            =   30
      TabIndex        =   12
      Top             =   5220
      Width           =   16815
      _ExtentX        =   29660
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Arquivo"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ShortName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Banco"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Rec."
         Object.Width           =   1835
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CNPJ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Codigo"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Ano"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Mes"
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Dt.Venc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Valor"
         Object.Width           =   1834
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Exer."
         Object.Width           =   953
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Sq"
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Dup"
         Object.Width           =   776
      EndProperty
   End
   Begin VB.CommandButton btCadastro 
      Caption         =   "baixar"
      Height          =   330
      Left            =   1380
      TabIndex        =   10
      Top             =   4770
      Width           =   960
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2295
      TabIndex        =   7
      Top             =   3825
      Width           =   480
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type idUser
    id As Integer
    Nome As String
End Type
Dim aIdUser() As idUser

Private Type Proc
    Numero As Long
    Ano As Integer
    Cancelado As Boolean
End Type


Private Type Registro
    nNumDoc As Long
    nSeq As Integer
    sDataDoc As String
    sDataPag As Date
    sDataCred As Date
    nValorPago As Double
    sAgencia As String
    nValorTarifa As Double
    sSitRetorno As String
    bExiste As Boolean
    bIsentoMJ As Boolean
    sCNPJ As String
    nAno As Integer
    nMes As Integer
    sDataVencto As Date
    nValorTarifaBancaria As Double
    nSomaTributo As Double
    sDataPagCalc As Date
End Type

Private Type Documento
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    sDataVencto As String
    sSit As String
    nNumeroLivro As Integer
    nPaginaLivro As Integer
    bAjuizado As Boolean
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
    nValorTarifa As Double
    nValorDif As Double
    nValorCompensado As Double
    sBx As String
    sDp As String
    nSeqReg As Integer
    bExiste As Boolean
    sCNPJ As String
End Type

Private Type tAREA
    nSeq As Integer
    
End Type


Private Type SIMPLES
    nCodigo As Long
    nAno As Integer
End Type


Private Sub btArquivos_Click(Index As Integer)
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, x As Integer
Dim sNomeArq As String, nCodBanco As Integer, sFullPath As String, sReg As String, sCNPJ As String
Dim sDataIni As String, sDataFim As String, sEncerrada As String, sSuspensa As String
Dim nCodReduzido As Long, sClasse As String, sSimples As String, RdoAux3 As rdoResultset, nNumDoc As Long, FF1 As Integer

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
'GoTo Parte2
Sql = "truncate table mei2"
cn.Execute Sql, rdExecDirect


On Error Resume Next
FF1 = FreeFile()
Open "c:\trabalho\simples.txt" For Binary Access Read Write As FF1

    While Not EOF(FF1)
        Input #FF1, sReg
        sCNPJ = Left(sReg, 8)
        sDataIni = Mid(sReg, 9, 8)
        sDataFim = Mid(sReg, 17, 8)
        sDataIni = ConvDataSerial(sDataIni)
        sDataFim = ConvDataSerial(sDataFim)
        If sDataFim = "00/00/0000" Then sDataFim = "01/01/1900"
        
        Sql = "select cnpj from simples_codigo where cnpj='" & sCNPJ & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            RdoAux.Close
            GoTo proximo
        End If
        RdoAux.Close
        
        
        Sql = "SELECT codigomob,dataencerramento From mobiliario WHERE SUBSTRING(cnpj, 1, 8) = '" & sCNPJ & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            nCodReduzido = RdoAux!codigomob
            If IsNull(RdoAux!dataencerramento) Then
                sEncerrada = "N"
            Else
                sEncerrada = "S"
            End If
        Else
            nCodReduzido = 0
            sEncerrada = "N"
        End If
        RdoAux.Close

        Sql = "insert mei2 (cnpj,codigo,datainicio,datafim,encerrada) values('" & sCNPJ & "'," & nCodReduzido & ",'"
        Sql = Sql & Format(sDataIni, sDataFormat) & "','" & Format(sDataFim, sDataFormat) & "','" & sEncerrada & "')"
 '      Sql = "insert simples_codigo (cnpj) values('" & sCNPJ & "')"
        cn.Execute Sql, rdExecDirect
        DoEvents
        
       'suspenção
        Sql = "SELECT CODTIPOEVENTO,DATAEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduzido
        Sql = Sql & " ORDER BY DATAEVENTO DESC"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                sSuspensa = "N"
            Else
                If !CODTIPOEVENTO = 2 Then
                    sSuspensa = "S"
                Else
                    sSuspensa = "N"
                End If
            End If
           .Close
        End With
        Sql = "update mei2 set suspensa='" & sSuspensa & "' where codigo=" & nCodReduzido
        cn.Execute Sql, rdExecDirect
proximo:
    Wend
CloseFile2:
Close #FF1

'Exit Sub
parte2:
Sql = "select codigo from mei2 where codigo > 0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
Do Until RdoAux.EOF
    nCodReduzido = RdoAux!Codigo
    sSimples = IIf(SNCheck2(nCodReduzido), "S", "N")
    Sql = "update mei2 set esimples='" & sSimples & "' where codigo=" & nCodReduzido
    cn.Execute Sql, rdExecDirect
    RdoAux.MoveNext
Loop
RdoAux.Close


MsgBox "fim"

End Sub

Private Sub btCadastro_Click()
Dim Sql As String, RdoAux As rdoResultset, nAno1 As Integer, nLanc1 As Integer, RdoAux2 As rdoResultset, x As Integer, RdoAux3 As rdoResultset
Dim nCodReduz As Long, aCodigo(17) As Integer, nCodigo1 As Long, nSeq1 As Integer, nParc1 As Integer, nCompl1 As Integer, sDataVencto As String, nValor1 As Double
Dim aOrigem() As Documento, bFind As Boolean, Y As Integer, nCodigo2 As Long, nSeq2 As Integer, nParc2 As Integer, nCompl2 As Integer


If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
ReDim aIdUser(0)

Sql = "select id,nomelogin from usuario order by nomelogin"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aIdUser(UBound(aIdUser) + 1)
        aIdUser(UBound(aIdUser)).id = !id
        aIdUser(UBound(aIdUser)).Nome = !NomeLogin
       .MoveNext
    Loop
   .Close
End With
Mei 'SuspendeEmpresa
'EmiteBoletoCIP
'CorrigeIsencao
'CorrigeHistorico
'CorrigeObsCidadao
'CorrigeDebitoCancel
'CorrigeDebitoObservacao
'CorrigeDebitoParcela
'CorrigeUsuarioCC
'CorrigeTramite
'ProcessoGTI
'Codigo_Usuario
'ContaArea
'Corrige_Obs
'Relatorio_SanMarino
'NaoPagoParaPago
Exit Sub

aCodigo(0) = 1992
aCodigo(1) = 22498
aCodigo(2) = 22499
aCodigo(3) = 22501
aCodigo(4) = 22502
aCodigo(5) = 22503
aCodigo(6) = 22504
aCodigo(7) = 22505
aCodigo(8) = 22506
aCodigo(9) = 22507
aCodigo(10) = 22509
aCodigo(11) = 22510
aCodigo(12) = 22511
aCodigo(13) = 22512
aCodigo(14) = 22513
aCodigo(15) = 22514
aCodigo(16) = 22508


'carrega origem
ReDim aOrigem(0)
Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,debitoparcela.codcomplemento, debitoparcela.datavencimento, SUM(debitotributo.valortributo) AS Soma "
Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitotributo.codtributo <> 3) GROUP BY debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,"
Sql = Sql & "debitoparcela.CODCOMPLEMENTO , debitoparcela.DataVencimento Having (debitoparcela.CODREDUZIDO = 38258) ORDER BY debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, debitoparcela.codcomplemento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aOrigem(UBound(aOrigem) + 1)
        aOrigem(UBound(aOrigem)).nAno = !AnoExercicio
        aOrigem(UBound(aOrigem)).nLanc = !CodLancamento
        aOrigem(UBound(aOrigem)).nSeq = !SeqLancamento
        aOrigem(UBound(aOrigem)).nParc = !NumParcela
        aOrigem(UBound(aOrigem)).nCompl = !CODCOMPLEMENTO
        aOrigem(UBound(aOrigem)).sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
        aOrigem(UBound(aOrigem)).nValorPrincipal = Round(!soma, 2)
       .MoveNext
    Loop
   .Close
End With
'GoTo fim
'fim origem

Sql = "delete from transfere_debito"
cn.Execute Sql, rdExecDirect

For x = 0 To 16
    lblPB.Caption = x
    nCodReduz = aCodigo(x)
    Sql = "select * FROM debitoparcela where codreduzido=" & nCodReduz & " and statuslanc=13"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            nAno1 = !AnoExercicio
            nLanc1 = !CodLancamento
            nSeq1 = !SeqLancamento
            nParc1 = !NumParcela
            nCompl1 = !CODCOMPLEMENTO
            sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
            
            Sql = "select sum(valortributo) as soma from debitotributo where codreduzido=" & nCodReduz & " and anoexercicio=" & nAno1 & " and "
            Sql = Sql & "codlancamento=" & nLanc1 & " and seqlancamento=" & nSeq1 & " and numparcela=" & nParc1 & " and codcomplemento=" & nCompl1 & " and "
            Sql = Sql & "codtributo<>3"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            nValor1 = Round(RdoAux2!soma, 2)
            RdoAux2.Close
            
            'Localizar estes valores na matriz de origem
            bFind = False
            For Y = 1 To UBound(aOrigem)
                With aOrigem(Y)
                    If .nAno = nAno1 And .nLanc = nLanc1 And .nParc = nParc1 And .nCompl = nCompl1 And .sDataVencto = sDataVencto And .nValorPrincipal = nValor1 Then
                        bFind = True
                        Sql = "insert transfere_debito (codigo1,ano1,lanc1,seq1,parc1,comp1,datavencto1,valor1,codigo2,ano2,lanc2,seq2,parc2,comp2,datavencto2,valor2,statuslanc) "
                        Sql = Sql & "values(" & 38258 & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & ",'" & Format(.sDataVencto, "mm/dd/yyyy") & "',"
                        Sql = Sql & Virg2Ponto(CStr(.nValorPrincipal)) & "," & nCodReduz & "," & nAno1 & "," & nLanc1 & "," & nSeq1 & "," & nParc1 & "," & nCompl1 & ",'"
                        Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(nValor1)) & "," & RdoAux!statuslanc & ")"
                        cn.Execute Sql, rdExecDirect
                        
                        Exit For
                    End If
                End With
            Next
           
            If Not bFind Then
                MsgBox "não achei"
            End If
           
           
            DoEvents
           .MoveNext
        Loop
        RdoAux.Close
    End With
Next

Etapa2:

Sql = "SELECT transfere_debito.codigo1, transfere_debito.ano1, transfere_debito.lanc1, transfere_debito.seq1, transfere_debito.parc1, transfere_debito.comp1,transfere_debito.datavencto1, transfere_debito.valor1,"
Sql = Sql & "transfere_debito.codigo2, transfere_debito.ano2, transfere_debito.lanc2, transfere_debito.seq2,transfere_debito.parc2, transfere_debito.comp2, transfere_debito.datavencto2, transfere_debito.valor2, transfere_debito.statuslanc,"
Sql = Sql & "debitoparcela.statuslanc AS sit2 FROM transfere_debito INNER JOIN debitoparcela ON transfere_debito.codigo1 = debitoparcela.codreduzido AND transfere_debito.ano1 = debitoparcela.anoexercicio AND "
Sql = Sql & "transfere_debito.lanc1 = debitoparcela.codlancamento AND transfere_debito.seq1 = debitoparcela.seqlancamento AND transfere_debito.parc1 = debitoparcela.NumParcela And transfere_debito.comp1 = debitoparcela.CODCOMPLEMENTO "
Sql = Sql & "ORDER BY transfere_debito.codigo2, transfere_debito.ano2, transfere_debito.lanc2, transfere_debito.parc2"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCodigo1 = !codigo1: nAno1 = !ano1: nLanc1 = !lanc1: nSeq1 = !seq1: nParc1 = !parc1: nCompl1 = !comp1
        nCodigo2 = !codigo2: nAno2 = !ano2: nLanc2 = !lanc2: nSeq2 = !seq2: nParc2 = !parc2: nCompl2 = !comp2
        
        'ATUALIZA PARCELADOCUMENTO
        Sql = "UPDATE PARCELADOCUMENTO SET CODREDUZIDO=" & nCodigo2 & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute Sql, rdExecDirect
        'ATUALIZA DEBITOPAGO
        Sql = "UPDATE DEBITOPAGO SET CODREDUZIDO=" & nCodigo2 & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute Sql, rdExecDirect
        'ATUALIZA OBS
        Sql = "UPDATE obsparcela SET CODREDUZIDO=" & nCodigo2 & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute Sql, rdExecDirect
                
        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & nCodigo1 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno1 & " AND CODLANCAMENTO=" & nLanc1 & " AND SEQLANCAMENTO=" & nSeq1 & " AND NUMPARCELA=" & nParc1 & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl1
        cn.Execute Sql, rdExecDirect
        
        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & !SIT2 & " WHERE CODREDUZIDO=" & nCodigo2 & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno2 & " AND CODLANCAMENTO=" & nLanc2 & " AND SEQLANCAMENTO=" & nSeq2 & " AND NUMPARCELA=" & nParc2 & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl2
        cn.Execute Sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With


fim:
MsgBox "fim"

Exit Sub
Erro:
MsgBox rdoErrors(0).Description

End Sub

Private Sub cmdFase4_Click()

Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset
Dim nPagas As Integer


If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
Pb.value = 0: lblPB.Caption = "0 %": nPos = 1


Sql = "SELECT * from daf_reg order by codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 100 = 0 Then CallPb nPos, nTot
                
        Sql = "select count(*) as contador from debitopago where codreduzido=" & !CODREDUZIDO & " and "
        Sql = Sql & "datapagamento='" & Format(!datapagto, "mm/dd/yyyy") & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!contador) Then
            nPagas = 0
        Else
            nPagas = RdoAux2!contador
        End If
                
        Sql = "update daf_reg set pagas=" & nPagas & " where codreduzido=" & !CODREDUZIDO & " and "
        Sql = Sql & "datapagto='" & Format(!datapagto, "mm/dd/yyyy") & "'"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With



MsgBox "fim"

Exit Sub
Erro:
MsgBox rdoErrors(0).Description



End Sub

Private Sub cmdFase5_Click()
Dim Sql As String, RdoAux As rdoResultset
If NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Erro fatal."
    Exit Sub
End If
cmdFase5.Enabled = False

Sql = "DELETE FROM SIMPLESCNPJ"
cn.Execute Sql, rdExecDirect

Sql = "SELECT * FROM RESUMOARQSN WHERE nome='CNPJ NÃO LOCALIZADO' ORDER BY CNPJ"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
        Sql = Sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(!Cnpj) & "','" & !ArquivoShort & "'," & !Banco & ",'" & Format(!DataArrecada, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(!DataVencto, "mm/dd/yyyy") & "'," & !AnoComp & "," & !MesComp & "," & Virg2Ponto(!principal) & "," & Virg2Ponto(!Juros) & "," & Virg2Ponto(!Multa) & ",'"
        Sql = Sql & !Agencia & "'," & 0 & ")"
        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

MsgBox "FIM"
End Sub

Private Sub btCorrige_Click(Index As Integer)
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
'Restaura
'CorrigeProcesso
'BaixaEicon
NaoPagoParaPago
End Sub

Private Sub Restaura()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, nPc As Integer, nCp As Integer
Dim nTot As Long, nPos As Long, sNumProc As String, nNumproc As Long, nAnoproc As Integer

ConectaBkp

GoTo FASE8

FASE1:
Sql = "SELECT * from facequadra order by coddistrito,codsetor,codquadra,codface"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from facequadra where coddistrito=" & !CODDISTRITO & " and codsetor=" & !CODSETOR & " and codquadra=" & !CODQUADRA & " and codface=" & !CODFACE
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogr <> !CodLogr Then
                DoEvents
                Sql = "update facequadra set codlogr=" & !CodLogr & " where coddistrito=" & !CODDISTRITO & " and codsetor=" & !CODSETOR & " and codquadra=" & !CODQUADRA & " and codface=" & !CODFACE
                cn.Execute Sql, rdExecDirect
            End If
        Else
            Sql = "insert facequadra select coddistrito,codsetor,codquadra,codface,codlogr,codagrupa,pavimento,quadras from tributacaobkp..facequadra where "
            Sql = Sql & "coddistrito=" & !CODDISTRITO & " and codsetor=" & !CODSETOR & " and codquadra=" & !CODQUADRA & " and codface=" & !CODFACE
            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

FASE2:
Sql = "SELECT * from logradouro"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from logradouro where codlogradouro=" & !CodLogradouro
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!NomeLogradouro <> !NomeLogradouro Then
                DoEvents
         '       MsgBox "BKP:" & !CodLogradouro & "-" & !NomeLogradouro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!NomeLogradouro
 '               Sql = "update logradouro set codtipolog=" & Val(SubNull(!codtipolog)) & ",codtitlog=" & Val(SubNull(!codtitlog)) & ",nomelogradouro='" & Mask(!NomeLogradouro) & "',endereco='" & Mask(!Endereco) & "' where codlogradouro=" & !CodLogradouro
'                cn.Execute Sql, rdExecDirect
            End If
        Else
            'MsgBox "Não existe:" & !CodLogradouro & "-" & !NomeLogradouro
            Sql = "insert logradouro select codlogradouro,endereco,dataofic,numofic,codtipolog,codtitlog,nomelogradouro from tributacaobkp..logradouro where "
            Sql = Sql & "codlogradouro=" & !CodLogradouro
            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

FASE3:
Sql = "SELECT * from bairro where siglauf='SP' and codcidade=413 order by codbairro"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from bairro where siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!DescBairro <> !DescBairro Then
                DoEvents
'                MsgBox "BKP:" & !CodBairro & "-" & !DescBairro & " -> Atual:" & RdoAux2!CodBairro & "-" & RdoAux2!DescBairro
                Sql = "update bairro set descbairro='" & Mask(!DescBairro) & "' where siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
                cn.Execute Sql, rdExecDirect
            End If
        Else
 '           MsgBox "Não existe:" & !CodBairro & "-" & !DescBairro
            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


FASE4:
Sql = "SELECT * from cadimob order by codreduzido"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from cadimob where codreduzido=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!Li_CodBairro <> !Li_CodBairro Then
                DoEvents
              '  MsgBox "BKP:" & !CODREDUZIDO & "-" & !Li_CodBairro & " -> Atual:" & RdoAux2!CODREDUZIDO & "-" & RdoAux2!Li_CodBairro
                Sql = "update cadimob set li_codbairro=" & !Li_CodBairro & " where codreduzido=" & !CODREDUZIDO
                cn.Execute Sql, rdExecDirect
            End If
        Else
            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


FASE5:
Sql = "SELECT * from cidadao order by codcidadao"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from cidadao where codcidadao=" & !CodCidadao
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogradouro <> !CodLogradouro Or RdoAux2!CodBairro <> !CodBairro Then
                DoEvents
                If Val(SubNull(!CodLogradouro)) > 0 Then
                    'MsgBox !CodCidadao & " BKP:" & !CodLogradouro & "-" & !CodBairro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!CodBairro
                Sql = "update cidadao set codlogradouro=" & !CodLogradouro & ",codbairro=" & IIf(IsNull(!CodBairro), 999, !CodBairro) & " where codcidadao=" & !CodCidadao
                cn.Execute Sql, rdExecDirect
               End If
            End If
        Else
            MsgBox "Não existe:" & !CodCidadao
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
GoTo fim
FASE6:
Sql = "SELECT * from endentrega where ee_cidade=413 order by codreduzido"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from endentrega where ee_cidade=413 and codreduzido=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!Ee_CodLog <> !Ee_CodLog Or RdoAux2!Ee_Bairro <> !Ee_Bairro Then
                DoEvents
                If Val(SubNull(!Ee_CodLog)) > 0 Then
               '     MsgBox !CODREDUZIDO & " BKP:" & !Ee_CodLog & "-" & !Ee_Bairro & " -> Atual:" & RdoAux2!Ee_CodLog & "-" & RdoAux2!Ee_Bairro
                Sql = "update endentrega set ee_codlog=" & !Ee_CodLog & ",ee_bairro=" & IIf(IsNull(!Ee_Bairro), 999, !Ee_Bairro) & " where codreduzido=" & !CODREDUZIDO
                cn.Execute Sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

GoTo fim

FASE7:
Sql = "SELECT * from mobiliario order by codigomob"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from mobiliario where codigomob=" & !codigomob
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogradouro <> !CodLogradouro Or RdoAux2!CodBairro <> !CodBairro Then
                DoEvents
                If Val(SubNull(!CodLogradouro)) > 0 Then
                 '   MsgBox !CODigomob & " BKP:" & !CodLogradouro & "-" & !CodBairro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!CodBairro
                Sql = "update mobiliario set codlogradouro=" & !CodLogradouro & ",codbairro=" & IIf(IsNull(!CodBairro), 999, !CodBairro) & " where codigomob=" & !codigomob
                cn.Execute Sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
GoTo fim

FASE8:
Sql = "SELECT * from mobiliarioendentrega order by codmobiliario"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from mobiliarioendentrega where codmobiliario=" & !codmobiliario
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogradouro <> !CodLogradouro Or RdoAux2!CodBairro <> !CodBairro Then
                DoEvents
                If Val(SubNull(!CodLogradouro)) > 0 Then
                  '  MsgBox !CODmobiliario & " BKP:" & !CodLogradouro & "-" & !CodBairro & " -> Atual:" & RdoAux2!CodLogradouro & "-" & RdoAux2!CodBairro
                Sql = "update mobiliarioendentrega set codlogradouro=" & !CodLogradouro & ",codbairro=" & IIf(IsNull(!CodBairro), 999, !CodBairro) & " where codmobiliario=" & !codmobiliario
                cn.Execute Sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
GoTo fim
FASE9:
Sql = "SELECT * from processoend order by Ano,numprocesso"
Set RdoAux = cnBkp.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "select * from processoend where ano=" & !Ano & " and numprocesso=" & !numprocesso
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            If RdoAux2!CodLogr <> !CodLogr Then
                DoEvents
                On Error Resume Next
                If Val(SubNull(!CodLogr)) > 0 Then
        '            MsgBox !numprocesso & "/" & !Ano & " BKP:" & !CodLogr & " -> Atual:" & RdoAux2!CodLogr
                Sql = "update processoend set codlogr=" & !CodLogr & " where ano=" & !Ano & " and numprocesso=" & !numprocesso & " and numero=" & !Numero
  '              cn.Execute Sql, rdExecDirect
               End If
            End If
        Else
'            MsgBox "Não existe:" & !CODREDUZIDO
'            Sql = "insert bairro select siglauf,codcidade,codbairro,descbairro from tributacaobkp..bairro where "
'            Sql = Sql & "siglauf='SP' and codcidade=413 and codbairro=" & !CodBairro
'            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


fim:
MsgBox "fim"

End Sub


Private Sub CorrigeProcesso()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, nPc As Integer, nCp As Integer
Dim nTot As Long, nPos As Long, sNumProc As String, nNumproc As Long, nAnoproc As Integer


FASE1:
Sql = "SELECT distinct NUMPROCESSO FROM origemreparc WHERE ANOPROC IS NULL ORDER BY NUMPROCESSO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sNumProc = !numprocesso
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        Sql = "update origemreparc set numproc=" & nNumproc & ",anoproc=" & nAnoproc & " where numprocesso='" & sNumProc & "'"
        cn.Execute Sql, rdExecDirect

        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

FASE2:

Sql = "SELECT distinct NUMPROCESSO FROM destinoreparc WHERE ANOPROC IS NULL ORDER BY NUMPROCESSO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sNumProc = !numprocesso
        nNumproc = Val(Left$(sNumProc, InStr(1, sNumProc, "/", vbBinaryCompare) - 1))
        nAnoproc = Val(Right$(sNumProc, 4))
        
        Sql = "update destinoreparc set numproc=" & nNumproc & ",anoproc=" & nAnoproc & " where numprocesso='" & sNumProc & "'"
        cn.Execute Sql, rdExecDirect

        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


fim:
MsgBox "fim"

End Sub


Private Sub cmdGravar_Click()
Dim RdoAux As rdoResultset, x As Integer

s = Mid(lstParam.Text, 2, 6)

If txtOld.Text = txtValor.Text Then
    MsgBox "Nenhuma alteração foi feita neste parâmetro.", vbInformation, "Atenção"
    Exit Sub
Else
    If MsgBox("Deseja alterar o Valor de " & txtOld.Text & " para " & txtValor.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "UPDATE PARAMETROS SET VALPARAM='" & txtValor.Text & "' WHERE NOMEPARAM='" & s & "'"
        cn.Execute Sql, rdExecDirect
        txtOld.Text = txtValor.Text
    End If
End If

End Sub


Private Sub cmdPagos_Click()

Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, RdoAux3 As rdoResultset
Dim nPc As Integer, nCp As Integer, sCnae As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String, nNumDoc As Long
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

BaixaEicon
'CancelaUnica
'EmiteBoleto
Exit Sub

frmComercioEletronico.BoletoNome = "DANIELA LAURENTIZ TERRA"
frmComercioEletronico.BoletoCidade = "CENTRO"
frmComercioEletronico.BoletoCep = "14887-888"
frmComercioEletronico.BoletoCpfCnpj = "151.729.278-67"
frmComercioEletronico.BoletoEndereco = "AV. MARECHAL DEODORO, 573 BLOCO A"
frmComercioEletronico.BoletoNumDoc = 15712545
frmComercioEletronico.BoletoUF = "PR"
frmComercioEletronico.BoletoValor = 12658.3
frmComercioEletronico.BoletoVencto = "17/01/2018"
frmComercioEletronico.show 1
'CorrigeRefis
'NaoPagoParaPago
'MsgBox "fim"


Exit Sub


Sql = "truncate table taxa_lixo"
cn.Execute Sql, rdExecDirect

Sql = "SELECT areas.codreduzido, SUM(areas.areaconstr) AS soma, vwFULLIMOVEL.INSCRICAO, vwFULLIMOVEL.LOGRADOURO,    vwFULLIMOVEL.Li_Num , vwFULLIMOVEL.Li_Compl, vwFULLIMOVEL.DescBairro, vwFULLIMOVEL.CodLogr "
Sql = Sql & "FROM areas INNER JOIN vwFULLIMOVEL ON areas.codreduzido = vwFULLIMOVEL.codreduzido Where (vwFULLIMOVEL.Inativo = 0) GROUP BY areas.codreduzido, vwFULLIMOVEL.INSCRICAO, vwFULLIMOVEL.LOGRADOURO, vwFULLIMOVEL.li_num, vwFULLIMOVEL.li_compl,"
Sql = Sql & "vwFULLIMOVEL.DescBairro , vwFULLIMOVEL.CodLogr ORDER BY areas.codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        On Error Resume Next
        Sql = "select TOP (1) areas.codreduzido, usoconstr.descusoconstr FROM areas INNER JOIN usoconstr ON areas.usoconstr = usoconstr.codusoconstr Where CODREDUZIDO = " & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        Sql = "insert taxa_lixo (codigo_imovel,inscricao,codigo_logradouro,endereco,numero,complemento,bairro,cep,area,uso) values(" & !CODREDUZIDO & ",'" & !Inscricao & "'," & !CodLogr & ",'" & !Logradouro & "',"
        Sql = Sql & !Li_Num & ",'" & Left(SubNull(!Li_Compl), 50) & "','" & !DescBairro & "','" & RetornaCEP(!CodLogr, !Li_Num) & "'," & Virg2Ponto(CStr(Round(!soma, 2))) & ",'" & RdoAux2!descusoconstr & "')"
        cn.Execute Sql, rdExecDirect

PROXIMO2:
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With
'PrintExcel
fim:
MsgBox "fim"
End Sub


Private Sub Form_Load()

'File1.Path = "c:\trabalho\daf\Arq1\"
Centraliza Me
With lstParam
    .AddItem "(SEQ237) Sequência Arquivo DA Bradesco"
    .AddItem "(SEQ341) Sequência Arquivo DA Itaú"
    .AddItem "(SEQ409) Sequência Arquivo DA Unibanco"
    .AddItem "(SEQ033) Sequência Arquivo DO Banespa"
    .AddItem "(SEQ399) Sequência Arquivo DO HSBC"
End With
lstParam.ListIndex = 0

End Sub

Private Sub lstParam_Click()
Dim s As String, Sql As String, RdoAux As rdoResultset

If lstParam.ListIndex = -1 Then Exit Sub
s = Mid(lstParam.Text, 2, 6)

Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='" & s & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtValor.Text = !valparam
    Else
        txtValor.Text = "Não Cadastrado"
    End If
   .Close
End With
txtOld.Text = txtValor.Text

End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

   On Error GoTo FillSpace_Error

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

   On Error GoTo 0
   Exit Function

FillSpace_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FillSpace of Formulário frmConfig"

End Function

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

'Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function

Public Function SNCheck(nCodigo As Long) As Boolean
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNASN(" & Format(nCodigo, "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
     If RdoAux!RETORNO = 1 Then
        SNCheck = True
     Else
        SNCheck = False
     End If
    .Close
End With

End Function

Private Sub simples_ano()
Sql = "delete from simples_ano"
cn.Execute Sql, rdExecDirect
ReDim aAno(0)
Sql = "SELECT * from periodosn order by codigo,dataini"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        
        nCodReduz = !Codigo
        nAnoIni = Year(!dataini)
        If Not IsNull(!datafim) Then
            nAnoFim = Year(!datafim)
             
            For Y = nAnoIni To nAnoFim
                nAnoTmp = Y
                GoSub AddMatrix
            Next
        Else
            nAnoTmp = nAnoIni
            GoSub AddMatrix
        End If
        
       .MoveNext
        DoEvents
    Loop
   .Close
   
    For x = 1 To UBound(aAno)
        Sql = "insert simples_ano (codigo,ano) values(" & aAno(x).nCodigo & "," & aAno(x).nAno & ")"
        cn.Execute Sql, rdExecDirect
    Next
   
 MsgBox "fim"
Exit Sub

AddMatrix:
    bFind = False
    For x = 0 To UBound(aAno)
        If aAno(x).nCodigo = nCodReduz And aAno(x).nAno = nAnoTmp Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        ReDim Preserve aAno(UBound(aAno) + 1)
        aAno(UBound(aAno)).nCodigo = nCodReduz
        aAno(UBound(aAno)).nAno = nAnoTmp
    End If
    Return
   
End With
End Sub

Private Sub PrintExcel()

If lvMain.ListItems.Count = 0 Then Exit Sub

Dim x As Long, Y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For Y = 1 To lvMain.ColumnHeaders.Count
    ax = ax & RemoveSpace(lvMain.ColumnHeaders(Y).Text) & " char(255), "
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "Table1$", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To lvMain.ListItems.Count
    Rs.AddNew
    nCont = 0
    Rs.Fields(nCont).value = lvMain.ListItems(x).Text
    nCont = nCont + 1
    For Y = 2 To lvMain.ColumnHeaders.Count
         
         Rs.Fields(nCont).value = lvMain.ListItems(x).SubItems(Y - 1)
         nCont = nCont + 1
    
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)


End Sub

Private Sub LeArquivo(sFullPath As String, sArq As String, nCodBanco As Integer, sDataCredito As String)

Dim sReg As String, FF1 As Integer, bExec As Boolean, sTipoArq As String, kk As Integer
Dim nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorGuia As Double, nNumDoc As Long, nErro As Integer, RdoAux4 As rdoResultset
Dim sAno As String, sMes As String, sAgencia As String, bLayoutNovo As Boolean, RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, RdoAux3 As rdoResultset
Dim nNumParc As Integer, bAchou As Boolean, nSeq As Integer, nCompl As Integer, nCodReduz As Long, sDataVencto As String, nRetorno As Integer, sRetorno As String
Dim nValorEfetivo As Double, nSeqReg As Integer, itmX As ListItem, nValorTaxa As Double, R As Integer, sDataGeracao As String, sLinhaT As String, sLinhaU As String, aRegistro() As Registro, aDoc() As Documento


ReDim aRegistro(0): ReDim aDoc(0)
nSeqReg = 1

'*** VERIFICA EXISTENCIA DO ARQUIVO

sFullPath = Replace(sFullPath, "/", "\")

If Dir$(sFullPath) = "" Then
    MsgBox "Não localizado o arquivo em " & sFullPath, vbCritical, "ERRO FATAL !!!"
    Exit Sub
End If


Ocupado

sReg = ""

'*****************************************
'****** ARQUIVO DO SIMPLES NACIONAL ******
'*****************************************
FF1 = FreeFile()
Open sFullPath For Binary Access Read Write As FF1

    While Not EOF(FF1)
        On Error GoTo CloseFile2
        bExec = False
        If Left(sReg, 1) = "9" Then GoTo CloseFile2
        Input #FF1, sReg
        If Left(sReg, 1) = "1" Then
            sSeqArq = Mid(sReg, 2, 8)
        ElseIf Left(sReg, 1) = "2" Then
           'LE OS REGISTROS
            With grdReg
                nValorPrincipal = CDbl(Mid(sReg, 107, 17)) / 100
                nValorJuros = CDbl(Mid(sReg, 124, 17)) / 100
                nValorMulta = CDbl(Mid(sReg, 141, 17)) / 100
                nValorGuia = nValorPrincipal + nValorJuros + nValorMulta
                sAno = Mid(sReg, 101, 4)
                sMes = Mid(sReg, 105, 2)
                sAgencia = Mid(sReg, 223, 4)
                sCNPJ = Mid(sReg, 75, 14)
                
                nSeq = 0
                bAchou = False
                For R = 1 To UBound(aRegistro)
                    If aRegistro(R).sCNPJ = Mid(sReg, 75, 14) Then
                        bAchou = True
                        nSeq = nSeq + 1
                    End If
                Next
                
                ReDim Preserve aRegistro(UBound(aRegistro) + 1)
                aRegistro(UBound(aRegistro)).sDataVencto = ConvDataSerial(Mid(sReg, 18, 8))
                aRegistro(UBound(aRegistro)).sDataCred = ConvDataSerial(Mid(sReg, 10, 8))
                aRegistro(UBound(aRegistro)).sDataPag = ConvDataSerial(Mid(sReg, 10, 8))
                aRegistro(UBound(aRegistro)).nValorPago = nValorGuia
                aRegistro(UBound(aRegistro)).sAgencia = Mid(sReg, 223, 4)
                aRegistro(UBound(aRegistro)).sCNPJ = Mid(sReg, 75, 14)
                aRegistro(UBound(aRegistro)).nAno = Val(Mid(sReg, 101, 4))
                aRegistro(UBound(aRegistro)).nMes = Val(Mid(sReg, 105, 2))
                aRegistro(UBound(aRegistro)).nValorTarifaBancaria = 0
                aRegistro(UBound(aRegistro)).sSitRetorno = "CNPJ: " & Format(aRegistro(UBound(aRegistro)).sCNPJ, "0#\.###\.###/####-##")
                aRegistro(UBound(aRegistro)).bExiste = True
                aRegistro(UBound(aRegistro)).sDataPagCalc = RetornaDiaUtil(CDate(aRegistro(UBound(aRegistro)).sDataPag))
                If Not bAchou Then
                    aRegistro(UBound(aRegistro)).nSeq = 0
                Else
                    aRegistro(UBound(aRegistro)).nSeq = nSeq
                End If
                
                With aRegistro(UBound(aRegistro))
                    
                    
                    Set itmX = lvMain.ListItems.Add(, , sFullPath)
                    itmX.SubItems(1) = sArq
                    itmX.SubItems(2) = nCodBanco
                    itmX.SubItems(3) = sDataCredito
                    itmX.SubItems(4) = .sCNPJ
                    itmX.SubItems(6) = .nAno
                    itmX.SubItems(7) = .nMes
'                    itmX.SubItems(5) = !CODREDUZIDO
                    itmX.SubItems(8) = aRegistro(UBound(aRegistro)).sDataVencto
                    itmX.SubItems(9) = nValorGuia

                    
                End With
                'PROCURA SE O DEBITO JA FOI BAIXADO
                Sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE ARQUIVOBANCO='" & sArq & "' AND DATACREDITO='" & Format(ConvDataSerial(Mid(sReg, 10, 8)), "mm/dd/yyyy") & "' AND "
                Sql = Sql & "CNPJ='" & Mid(sReg, 75, 14) & "' AND ANO=" & Val(Mid(sReg, 101, 4)) & " AND MES=" & Val(Mid(sReg, 105, 2))
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        'CARREGA PARCELA GRAVADA
                        ReDim Preserve aDoc(UBound(aDoc) + 1)
                        aDoc(UBound(aDoc)).sCNPJ = aRegistro(UBound(aRegistro)).sCNPJ
                        aDoc(UBound(aDoc)).nCodReduz = !CODREDUZIDO
                        aDoc(UBound(aDoc)).nAno = !AnoExercicio
                        aDoc(UBound(aDoc)).nLanc = !CodLancamento
                        aDoc(UBound(aDoc)).nSeq = !SeqLancamento
                        aDoc(UBound(aDoc)).nParc = !NumParcela
                        aDoc(UBound(aDoc)).nCompl = !CODCOMPLEMENTO
                        aDoc(UBound(aDoc)).sDataVencto = aRegistro(UBound(aRegistro)).sDataVencto
                        aDoc(UBound(aDoc)).sSit = 2
                        aDoc(UBound(aDoc)).nValorPrincipal = nValorPrincipal
                        aDoc(UBound(aDoc)).nValorMulta = nValorMulta
                        aDoc(UBound(aDoc)).nValorJuros = nValorJuros
                        aDoc(UBound(aDoc)).nValorCorrecao = 0
                        aDoc(UBound(aDoc)).nValorTotal = nValorGuia
                        aDoc(UBound(aDoc)).nValorTarifa = 0
                        aDoc(UBound(aDoc)).nValorDif = 0
                        aDoc(UBound(aDoc)).nValorCompensado = nValorGuia
                        aDoc(UBound(aDoc)).sBx = "S"
                        aDoc(UBound(aDoc)).sDp = "N"
                        aDoc(UBound(aDoc)).bExiste = True
                        aDoc(UBound(aDoc)).nSeqReg = aRegistro(UBound(aRegistro)).nSeq
                    Else
                        'DEFINIR NOVA PARCELA
                        'BUSCA CÓDIGO
                        Sql = "SELECT CODIGOMOB,CNPJ FROM MOBILIARIO WHERE DATAENCERRAMENTO IS NULL and CONVERT(BIGINT, cnpj) = " & Val(aRegistro(UBound(aRegistro)).sCNPJ)
                        Sql = Sql & " OR CNPJ='" & Format(aRegistro(UBound(aRegistro)).sCNPJ, "00\.000\.000/0000-00") & "' AND DATAENCERRAMENTO IS NULL ORDER BY CODIGOMOB DESC"
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                nCodReduz = !codigomob
                                .Close
                            Else
                                .Close
                                Sql = "SELECT CODCIDADAO,CNPJ FROM CIDADAO WHERE CNPJ = '" & RetornaNumero(aRegistro(UBound(aRegistro)).sCNPJ) & "' OR "
                                Sql = Sql & "CNPJ='" & Format(aRegistro(UBound(aRegistro)).sCNPJ, "00\.000\.000/0000-00") & "' ORDER BY CODCIDADAO DESC"
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    If .RowCount > 0 Then
                                        nCodReduz = !CodCidadao
                                    Else
                                        'CNPJ NÃO LOCALIZADO
                                        aRegistro(UBound(aRegistro)).bExiste = False
                                        Sql = "SELECT * FROM SIMPLESCNPJ WHERE CNPJ='" & aRegistro(UBound(aRegistro)).sCNPJ & "' AND ANOCOMP=" & aRegistro(UBound(aRegistro)).nAno & " AND MESCOMP=" & aRegistro(UBound(aRegistro)).nMes
                                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                        If RdoAux3.RowCount = 0 Then
                                            Sql = "INSERT SIMPLESCNPJ (CNPJ,ARQUIVOSHORT,BANCO,DATAARRECADA,DATAVENCTO,ANOCOMP,MESCOMP,PRINCIPAL,JUROS,"
                                            Sql = Sql & "MULTA,AGENCIA,CODREDUZIDO) VALUES('" & RetornaNumero(aRegistro(UBound(aRegistro)).sCNPJ) & "','" & lstArq.Text & "'," & Val(Left(lblBanco.Caption, 3)) & ",'" & Format(aRegistro(UBound(aRegistro)).sDataCred, "mm/dd/yyyy") & "','"
                                            Sql = Sql & Format(aRegistro(UBound(aRegistro)).sDataVencto, "mm/dd/yyyy") & "'," & aRegistro(UBound(aRegistro)).nAno & "," & aRegistro(UBound(aRegistro)).nMes & "," & Virg2Ponto(CStr(aRegistro(UBound(aRegistro)).nValorPago)) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & ",'"
                                            Sql = Sql & aRegistro(UBound(aRegistro)).sAgencia & "'," & 0 & ")"
                      '                      cn.Execute Sql, rdExecDirect
                                        End If
                                        RdoAux3.Close
                                        GoTo CONTSN
                                    End If
                                    
                                    
                                    
                                End With
                            End If
                           
                        End With
                                
                        'BUSCA LANCAMENTO
                         Sql = "SELECT debitoparcela.codreduzido,debitoparcela.anoexercicio, debitoparcela.codlancamento,DEBITOPARCELA.SEQLANCAMENTO,debitoparcela.numparcela,DEBITOPARCELA.CODCOMPLEMENTO,debitoparcela.datavencimento, debitoparcela.statuslanc, debitotributo.valortributo "
                         Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                         Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And "
                         Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.codlancamento = 5) AND (MONTH(debitoparcela.datavencimento) = " & Month(CDate(aRegistro(UBound(aRegistro)).sDataVencto)) & ") AND "
                         Sql = Sql & "(YEAR(debitoparcela.datavencimento) = " & Year(CDate(aRegistro(UBound(aRegistro)).sDataVencto)) & ") AND (debitotributo.codtributo = 13) and debitotributo.valortributo =" & Virg2Ponto(CStr(nValorGuia)) & " AND statuslanc<>6"
                         Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                         With RdoAux3
                            'EXISTE LANCAMENTO NESTE MÊS/ANO?
                             If .RowCount > 0 Then 'SIM
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
                                 Else
                                    'SE NÃO ACHAR
                                    .MoveFirst
                                     nCompl = 0
                                    'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
                                     Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (MONTH(datavencimento) = " & Month(dDataVencto) & ") AND "
                                     Sql = Sql & "(YEAR(datavencimento) = " & Year(dDataVencto) & ")"
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
                                 Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND (codlancamento = 5) AND (ANOEXERCICIO = " & Val(sAno) & ")"
                                 Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                 With RdoAux4
                                     If IsNull(!maximo) Then
                                         nSeq = 1
                                     Else
                                         nSeq = !maximo + 1
                                     End If
                                    .Close
                                 End With
                                 'VERIFICA SE A SEQ JA NÃO EXISTE NA MATRIZ
                                 For R = 1 To UBound(aDoc)
                                    If aDoc(R).nCodReduz = nCodReduz And aDoc(R).nAno = Val(sAno) Then
                                        nSeq = aDoc(R).nSeq + 1
                                    End If
                                 Next
                                 
                                 nCompl = 0
                                 nNumParc = 1
                             End If
                             ReDim Preserve aDoc(UBound(aDoc) + 1)
                             aDoc(UBound(aDoc)).sCNPJ = aRegistro(UBound(aRegistro)).sCNPJ
                             aDoc(UBound(aDoc)).nCodReduz = nCodReduz
                             aDoc(UBound(aDoc)).nAno = Val(sAno)
                             aDoc(UBound(aDoc)).nLanc = 5
                             aDoc(UBound(aDoc)).nSeq = nSeq
                             aDoc(UBound(aDoc)).nParc = nNumParc
                             aDoc(UBound(aDoc)).nCompl = nCompl
                             aDoc(UBound(aDoc)).sDataVencto = aRegistro(UBound(aRegistro)).sDataVencto
                             aDoc(UBound(aDoc)).sSit = 3
                             aDoc(UBound(aDoc)).nValorPrincipal = nValorPrincipal
                             aDoc(UBound(aDoc)).nValorMulta = nValorMulta
                             aDoc(UBound(aDoc)).nValorJuros = nValorJuros
                             aDoc(UBound(aDoc)).nValorCorrecao = 0
                             aDoc(UBound(aDoc)).nValorTotal = nValorGuia
                             aDoc(UBound(aDoc)).nValorTarifa = 0
                             aDoc(UBound(aDoc)).nValorDif = 0
                             aDoc(UBound(aDoc)).nValorCompensado = nValorGuia
                             aDoc(UBound(aDoc)).sBx = ""
                             aDoc(UBound(aDoc)).sDp = ""
                             aDoc(UBound(aDoc)).bExiste = True
                             aDoc(UBound(aDoc)).nSeqReg = aRegistro(UBound(aRegistro)).nSeq
                            itmX.SubItems(5) = nCodReduz
                            itmX.SubItems(10) = Val(sAno)
                            itmX.SubItems(11) = nSeq
                            Sql = "SELECT debitoparcela.codreduzido,debitotributo.valortributo From  debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
                            Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
                            Sql = Sql & "WHERE debitoparcela.codreduzido = " & nCodReduz & " AND debitoparcela.anoexercicio = " & Val(sAno) & " AND (debitoparcela.codlancamento = 5) and codtributo=13 AND datavencimento = '" & Format(CDate(aRegistro(UBound(aRegistro)).sDataVencto), "mm/dd/yyyy") & "' and valortributo=" & Virg2Ponto(CStr(nValorGuia))
                            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            If RdoAux4.RowCount > 0 Then
                                itmX.SubItems(12) = "S"
                            Else
                                itmX.SubItems(12) = "N"
                            End If
                            RdoAux4.Close
                            .Close
                         End With
CONTSN:
'**********************************
                    End If
                   .Close
                End With
               
            End With
        ElseIf Left(sReg, 1) = "9" Then
           'LE O RODAPÉ DO ARQUIVO
'            lblNumReg.Caption = Format(Val(Mid(sReg, 10, 6)) - 2, "000000")
'            lblValorTotal.Caption = FormatNumber(CDbl(Mid(sReg, 16, 17) / 100), 2)
        End If
         
        
        'nPos = nPos + 1
    Wend
CloseFile2:
Close #FF1
'Pb.value = 0
Liberado
nErro = 0




End Sub


Private Sub cmdPagosold_Click()

Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer
Dim nPc As Integer, nCp As Integer, sEvento As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub


GoTo Rotina2
Sql = "SELECT DISTINCT seq, datahoraevento, computador, usuario, form, evento, secevento, logevento "
Sql = Sql & "From logevento WHERE (evento = 3) AND (secevento = 2) order by seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sEvento = !LOGEVENTO
        nIni = InStr(sEvento, " em ")
        If nIni = 0 Then GoTo proximo
        nFim = InStr(sEvento, " Ano:")
        If nFim = 0 Then GoTo proximo
        nCodReduz = Val(Mid(sEvento, nIni + 4, nFim - nIni - 4))
        
        If nCodReduz < 500000 Or nCodReduz > 700000 Then GoTo proximo
        sMotivo = Left(sEvento, nIni - 1)
        
        nAno = Val(Mid(sEvento, InStr(sEvento, "Ano:") + 4, 4))
        nLc = Val(Mid(sEvento, InStr(sEvento, "Lc:") + 3, 2))
        nSq = Val(Mid(sEvento, InStr(sEvento, "Sq:") + 3, 2))
        nPc = Val(Mid(sEvento, InStr(sEvento, "Pc:") + 3, 2))
        nCp = Val(Mid(sEvento, InStr(sEvento, "Cp:") + 3, 2))
        sEvento = sMotivo & " - Ex:" & nAno & " Lc:" & nLc & " Sq:" & nSq & " Pc:" & nPc & " Cp:" & nCp
        
        Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCodReduz & ",'" & Format(!DATAHORAEVENTO, sDataFormat & " hh:mm:ss") & "','" & !Usuario & "','" & sEvento & "')"
        'cn.Execute Sql, rdExecDirect

proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With

Rotina2:
Sql = "SELECT DISTINCT seq, datahoraevento, computador, usuario, form, evento, secevento, logevento "
Sql = Sql & "From logevento WHERE (evento = 3) AND (secevento = 3) order by seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rydConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        sEvento = !LOGEVENTO
        nIni = InStr(sEvento, "Código:")
        If nIni = 0 Then GoTo PROXIMO2
        nFim = InStr(sEvento, "Lançamento:")
        If nFim = 0 Then GoTo PROXIMO2
        nCodReduz = Val(Mid(sEvento, InStr(sEvento, "Código:") + 7, 6))
        'nCodReduz = Val(Mid(sEvento, nIni + 4, nFim - nIni - 4))
        
        If nCodReduz < 500000 Or nCodReduz > 700000 Then GoTo PROXIMO2
        'sMotivo = Left(sEvento, nIni - 1)
        sMotivo = Mid(sEvento, InStr(sEvento, "Motivo:") + 8, Len(sEvento) - InStr(sEvento, "Motivo:"))
        nAno = Val(Mid(sEvento, InStr(sEvento, "Ano:") + 4, 4))
        nLc = Val(Mid(sEvento, InStr(sEvento, "Lançamento:") + 11, 2))
        nSq = Val(Mid(sEvento, InStr(sEvento, "Seq:") + 4, 2))
        nPc = Val(Mid(sEvento, InStr(sEvento, "Parcela:") + 8, 2))
        nCp = Val(Mid(sEvento, InStr(sEvento, "Compl:") + 5, 2))
        sEvento = sMotivo & " - Ex:" & nAno & " Lc:" & nLc & " Sq:" & nSq & " Pc:" & nPc & " Cp:" & nCp
        
        'Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCodReduz & ",'" & Format(!DATAHORAEVENTO, sDataFormat & " hh:mm:ss") & "','" & !Usuario & "','" & Mask(sEvento) & "')"
        Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCodReduz & ",'" & Format(!DATAHORAEVENTO, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(!Usuario) & ",'" & Mask(sEvento) & "')"
        cn.Execute Sql, rdExecDirect

PROXIMO2:
        nPos = nPos + 1
       .MoveNext
    Loop
    
   .Close
End With
'PrintExcel

MsgBox "fim"
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
Dim nPagina As Integer, nLivro As Integer, sDataDam As String, xImovel As clsImovel


'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
cn.Execute Sql, rdExecDirect

Sql = "delete from boletoguiacapa where sid=" & nSid
cn.Execute Sql, rdExecDirect

sLib = "LIBERACAO"


'sNumProc = lblNumProc.Caption & "/" & lblAnoProc.Caption
'dDataProc = lblDataParc.Caption
Sql = "SELECT cadimob.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento, parceladocumento.numparcela, "
Sql = Sql & "parceladocumento.CODCOMPLEMENTO , parceladocumento.NumDocumento, debitoparcela.DataVencimento, debitotributo.ValorTributo FROM cadimob INNER JOIN "
Sql = Sql & "parceladocumento ON cadimob.codreduzido = parceladocumento.codreduzido INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN "
Sql = Sql & "debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.NumParcela = debitotributo.NumParcela And debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO "
'Sql = Sql & "Where (cadimob.codreduzido in (35654,35565,35566)) "
Sql = Sql & "Where (cadimob.li_codbairro =1069) "
Sql = Sql & " And (parceladocumento.AnoExercicio = 2018) And (parceladocumento.CodLancamento =1) And (parceladocumento.SeqLancamento = 0) "
Sql = Sql & "AND  statuslanc=3 ORDER BY cadimob.codreduzido, parceladocumento.numparcela"

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
    '**** GERADOR DE CÓDIGO DE BARRAS ********
    sNossoNumero = "2678478"
    sDigitavel = "001900000"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv
    
    dDataBase = "07/10/1997"
    nFatorVencto = CDate(sDataDam) - CDate(dDataBase)
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
    
    '*******************************************

        Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numproc) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 0, nParc) & "," & 12 & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & "," & "'','')"
        cn.Execute Sql, rdExecDirect
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport2 "boletoguia2", frmMdi.hwnd, Me.hwnd, nSid
Liberado

End Sub

Private Function SNCheck2(nCodigo As Long) As Boolean
Dim RdoAux As rdoResultset, Sql As String, sReturn As Boolean
ConectaEicon
Sql = "select * from  tb_inter_empr_snacional_giss Where NUM_CADASTRO=" & nCodigo & " order by timestamp desc"
Set RdoAux = cnEicon.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    If RdoAux.RowCount > 0 Then
        If IsNull(!data_fim) Then
            sReturn = True
        Else
            sReturn = False
        End If
     Else
        sReturn = False
     End If
    .Close
End With
cnEicon.Close
SNCheck2 = sReturn
End Function

Private Sub TransfereLancamento()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, x As Integer
Dim nCodReduz As Long, nFicha As Long, sProc As String, nAno As Integer, nNumero As Long, sStatus As String, bCancelado As Boolean
Dim sCep As String

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
Pb.value = 0
nPos = 1

Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento From debitotributo Where CodTributo = 527 AND CODLANCAMENTO=11 ORDER BY codreduzido, anoexercicio, numparcela"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "update debitoparcela set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        Sql = "update debitotributo set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        Sql = "update parceladocumento set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        Sql = "update debitopago set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        Sql = "update obsparcela set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        Sql = "update origemreparc set codlancamento=81 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=11 and "
        Sql = Sql & "numsequencia=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
    RdoAux.Close
End With

MsgBox "fim"

Exit Sub
Erro:
MsgBox rdoErrors(0).Description

End Sub

Private Sub NaoPagoParaPago()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long
Pb.value = 0
nPos = 1

Sql = "SELECT DISTINCT  debitopago.codreduzido, debitopago.contacorrente, debitoparcela.statuslanc, debitopago.anoexercicio, debitopago.seqlancamento, debitopago.numparcela,"
Sql = Sql & "debitopago.CODCOMPLEMENTO , debitopago.CodLancamento FROM  debitopago INNER JOIN debitoparcela ON debitopago.codreduzido = debitoparcela.codreduzido AND debitopago.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "debitopago.codlancamento = debitoparcela.codlancamento AND debitopago.seqlancamento = debitoparcela.seqlancamento AND "
Sql = Sql & "debitopago.NumParcela = debitoparcela.NumParcela And debitopago.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO Where (debitoparcela.statuslanc = 3)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        Sql = "update debitoparcela set statuslanc=2 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

End Sub


Private Sub BaixaEicon()
Dim Sql As String, RdoAux As rdoResultset

ConectaEicon

Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, seqpag, datapagamento, datarecebimento, valorpago,CodBanco,"
Sql = Sql & "CodAgencia, restituido, NumDocumento, valorpagoreal, intacto, ValorTarifa, arquivobanco, valordif, datapagamentocalc, dataintegracao, contacorrente "
Sql = Sql & "From debitopago WHERE (numdocumento BETWEEN 2000000 AND 2100000) AND (numdocumento NOT IN (SELECT num_documento FROM GTI_Eicon.dbo.tb_inter_baixa_detalhe)) "
Sql = Sql & " AND (anoexercicio > 2015) ORDER BY numdocumento"

'Sql = "SELECT codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, seqpag, datapagamento, datarecebimento, valorpago, codbanco, "
'Sql = Sql & "CodAgencia , restituido, NumDocumento, valorpagoreal, intacto, ValorTarifa, arquivobanco, valordif, datapagamentocalc, dataintegracao, contacorrente "
'Sql = Sql & "From debitopago WHERE (codreduzido BETWEEN 100000 AND 300000) AND (datapagamento BETWEEN '03/01/2017' AND '03/31/2017') AND (codlancamento = 5)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        DoEvents
        '***** GRAVA BAIXA NA GISS ***************
        Sql = "insert tb_inter_baixa(cod_cliente,cod_banco,num_sequencia,timestamp,data_geracao,nome_arquivo,data_movimento) values("
        Sql = Sql & 2177 & "," & !CodBanco & "," & 0 & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "','" & Format(Now, "mm/dd/yyyy") & "','"
        Sql = Sql & !arquivobanco & "','" & Format(!datarecebimento, "mm/dd/yyyy") & "')"
        cnEicon.Execute Sql, rdExecDirect
        
        Sql = "insert tb_inter_baixa_detalhe(cod_cliente,cod_banco,num_sequencia,num_documento,linha,timestamp,valor_titulo,valor_pago,data_pagamento,valor_encargos,"
        Sql = Sql & "descricao_linha_t,descricao_linha_u) values(" & 2177 & "," & !CodBanco & "," & 0 & "," & !NumDocumento & "," & !SEQPAG & ",'" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "',"
        Sql = Sql & Virg2Ponto(CStr(!valorpagoreal)) & "," & Virg2Ponto(CStr(!valorpagoreal)) & ",'" & Format(!DataPagamento, "mm/dd/yyyy") & "'," & 0 & ",'"
        Sql = Sql & "" & "','" & "" & "')"
        cnEicon.Execute Sql, rdExecDirect
    
    
       .MoveNext
    Loop
   .Close
End With
                   
End Sub

Private Sub ISSpagoPorDAM()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nAno As Integer, nLc As Integer, nSq As Integer, RdoAux3 As rdoResultset
Dim nPc As Integer, nCp As Integer, sCnae As String, nPos As Long, nTot As Long, nIni As Integer, nFim As Integer, sMotivo As String, nNumDoc As Long
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

Sql = "SELECT * From debitopago WHERE (anoexercicio > 2014) AND (codlancamento = 5) AND (numdocumento > 4000000) AND (codbanco NOT IN (90, 91, 92, 93, 94, 95, 96, 97, 98, 99)) AND "
Sql = Sql & " (seqpag = 0) AND (codcomplemento = 0) order by numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nNumDoc = !NumDocumento
        Sql = "select * from parceladocumento where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO & " and numdocumento between 2000000 and 3000000"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                MsgBox "aqui"
            End If
           .Close
        End With
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"


End Sub

Private Sub CorrigeRefis()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPercIsencao As Integer, nPlano As Integer, nSq As Integer, RdoAux3 As rdoResultset


Sql = "SELECT DISTINCT parceladocumento.plano, numdocumento.percisencao, numdocumento.emissor, numdocumento.numdocumento FROM parceladocumento INNER JOIN "
Sql = Sql & "numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento Where (parceladocumento.plano = 0) And (NumDocumento.percisencao > 0) And (Year(NumDocumento.datadocumento) = 2017)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nNumDoc = !NumDocumento
        nPercIsencao = !percisencao
        If nPercIsencao = 100 Then
            nPlano = 16
        ElseIf nPercIsencao = 80 Then
            nPlano = 17
        ElseIf nPercIsencao = 60 Then
            nPlano = 18
        ElseIf nPercIsencao = 50 Then
            nPlano = 19
        End If
        Sql = "update parceladocumento set plano=" & nPlano & " where numdocumento=" & nNumDoc
        cn.Execute Sql, rdExecDirect
        
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"



End Sub


Private Sub CancelaUnica()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, nPercIsencao As Integer, nPlano As Integer, nSq As Integer, RdoAux3 As rdoResultset


Sql = "SELECT * FROM debitoparcela WHERE (codreduzido < 100000) AND (anoexercicio = 2018) AND (codlancamento = 1) AND (numparcela > 0) AND (statuslanc = 2) order by codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "SELECT * FROM debitoparcela WHERE codreduzido = " & !CODREDUZIDO & " AND (anoexercicio = 2018) AND (codlancamento = 1) AND (numparcela = 0)  AND (statuslanc = 3)"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            Sql = "update debitoparcela set statuslanc=5 WHERE codreduzido = " & !CODREDUZIDO & " AND (anoexercicio = 2018) AND (codlancamento = 1) AND (numparcela = 0)  AND (statuslanc = 3)"
            cn.Execute Sql, rdExecDirect
        End If
'        Sql = "update parceladocumento set plano=" & nPlano & " where numdocumento=" & nNumDoc
'        cn.Execute Sql, rdExecDirect
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"



End Sub

Private Sub Relatorio_SanMarino()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String

Sql = "truncate table relatorio_sanmarino"
cn.Execute Sql, rdExecDirect

Sql = "SELECT codreduzido, nomecidadao, LOGRADOURO, li_num, descbairro, li_quadras, li_lotes From vwFULLIMOVEL2 WHERE (li_codbairro IN (81, 1056, 1062, 1064, 1074, 1075, 1077)) ORDER BY li_codbairro, codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sExercicio = ""
        Sql = "SELECT distinct anoexercicio FROM debitoparcela WHERE codreduzido = " & !CODREDUZIDO & " AND (statuslanc = 3) and datavencimento<getdate()  order by anoexercicio"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            Do Until RdoAux2.EOF
                sExercicio = sExercicio & RdoAux2!AnoExercicio & ","
                RdoAux2.MoveNext
            Loop
            RdoAux2.Close
        End If
        If sExercicio <> "" Then
            sExercicio = Left(sExercicio, Len(sExercicio) - 1)
        End If
        Sql = "INSERT relatorio_sanmarino (codreduzido,nome,endereco,numero,bairro,quadras,lotes ,exercicio) values (" & !CODREDUZIDO & ",'" & Mask(!nomecidadao) & "','"
        Sql = Sql & !Logradouro & "'," & !Li_Num & ",'" & !DescBairro & "','" & !Li_Quadras & "','" & !Li_Lotes & "','" & sExercicio & "')"
        cn.Execute Sql, rdExecDirect
        
        DoEvents
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"

End Sub

Private Sub Corrige_Obs()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

GoTo Debito:

Sql = "truncate table obsparcela2"
'cn.Execute Sql, rdExecDirect

Sql = "SELECT * from obsparcela where anoexercicio>=2018 order by codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        Sql = "INSERT obsparcela2 (codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,seq,obs,usuario,data) values(" & !CODREDUZIDO & ","
        Sql = Sql & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & !Seq & ",'" & Mask(!obs) & "','"
        Sql = Sql & !Usuario & "','" & Format(!Data, "mm/dd/yyyy") & "')"
'        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

Debito:
Sql = "SELECT * from debitoobservacao  order by codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        Sql = "INSERT debitoobservacao2 (codreduzido,seq,usuario,dataobs,obs) values(" & !CODREDUZIDO & ","
        Sql = Sql & !Seq & ",'" & !Usuario & "','" & Format(!DATAOBS, "mm/dd/yyyy") & "','" & Mask(!obs) & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



MsgBox "fim"

Exit Sub

Erro:
'MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub ContaArea()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
Dim bR As Boolean, bC As Boolean, bId As Boolean, bIn As Boolean
Dim nCountR As Integer, nCountC As Integer, nCountId As Integer, nCountIn As Integer, nCountM As Integer, nCountTmp As Integer
On Error GoTo Erro


Sql = "SELECT codreduzido from cadimob where inativo=0 order by codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    nCountR = 0: nCountC = 0: nCountId = 0: nCountIn = 0: nCountM = 0
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !CODREDUZIDO
        nCountTmp = 0
        bR = False: bC = False: bId = False: bIn = False
        Sql = "select * from areas where codreduzido=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                If !USOCONSTR = 1 Then
                    bR = True
                ElseIf !USOCONSTR = 2 Then
                    bId = True
                ElseIf !USOCONSTR = 3 Then
                    bC = True
                ElseIf !USOCONSTR = 4 Then
                    bIn = True
                End If
               .MoveNext
            Loop
           .Close
        End With
        If bR Then
            nCountTmp = nCountTmp + 1
        End If
        If bC Then
            nCountTmp = nCountTmp + 1
        End If
        If bId Then
            nCountTmp = nCountTmp + 1
        End If
        If bIn Then
            nCountTmp = nCountTmp + 1
        End If
        If nCountTmp > 1 Then
            nCountM = nCountM + 1
        Else
            If bR Then
                nCountR = nCountR + 1
            End If
            If bC Then
                nCountC = nCountC + 1
            End If
            If bIn Then
                nCountIn = nCountIn + 1
            End If
            If bId Then
                nCountId = nCountId + 1
            End If
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



MsgBox "Residencial: " & nCountR & vbCrLf & "Comercial: " & nCountC & vbCrLf & "Industrial: " & nCountId & vbCrLf & "Institucional: " & nCountIn & vbCrLf & "Misto: " & nCountM

Exit Sub

Erro:
'MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub Codigo_Usuario()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String, x As Integer
x = 1

Sql = "SELECT * from usuario order by nomelogin"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "update usuario set id=" & x & " where nomelogin='" & !NomeLogin & "'"
        cn.Execute Sql, rdExecDirect
        
        x = x + 1
       .MoveNext
    Loop
   .Close
End With
MsgBox "fim"

End Sub

Private Sub ProcessoGTI()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT ano,numero,responsavel from processogti where responsavel is not null and userid is null order by ano,numero"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = idFromLogin(UCase(Trim(!RESPONSAVEL)))
        'nCodReduz = RetornaUsuarioID(UCase(Trim(!RESPONSAVEL)))
        If nCodReduz > 0 Then
            Sql = "update processogti set userid=" & RetornaUsuarioID(!RESPONSAVEL) & " where ano=" & !Ano & " and numero=" & !Numero
            cn.Execute Sql, rdExecDirect
        End If
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeTramite()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT ano,numero,seq,usuario,usuario2,userid,userid2 from tramitacao where usuario2 is not null and userid2 is null order by ano,numero"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = 0
        
        If Not IsNull(!Usuario2) Then
            nCodReduz = idFromLogin(UCase(Trim(!Usuario2)))
        End If
        
        Sql = "update tramitacao set userid2=" & nCodReduz & " where ano=" & !Ano & " and numero=" & !Numero & " and seq=" & !Seq
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub


Private Sub CorrigeUsuarioCC()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT nome,codigocc from usuariocc order by nome,codigocc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 100 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Nome)))
        
        Sql = "update usuariocc  set userid=" & nCodReduz & " where nome='" & !Nome & "' and codigocc=" & !codigocc
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeDebitoParcela()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT distinct usuario from debitoparcela where usuario is not null and userid is null order by usuario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Usuario)))
        
        Sql = "update debitoparcela  set userid=" & nCodReduz & " where usuario='" & !Usuario & "'"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeDebitoObservacao()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT distinct usuario from debitoobservacao where usuario is not null and userid =0 order by usuario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Usuario)))
        
        Sql = "update debitoobservacao  set userid=" & nCodReduz & " where usuario='" & !Usuario & "'"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeDebitoCancel()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT distinct usuario from debitocancel where usuario is not null and userid is null order by usuario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Usuario)))
        
        Sql = "update debitocancel  set userid=" & nCodReduz & " where usuario='" & !Usuario & "'"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeObsCidadao()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT distinct usuario from obsparcela where usuario is not null and userid =0 order by usuario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Usuario)))
        
        Sql = "update obsparcela  set userid=" & nCodReduz & " where usuario='" & !Usuario & "'"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeHistorico()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT distinct usuario from Historicocidadao where usuario is not null and userid is null order by usuario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!Usuario)))
        
        Sql = "update Historicocidadao  set userid=" & nCodReduz & " where usuario='" & !Usuario & "'"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub

Private Sub CorrigeIsencao()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long
On Error GoTo Erro

Sql = "SELECT distinct funcionario from processoreparc where funcionario is not null and userid is null order by funcionario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = idFromLogin(UCase(Trim(!funcionario)))
        
        Sql = "update processoreparc  set userid=" & nCodReduz & " where funcionario='" & !funcionario & "'"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox rdoErrors(1).Description
Resume Next

End Sub



Private Function idFromLogin(sNomeLogin As String) As Integer
Dim x As Integer, nRet As Integer

nRet = 0
For x = 1 To UBound(aIdUser)
    If aIdUser(x).Nome = sNomeLogin Then
        nRet = aIdUser(x).id
        Exit For
    End If
Next
idFromLogin = nRet

End Function


Private Sub EmiteBoletoCIP()

Dim nValorTaxa As Double, x As Integer, nSituacao As Integer, dDataProc As Date, sDescImposto As String, RdoAux2 As rdoResultset, Y As Integer, nPercTrib As Double
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, sDataVencto As String, nCodTrib As Integer, nValorTributo As Double
Dim NumBarra1 As String, StrBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
Dim sCodReduz As String, sNomeResp As String, sTipoImposto As String, sEndImovel As String, nNumImovel As Integer, sComplImovel As String, sBairroImovel As String
Dim nCodLogr As Long, sEndEntrega As String, nNumEntrega As Integer, sBairroEntrega As String, sComplEntrega As String, sCepEntrega As String, sCidadeEntrega As String
Dim sUFEntrega As String, sNumInsc As String, sValorParc As String, nNumDoc As Long, sBarra As String, sDigitavel2 As String, nValorGuia As Double, nNumGuia As Long
Dim sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long, sNumDoc As String, nSid As Long, sDigitavel As String, sNossoNumero As String, sCPF As String, sObs As String
Dim clsImovel As New clsImovel, nCodReduz As Long, sSetor As String, sRG As String, dDataPrimeiraParc As String, nValorTotalHon As Double, RdoAux3 As rdoResultset
Dim nPagina As Integer, nLivro As Integer, sDataDam As String, xImovel As clsImovel


'LIMPA TEMPORARIO
nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boletoguia where sid=" & nSid
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
Sql = Sql & "Where (cadimob.codreduzido in (select codigo from cip_semregistro where ano=2018)) "
'Sql = Sql & "Where (cadimob.li_codbairro =1069) "
Sql = Sql & " And (parceladocumento.AnoExercicio = 2018) And (parceladocumento.CodLancamento =79) And (parceladocumento.SeqLancamento = 0) "
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
    sDadosLanc = "CONTRIBUIÇÃO DE ILUMINAÇÃO PÚBLICA 2018"
    NumBarra2 = Gera2of5Cod(CStr(sValor), CDate(dDataVencto), CLng(nNumDoc), CLng(nCodReduz))
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)

    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
    sBarra = StrBarra2

    '*******************************************

        Sql = "insert boletoguia(usuario,computer,sid,seq,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,fulllanc,numdoc,numparcela,totparcela,datavencto,numdoc2,"
        Sql = Sql & "digitavel,codbarra,valorguia,obs,numbarra2a,numbarra2b,numbarra2c,numbarra2d) values('" & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & x & "," & nCodReduz & ",'" & Left(Mask(sNomeResp), 80) & "','" & sCPF & "','"
        Sql = Sql & Left(Mask(sEndImovel), 80) & "'," & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','" & Mask(sDescImposto) & "','"
        Sql = Sql & CStr(nNumGuia) & "'," & IIf(nParc = 0, 0, nParc) & "," & 3 & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & sNumDoc & "','" & sDigitavel2 & "','" & Mask(sBarra) & "',"
        Sql = Sql & Virg2Ponto(Format(nValorGuia, "#0.00")) & ",'" & "contrib" & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "')"
        cn.Execute Sql, rdExecDirect
        x = x + 1
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport2 "BOLETOGUIA_CIP", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
Liberado

End Sub

Private Sub SuspendeEmpresa()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long, nSeq As Integer
On Error GoTo Erro

Sql = "SELECT codigo from codtmp order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        nCodReduz = !Codigo
        
        Sql = "SELECT MAX(SEQEVENTO) AS MAXIMO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & nCodReduz
        Sql = Sql & " AND CODTIPOEVENTO=2"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!maximo) Then
                nSeq = 0
            Else
                nSeq = !maximo + 1
            End If
        End With

        Sql = "INSERT MOBILIARIOEVENTO (CODMOBILIARIO,CODTIPOEVENTO,SEQEVENTO,DATAEVENTO,NUMPROCEVENTO,DATAPROCEVENTO,TIPOCALCULO) VALUES("
        Sql = Sql & nCodReduz & "," & 2 & "," & nSeq & ",'" & Format("18/05/2018", "mm/dd/yyyy") & "','" & "23273/2017" & "','" & Format("14/12/2017", "mm/dd/yyyy") & "'," & 0 & ")"
        cn.Execute Sql, rdExecDirect


        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!maximo) Then
            nSeq = 0
        Else
            nSeq = RdoAux!maximo + 1
        End If
            
        sTexto1 = "A Empresa foi suspensa através do processo nº 23273-4/2017 em 18/05/2018."
            
        Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES(" & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(CStr(sTexto1)) & "',236)"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Mei()
Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, sExercicio As String
Dim nPos As Long, nTot As Long, nSeq As Integer, sHist As String, nIni As Integer, nFim As Integer, nSize As Integer

On Error GoTo Erro

Sql = "SELECT codigo from codtmp order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Do Until .EOF
        If nPos Mod 10 = 0 Then
           CallPb nPos, nTot
        End If
        
        nCodReduz = !Codigo
        sHist = ""
        Sql = "select distinct anoexercicio from debitoparcela where codreduzido=" & nCodReduz & " and "
        Sql = Sql & "anoexercicio between 2015 and 2018 and statuslanc=3 order by anoexercicio"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sHist = sHist & !AnoExercicio & ","
               .MoveNext
            Loop
           .Close
        End With
        If sHist <> "" Then
            sHist = Left(sHist, Len(sHist) - 1)
            Sql = "update codtmp set debito='" & sHist & "' where codigo=" & nCodReduz
            cn.Execute Sql, rdExecDirect
        End If
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "Fim"

Exit Sub

Erro:
MsgBox Err.Description
Resume Next

End Sub

