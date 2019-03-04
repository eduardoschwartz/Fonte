VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProdutividadeMensal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtividade - Extrato Mensal"
   ClientHeight    =   3585
   ClientLeft      =   945
   ClientTop       =   2205
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   5475
   Begin VB.TextBox txtTransportar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   9
      Top             =   3105
      Width           =   1230
   End
   Begin VB.TextBox txtReceber 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   8
      Top             =   2790
      Width           =   1230
   End
   Begin VB.TextBox txtResultado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   7
      Top             =   2475
      Width           =   1230
   End
   Begin VB.TextBox txtPontos 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   6
      Top             =   2160
      Width           =   1230
   End
   Begin VB.TextBox txtSaldo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1845
      Width           =   1230
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1530
      Width           =   1230
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmProdutividadeMensal.frx":0000
      Left            =   870
      List            =   "frmProdutividadeMensal.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1050
      Width           =   2715
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmProdutividadeMensal.frx":002E
      Left            =   3900
      List            =   "frmProdutividadeMensal.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   630
      Width           =   1245
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmProdutividadeMensal.frx":0032
      Left            =   870
      List            =   "frmProdutividadeMensal.frx":005D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.ComboBox cmbFiscal 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   4305
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   3900
      TabIndex        =   10
      ToolTipText     =   "Imprimir extrato"
      Top             =   1080
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmProdutividadeMensal.frx":00C6
      PICN            =   "frmProdutividadeMensal.frx":00E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblMatricula 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   5715
      TabIndex        =   22
      Top             =   4455
      Width           =   555
   End
   Begin VB.Label Label2 
      Caption         =   "Matricula"
      Height          =   240
      Left            =   4815
      TabIndex        =   21
      Top             =   4455
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A Transportar:"
      Height          =   270
      Index           =   9
      Left            =   195
      TabIndex        =   20
      Top             =   3165
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A Receber:"
      Height          =   270
      Index           =   8
      Left            =   195
      TabIndex        =   19
      Top             =   2850
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Resultado:"
      Height          =   270
      Index           =   7
      Left            =   195
      TabIndex        =   18
      Top             =   2520
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pontos Negativos:"
      Height          =   270
      Index           =   6
      Left            =   195
      TabIndex        =   17
      Top             =   2205
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo anterior:"
      Height          =   270
      Index           =   5
      Left            =   195
      TabIndex        =   16
      Top             =   1875
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total no mês:"
      Height          =   270
      Index           =   4
      Left            =   195
      TabIndex        =   15
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo......:"
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   14
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano......:"
      Height          =   225
      Index           =   2
      Left            =   3180
      TabIndex        =   13
      Top             =   690
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mês......:"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   12
      Top             =   660
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fiscal....:"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   11
      Top             =   240
      Width           =   705
   End
End
Attribute VB_Name = "frmProdutividadeMensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ExtratoType
    nDia As Integer
    sDia As String
    sEvento As String
    nPontos As Double
End Type

Private Type aEventoType
    sNome As String
    nPontos As Double
End Type

Private Type ExtratoItemType
    nDia As Integer
    sItem As String
    sDesc As String
    nQtde As Double
    nValor As Double
    nPontos As Double
    sProcesso As String
    sObs As String
End Type

Private Type SaldoType
    sMesAno As String
    nSaldo As Double
End Type

Dim sNomeArq As String, FF1 As Integer, aExtrato() As ExtratoType, aEvento() As aEventoType, aPontos() As Double
Dim nLastDay As Integer, nSomaMes As Double, nSaldoAnterior As Double, aExtratoItem() As ExtratoItemType

Private Sub CarregaEvento()
Dim Sql As String, RdoAux As rdoResultset

ReDim aEvento(0)
Sql = "select codigo,nome,pontodia from produtividadeevento order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aEvento(UBound(aEvento) + 1)
        aEvento(UBound(aEvento)).sNome = !Nome
        aEvento(UBound(aEvento)).nPontos = !pontodia
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbFiscal_Click()
Dim Sql As String, RdoAux As rdoResultset

Sql = "select nome,matricula from produtividadefiscal where nome='" & RetornaUsuarioLoginName(cmbFiscal.Text) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux.RowCount > 0 Then
        lblMatricula.Caption = Val(SubNull(!MATRICULA))
    End If
   .Close
End With

End Sub

Private Sub cmdPrint_Click()
Dim z As Long, sFileName As String

If cmbFiscal.ListIndex = -1 Then
    MsgBox "Selecione um fiscal.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbTipo.ListIndex = -1 Then
    MsgBox "Selecione um tipo de relatório.", vbCritical, "Atenção"
    Exit Sub
End If

Select Case cmbTipo.ListIndex
    Case 0
        Ocupado
        FillPontosMes
        FillArray
        sFileName = "ExtratoProd1.txt"
        FF1 = FreeFile()
        Open sPathBin & "\" & sFileName For Output As FF1
        PrintExtrato1
        Close #FF1
        Liberado
        z = Shell("NOTEPAD" & " " & sPathBin & "\" & sFileName, vbNormalFocus)
    Case 1
        PrintExtrato2
    Case 2
        PrintSaldoMes
End Select

End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset, x As Integer
On Error GoTo Erro:

Centraliza Me

Sql = "select codigo,nome,nomecompleto from produtividadefiscal inner join "
Sql = Sql & "usuario on produtividadefiscal.nome = usuario.nomelogin order by nomecompleto "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFiscal.AddItem !NomeCompleto
        cmbFiscal.ItemData(cmbFiscal.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

If NomeDeLogin <> "SCHWARTZ" And NomeDeLogin <> "NOELI" Then
    If Not ProdIsBossLogin() Then
        cmbFiscal.Text = RetornaUsuarioFullName()
        cmbFiscal.Enabled = False
    End If
End If

For x = 2011 To 2019
    cmbAno.AddItem (CStr(x))
Next

cmbMes.ListIndex = Month(Now) - 1
cmbAno.Text = Year(Now)
CarregaEvento

Exit Sub
Erro:
MsgBox "Erro Fatal!"
cmbFiscal.Enabled = False
cmdPrint.Enabled = False

End Sub

Private Sub FillArray()
Dim sData As String, nMes As Integer, nDay As Long, nWeekDay As Long
Dim nCodEvento As Integer, nCodFiscal As Integer, sEvento As String, nPontos As Double
Dim bIsBoss As Boolean, RdoAux2 As rdoResultset

ReDim aExtrato(0)
nSaldoAnterior = 423 'CORRIGIR
nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
bIsBoss = ProdIsBoss(cmbFiscal.ItemData(cmbFiscal.ListIndex))
nMes = cmbMes.ItemData(cmbMes.ListIndex)

If cmbAno.Text = 2015 And nMes = 2 Then
    nLastDay = 28
End If
For nDay = 1 To nLastDay
    ReDim Preserve aExtrato(UBound(aExtrato) + 1)
    aExtrato(nDay).nDia = nDay
    sData = Format(nDay, "00") & "/" & Format(nMes, "00") & "/" & cmbAno.Text
    If sData = "29/02/2019" Then sData = "28/02/2019"
    nWeekDay = Weekday(CDate(sData))
    aExtrato(nDay).sDia = WeekdayName(nWeekDay, True)

'    Sql = "select * from feriadodef where  dia=" & nDay & " and mes=" & nMes & " and ano=" & Val(cmbAno.Text)
'    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    If RdoAux2.RowCount > 0 Then
'        sEvento = "SÁB/DOM/FER."
'        nPontos = 0
'        aExtrato(nDay).sEvento = FillSpace(sEvento, 16)
'        aExtrato(nDay).nPontos = nPontos
'        RdoAux2.Close
'        GoTo Proximo
 '   End If

    nCodEvento = ProdEventoDia(nCodFiscal, CDate(sData))
    If nCodEvento > 0 Then
        If (nWeekDay = 1 Or nWeekDay = 7) And nCodEvento <> 2 And nCodEvento <> 3 Then
            sEvento = "SÁB/DOM/FER."
            nPontos = 0
'            nPontos = aEvento(nCodEvento).nPontos
        Else
            sEvento = aEvento(nCodEvento).sNome
            If nCodEvento = 3 Then 'LICENÇA PREMIO NÃO TEM PONTOS
                nPontos = 20
            Else
                If bIsBoss Then 'CHEFE SEMPRE TEM PONTOS CHEIOS
                    nPontos = 30
                Else
                    nPontos = aEvento(nCodEvento).nPontos
                End If
            End If
            
            If Len(sEvento) > 12 Then
                sEvento = Left(sEvento, 12) & "."
            Else
                sEvento = sEvento
            End If
        End If
        
        aExtrato(nDay).sEvento = FillSpace(sEvento, 16)
        aExtrato(nDay).nPontos = nPontos
    Else
        If nWeekDay = 1 Or nWeekDay = 7 Then
            sEvento = "SÁB/DOM/FER."
            nPontos = aPontos(nDay)
            'nPontos = 0
        Else
            If bIsBoss Then
                sEvento = "CHEFIA"
                nPontos = 30
            Else
                sEvento = "NORMAL"
                nPontos = aPontos(nDay)
            End If
        End If
        aExtrato(nDay).sEvento = FillSpace(sEvento, 16)
        aExtrato(nDay).nPontos = nPontos
    End If
proximo:
Next

End Sub

Private Sub PrintExtrato1()
Dim sNome As String, nDay As Integer, ax As String, nPos As Integer, nMes As Integer, nLastDay As Integer
Dim bAchou As Boolean, bPrint As Boolean, nSomaPontos As Double

If Len(cmbFiscal.Text) > 21 Then
    sNome = Left(cmbFiscal, 21) & "."
Else
    sNome = cmbFiscal.Text
End If

nMes = cmbMes.ItemData(cmbMes.ListIndex)
nLastDay = Val(Left(Format$(DateSerial(2012, Val(nMes) + 1, 0), "dd/mm/yyyy"), 2))
nSomaPontos = 0

'******* HEADER ********
Print #FF1, "EXTRATO MENSAL - PRODUTIVIDADE"
Print #FF1, "MÊS REFERÊNCIA: " & cmbMes.Text & "/" & cmbAno.Text
Print #FF1, "FISCAL: " & sNome
'Print #FF1, "SALDO ANTERIOR: " & Format(nSaldoAnterior, "000") & " PONTOS"
Print #FF1, "=============================="
Print #FF1, "DIA  SEM  EVENTO        PONTOS"
Print #FF1, "=============================="

'******* BODY ********

For nDay = 1 To UBound(aExtrato)
    With aExtrato(nDay)
        ax = " " & Format(.nDia, "00") & "  " & .sDia & "  " & .sEvento & " " & FormatNumber(.nPontos, 2)
        nSomaPontos = nSomaPontos + .nPontos
        Print #FF1, ax
    End With
Next

'******* FOOTER1 ********
Print #FF1, "=============================="
Print #FF1, "SOMA DE PONTOS NO MÊS: " & FormatNumber(nSomaPontos, 2)
'Print #FF1, "SALDO ACUMULADO: "; Format(nSaldoAnterior + nSomaMes, "000")
Print #FF1, ""
Print #FF1, ""
Print #FF1, "DESCRIMINAÇÃO DAS TAREFAS DIARIAS"
Print #FF1, "DESCRIÇÃO DA TAREFA         PROCESSO       QTDE VALOR    PONTOS OBS"
Print #FF1, "=================================================================================="

For nDay = 1 To nLastDay
    
    bPrint = False
    For nPos = 1 To UBound(aExtratoItem)
        bAchou = False
        With aExtratoItem(nPos)
            If .nDia = nDay Then
                bAchou = True
            End If
            If bAchou And Not bPrint Then
                Print #FF1, ""
                Print #FF1, "DIA: " & Format(nDay, "00") & " (" & aExtrato(nDay).sDia & ")"
                Print #FF1, "============="
                bPrint = True
            End If
            If bAchou Then
                ax = FillSpace(.sItem, 27) & " " & .sProcesso & "  " & FormatNumber(.nQtde, 2) & "  " & FormatNumber(.nValor, 2) & "     " & FormatNumber(.nPontos, 2) & " " & .sObs
                Print #FF1, ax
            End If
        End With
    Next
    If bPrint Then
    Print #FF1, "                                         TOTAL NO DIA ==> " & FormatNumber(aExtrato(nDay).nPontos, 2)
    End If
Next

'******* FOOTER2 ********
Print #FF1, ""
Print #FF1, "IMPRESSO EM: " & Format(Now, "dd/mm/yyyy")
Print #FF1, "Módulo EXPMEN01(G.T.I.)"

End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub FillPontosMes()
Dim nMes As Integer, Sql As String, RdoAux As rdoResultset, nCodFiscal As Integer, nSomaTarefa As Double, nDia As Integer
Dim bAchou As Boolean, nPos As Integer, sNome As String, RdoAux2 As rdoResultset

nSomaMes = 0
nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
nMes = cmbMes.ItemData(cmbMes.ListIndex)
nLastDay = Val(Left(Format$(DateSerial(2012, Val(nMes) + 1, 0), "dd/mm/yyyy"), 2))
ReDim aPontos(nLastDay)
ReDim aExtratoItem(0)

Sql = "SELECT produtividadetarefa.data, produtividadetarefa.item, produtividadetarefa.qtde, produtividadetarefa.valor, produtividadetarefa.processo,produtividadetarefa.obs, "
Sql = Sql & "produtividadedesc.descricao FROM produtividadetarefa INNER JOIN produtividadedesc ON produtividadetarefa.item = produtividadedesc.item "
Sql = Sql & "where year(data)=" & Val(cmbAno.Text) & " and month(data)=" & nMes & " and fiscal=" & nCodFiscal
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nDia = Day(!Data)
        nSomaTarefa = !Valor * !QTDE
        nSomaMes = nSomaMes + nSomaTarefa
        aPontos(nDia) = FormatNumber(aPontos(nDia) + nSomaTarefa, 2)
        
        ' ** EXTRATO ITEM
        ReDim Preserve aExtratoItem(UBound(aExtratoItem) + 1)
        nPos = UBound(aExtratoItem)
        aExtratoItem(nPos).nDia = nDia
        aExtratoItem(nPos).nQtde = FormatNumber(!QTDE, 2)
        aExtratoItem(nPos).nValor = FormatNumber(!Valor, 2)
        aExtratoItem(nPos).nPontos = FormatNumber(!QTDE * !Valor, 2)
        If Len(!descricao) > 21 Then
            sNome = Left(!descricao, 21) & "."
        Else
            sNome = !descricao
        End If
        
        aExtratoItem(nPos).sItem = !Item & " " & sNome
        aExtratoItem(nPos).sProcesso = SubNull(!Processo)
        If aExtratoItem(nPos).sProcesso = "" Then aExtratoItem(nPos).sProcesso = "000000-0/0000"
        aExtratoItem(nPos).sObs = SubNull(!obs)
        ' **
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub PrintExtrato2()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, nCodFiscal As Integer, nMes As Integer, nLastDay As Integer
Dim nSoma As Double, nTotal As Double, nSaldo As Double, nPontosNeg As Double, nResultado As Double, bIsBoss As Boolean
Dim nReceber As Double, nTransportar As Double, nContaCurso As Integer, nContaFerias As Integer, nContaChefia As Integer
Dim nContaDiaUtil As Integer, sDataTmp As String, nWeekDay As Integer, nCodEvento As Integer, nContaDiaPerdido As Integer
Dim sDataFerias As String, sDataExt As String, nAno As Integer, RdoAux2 As rdoResultset, nContaLic As Integer, sHistferias As String

nTotal = 0: nSaldo = 0: nPontosNeg = 0:
nResultado = 0: nReceber = 0: nTransportar = 0
nContaCurso = 0: nContaFerias = 0: nContaChefia = 0
nContaDiaUtil = 0: nContaDiaPerdido = 0: nContaLic = 0: nContaLicPremio = 0

ReDim aExtratoItem(0)
nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
bIsBoss = ProdIsBoss(nCodFiscal)

nMes = cmbMes.ItemData(cmbMes.ListIndex)
nLastDay = Val(Left(Format$(DateSerial(2012, Val(nMes) + 1, 0), "dd/mm/yyyy"), 2))

Sql = "select item,descricao from produtividadedesc order by item"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aExtratoItem(UBound(aExtratoItem) + 1)
        aExtratoItem(UBound(aExtratoItem)).sItem = !Item
        aExtratoItem(UBound(aExtratoItem)).sDesc = !descricao
       .MoveNext
    Loop
   .Close
End With

ReDim Preserve aExtratoItem(UBound(aExtratoItem) + 1)
aExtratoItem(UBound(aExtratoItem)).sItem = "Afastamento"
aExtratoItem(UBound(aExtratoItem)).sDesc = ""


Sql = "select data,item,qtde,valor from produtividadetarefa where fiscal=" & nCodFiscal & " and year(data)=" & Val(cmbAno.Text) & " and month(data)=" & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For x = 1 To UBound(aExtratoItem)
            If aExtratoItem(x).sItem = !Item Then
                aExtratoItem(x).nQtde = aExtratoItem(x).nQtde + !QTDE
                aExtratoItem(x).nValor = !Valor
                Exit For
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

'CONTA OS DIAS UTEIS, DIAS DE CURSO,CHEFIA E FERIAS
ReDim aFiscalEvento(0)

For x = 1 To nLastDay
    sDataTmp = Format(x, "00") & "/" & Format(nMes, "00") & "/" & cmbAno.Text
    If sDataTmp = "29/02/2019" Then sDataTmp = "28/02/2019"
    nWeekDay = Weekday(CDate(sDataTmp))
    nCodEvento = ProdEventoDia(nCodFiscal, CDate(sDataTmp))
    If nCodEvento = 0 And bIsBoss Then
        nCodEvento = 1
    End If
    
    Sql = "select * from feriadodef where  dia=" & x & " and mes=" & nMes & " and ano=" & Val(cmbAno.Text)
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
        RdoAux2.Close
        'nContaDiaPerdido = nContaDiaPerdido + 1
        GoTo proximo
    End If
    
    If nCodEvento > 0 Then
        If (nWeekDay = 1 Or nWeekDay = 7) And nCodEvento <> 2 And nCodEvento <> 3 Then
            'sabado,domingo
        Else
            nContaDiaUtil = nContaDiaUtil + 1
            If nCodEvento = 1 Then 'CHEFIA
                nContaChefia = nContaChefia + 1
            ElseIf nCodEvento = 2 Or nCodEvento = 7 Then 'FERIAS/doação de sangue
                nContaFerias = nContaFerias + 1
                nContaDiaPerdido = nContaDiaPerdido + 1
                sDataFerias = sDataTmp
            ElseIf nCodEvento = 3 Or nCodEvento = 4 Then 'Licença médica
                nContaLic = nContaLic + 1
                nContaDiaPerdido = nContaDiaPerdido + 1
                sDataFerias = sDataTmp
            ElseIf nCodEvento = 6 Then 'CURSO
                nContaCurso = nContaCurso + 1
                nContaDiaPerdido = nContaDiaPerdido + 1
            Else 'OUTRAS LICENÇAS
                nContaDiaPerdido = nContaDiaPerdido + 1
            End If
        End If
    End If
proximo:
Next


If nContaFerias > 0 Then
'    nContaFerias = nContaFerias - 1
    Sql = "SELECT codfiscal, seq, codevento, dataini, datafim From produtividadefiscalevento "
    Sql = Sql & "WHERE codfiscal = " & nCodFiscal & " AND codevento = 2 AND ('" & Format(sDataFerias, "mm/dd/yyyy") & "' BETWEEN dataini AND datafim)"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            'sDataFerias = "Período de: " & Format(!dataini, "dd/mm/yyyy") & " até " & Format(!datafim, "dd/mm/yyyy")
            sHistferias = "Período de: " & Format(!dataini, "dd/mm/yyyy") & " até " & Format(!datafim, "dd/mm/yyyy")
        End If
       .Close
    End With
End If

If nContaLic > 0 Then
    nContaLic = 0
    Sql = "SELECT codfiscal, seq, codevento, dataini, datafim From produtividadefiscalevento "
    Sql = Sql & "WHERE codfiscal = " & nCodFiscal & " AND (codevento = 4 or codevento=3) AND ('" & Format(sDataFerias, "mm/dd/yyyy") & "' BETWEEN dataini AND datafim)"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            nContaLic = nContaLic + DateDiff("d", !dataini, !datafim) + 1
            
           .MoveNext
        Loop
       .Close
    End With
End If

If bIsBoss Then
    nContaChefia = nContaDiaUtil - nContaDiaPerdido
End If

For x = 1 To UBound(aExtratoItem)
    If aExtratoItem(x).sItem = "14" And nContaCurso > 0 Then 'CURSO
        aExtratoItem(x).nQtde = nContaCurso
        aExtratoItem(x).nValor = 15
    ElseIf aExtratoItem(x).sItem = "22" And nContaChefia > 0 Then 'CHEFE
        aExtratoItem(x).nQtde = nContaChefia
        aExtratoItem(x).nValor = 30
    ElseIf aExtratoItem(x).sItem = "Afastamento" And (nContaFerias > 0 Or nContaLic > 0) Then 'FERIAS
        aExtratoItem(x).nQtde = nContaFerias
        If nContaLic > 0 Then
            aExtratoItem(x).nQtde = nContaLic
            aExtratoItem(x).nValor = 20
            aExtratoItem(x).sDesc = "Licença"
        Else
            aExtratoItem(x).nValor = 20
            aExtratoItem(x).sDesc = sHistferias
        End If
        
    End If
Next

Sql = "delete from produtividaderel1 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

For x = 1 To UBound(aExtratoItem)
    With aExtratoItem(x)
        Sql = "insert produtividaderel1(usuario,item,descricao,valor,qtde,total) values('" & NomeDeLogin & "','"
        Sql = Sql & .sItem & "','" & .sDesc & "'," & Virg2Ponto(CStr(.nValor)) & "," & Virg2Ponto(CStr(.nQtde)) & "," & Virg2Ponto(CStr(.nQtde * .nValor)) & ")"
        cn.Execute Sql, rdExecDirect
    
        nSoma = .nQtde * .nValor
        nTotal = nTotal + nSoma
    End With
Next

If nTotal > 900 Then nTotal = 900

'saldo anterior
If nMes = 1 Then
    nMes = 12
    nAno = Val(cmbAno.Text) - 1
Else
    nMes = nMes - 1
    nAno = Val(cmbAno.Text)
End If

Sql = "SELECT SUM(saldo) AS soma From produtividadesaldo Where codfiscal = " & nCodFiscal & " And  anoref=" & nAno & " and MesRef = " & cmbMes.ItemData(cmbMes.ListIndex) & " and mes=" & nMes
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!soma) Then
    nSaldo = 0
Else
    nSaldo = RdoAux!soma
End If
RdoAux.Close

'resultado mes
nResultado = nTotal + nSaldo

If nResultado >= 600 Then
    nReceber = 400
    nTransportar = nResultado - 600
Else
    nReceber = nResultado - 200
    nTransportar = 0
End If




'lblTotal.Caption = FormatNumber(nTotal, 2)
'lblSaldo.Caption = FormatNumber(nSaldo, 2)
'lblPontos.Caption = FormatNumber(nPontosNeg)
'lblResultado.Caption = FormatNumber(nResultado, 2)
'lblReceber.Caption = FormatNumber(nReceber, 2)
'lblTransportar.Caption = FormatNumber(nTransportar, 2)

frmReport.ShowReport2 "PRODMOB1", frmMdi.HWND, Me.HWND

Sql = "delete from produtividaderel1 where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub

Private Sub PrintSaldoMes()
Dim sFileName As String, z As Long, Sql As String, RdoAux As rdoResultset, nCodFiscal As Integer, nMes As Integer, nAno As Integer
Dim aSaldo(12) As SaldoType, sMesAno As String, x As Integer, dDataTmp As Date, ax As String
Exit Sub

nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
nMes = cmbMes.ItemData(cmbMes.ListIndex)
nAno = Val(cmbAno.Text)

Sql = "SELECT * FROM PRODUTIVIDADESALDO WHERE CODFISCAL=" & nCodFiscal & " AND MESREF=" & nMes & " AND ANOREF=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Não foi calculado o saldo para este período.", vbExclamation, "Atenção"
        .Close
        Exit Sub
    Else
        For x = 12 To 1 Step -1
            sMesAno = MonthName(nMes) & "/" & CStr(nAno)
            aSaldo(x).sMesAno = sMesAno
            nMes = nMes - 1
            If nMes = 0 Then
                nMes = 12
                nAno = nAno - 1
            End If
        Next
    End If
   .Close
End With


sFileName = "SaldoMes.txt"
FF1 = FreeFile()
Open sPathBin & "\" & sFileName For Output As FF1

For x = 1 To 12
    ax = aSaldo(x).sMesAno
    Print #1, ax
Next

Close #FF1
Liberado
z = Shell("NOTEPAD" & " " & sPathBin & "\" & sFileName, vbNormalFocus)

End Sub


Private Sub txtPontos_KeyPress(KeyAscii As Integer)
Tweak txtPontos, KeyAscii, DecimalAllowNegative, 2
End Sub

Private Sub txtReceber_KeyPress(KeyAscii As Integer)
Tweak txtReceber, KeyAscii, DecimalAllowNegative, 2
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
Tweak txtSaldo, KeyAscii, DecimalAllowNegative, 2
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
Tweak txtTotal, KeyAscii, DecimalAllowNegative, 2
End Sub

Private Sub txtTransportar_KeyPress(KeyAscii As Integer)
Tweak txtTransportar, KeyAscii, DecimalAllowNegative, 2
End Sub
