VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmAnaliseReceita 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analise da Receita"
   ClientHeight    =   2160
   ClientLeft      =   6000
   ClientTop       =   3180
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2160
   ScaleWidth      =   6030
   Begin VB.CheckBox chkSimples 
      Caption         =   "Somente SNac."
      Height          =   195
      Left            =   4365
      TabIndex        =   4
      Top             =   720
      Width           =   1545
   End
   Begin MSFlexGridLib.MSFlexGrid grdTmp 
      Height          =   4725
      Left            =   60
      TabIndex        =   15
      Top             =   2190
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   8334
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   "^DATA      |BANCO |CODREDUZIDO  |^ANO      |^LANC |^SEQ |^PARC  |^COMPL "
   End
   Begin VB.ComboBox cmbBanco 
      Height          =   315
      ItemData        =   "frmAnaliseReceita.frx":0000
      Left            =   2010
      List            =   "frmAnaliseReceita.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   630
      Width           =   2265
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Todos"
      Height          =   210
      Left            =   1005
      TabIndex        =   2
      Top             =   675
      Value           =   1  'Checked
      Width           =   840
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1185
      TabIndex        =   0
      Top             =   120
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmAnaliseReceita.frx":0004
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   3945
      TabIndex        =   1
      Top             =   135
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmAnaliseReceita.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4425
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   1215
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmAnaliseReceita.frx":003C
      PICN            =   "frmAnaliseReceita.frx":0058
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCalculo 
      Height          =   315
      Left            =   3195
      TabIndex        =   5
      ToolTipText     =   "Cancelar Edição"
      Top             =   1215
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gerar"
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
      MICON           =   "frmAnaliseReceita.frx":00C6
      PICN            =   "frmAnaliseReceita.frx":00E2
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
      Height          =   225
      Left            =   405
      TabIndex        =   10
      Top             =   1275
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7380
      TabIndex        =   14
      ToolTipText     =   "Cancelar Edição"
      Top             =   1170
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gerar"
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
      MICON           =   "frmAnaliseReceita.frx":0181
      PICN            =   "frmAnaliseReceita.frx":019D
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
      Caption         =   "Banco..:"
      Height          =   210
      Left            =   180
      TabIndex        =   13
      Top             =   675
      Width           =   765
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   75
      TabIndex        =   12
      Top             =   1275
      Width           =   270
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   1290
      Width           =   480
   End
   Begin VB.Label lblMsg 
      Caption         =   "Selecione as Datas de Inicio e Término"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   60
      TabIndex        =   9
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2925
      TabIndex        =   8
      Top             =   180
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   165
      Width           =   1035
   End
End
Attribute VB_Name = "frmAnaliseReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

Private Type Analise
    DataReceita As Date
    CodBanco As Integer
    CodTributo As Integer
    ValorTributo As Double
    ValorTotal As Double
    NumFicha As Long
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
End Type

Private Type Diferenca
    CodReduz As Long
    AnoExercicio As Integer
    CodLancamento As Integer
    SeqLancamento As Integer
    NumParcela As Integer
    CodCompl As Integer
End Type

Private Sub chk1_Click()

If chk1.value = 1 Then
    cmbBanco.Enabled = False
Else
    cmbBanco.Enabled = True
End If

End Sub

Private Sub cmdCalculo_Click()
Dim RdoAux As rdoResultset, Sql As String


If Not IsDate(mskDataIni.Text) Then
    MsgBox "Data de Inicio inválido", vbExclamation, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataFim.Text) Then
    MsgBox "Data de Fim inválido", vbExclamation, "atenção"
    Exit Sub
End If

If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    MsgBox "Data de Inicio tem que ser maior que data de termino", vbExclamation, "atenção"
    Exit Sub
End If

GeraAnalise
Ocupado

'Sql = "SELECT SUM(VALORTOTAL) AS TOTAL FROM ANALISERECEITA WHERE COMPUTER='" & NomeDeLogin & "'"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'MsgBox "A Soma total das analises = " & FormatNumber(RdoAux!Total, 2), vbInformation, "Soma das análises"

frmReport.ShowReport "Analise", frmMdi.HWND, Me.HWND
'frmReport.ShowReport "AnaliseSD", frmMdi.hwnd, Me.hwnd
Sql = "DELETE FROM ANALISERECEITA WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Liberado

End Sub

Private Sub GeraAnalise()
Dim cDataIni As Date, cDataFim As Date, cData As Date, xId As Long, nNumRec As Long, nSomaFicha As Double, i As Integer
Dim bDA As Boolean, bAj As Boolean, Matrix() As Analise, nCodFicha As Long, nValorPago As Double
Dim nCodFichaJM As Long, nCodFichaC As Long, x As Integer, bAchou As Boolean, ax As String
Dim nSomaTributo As Double, nDif As Double, nCodBanco As Integer, nPerc As Double, nValorJMC As Double
Dim nSomaPago As Double, nSomaDif As Double, nSomaMatriz As Double, nMaiorValor As Double, nIndMaior As Integer
Dim aBanco() As Integer, nContador As Integer, nContaBanco As Integer, t As Integer, sNumProc As String, nValorTaxa As Double
Dim nSomaDC As Double, nSomaRC As Double, nTotalMatriz As Double, sNatureza As String, sVinculo As String, nValorTotal As Double


nSomaPago = 0: nSomaTotal = 0

cDataIni = CDate(mskDataIni.Text): cDataFim = CDate(mskDataFim.Text)
cn.QueryTimeout = 0
Pb.value = 0
Sql = "DELETE FROM ANALISERECEITA WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

If chk1.value = 1 Then
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") For Output Shared As #1
'   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & "SD" For Output Shared As #2
Else
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") For Output Shared As #1
'   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") & "SD" For Output Shared As #2
End If

For cData = cDataIni To cDataFim
    lblMsg.Caption = "Gerando analise do dia.: " & cData
    lblMsg.Refresh
    If cGetInputState() <> 0 Then DoEvents
    
    If chk1.value = 1 Then
        Sql = "SELECT DISTINCT CODBANCO FROM DEBITOPAGO "
        Sql = Sql & "WHERE DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' AND RESTITUIDO IS NULL AND CODBANCO<>91 AND CODBANCO<>92 AND CODBANCO<>93 AND CODBANCO<>94 AND CODBANCO<>95 AND CODBANCO<>96 AND CODBANCO<>97 AND CODBANCO<>98 AND CODBANCO<>99"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            ReDim aBanco(0)
            Do Until .EOF
                ReDim Preserve aBanco(UBound(aBanco) + 1)
                aBanco(UBound(aBanco)) = !CodBanco
               .MoveNext
            Loop
           .Close
        End With
        nContaBanco = UBound(aBanco)

        Sql = "SELECT DISTINCT CODBANCO FROM RECEITACLASSIFICAR WHERE DATARECEITA ='" & Format(cData, "mm/dd/yyyy") & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                bAchou = False
                For t = 1 To UBound(aBanco)
                    If aBanco(t) = !CodBanco Then
                        bAchou = True
                        Exit For
                    End If
                Next
                If Not bAchou Then
                    ReDim Preserve aBanco(UBound(aBanco) + 1)
                    aBanco(UBound(aBanco)) = !CodBanco
                End If
               .MoveNext
            Loop
           .Close
        End With
        nContaBanco = UBound(aBanco)
    Else
        ReDim aBanco(1)
        aBanco(1) = cmbBanco.ItemData(cmbBanco.ListIndex)
        nContaBanco = 1
    End If
    
    For nContador = 1 To nContaBanco
        ReDim Matrix(0)
        If chk1.value = 1 Then
        Else
            ReDim aBanco(1)
            aBanco(1) = cmbBanco.ItemData(cmbBanco.ListIndex)
        End If
        
        Sql = "SELECT SUM(DEBITOPAGO.VALORPAGOREAL) AS VALORPAGOTOTAL FROM DEBITOPARCELA RIGHT OUTER JOIN DEBITOPAGO ON DEBITOPARCELA.CODREDUZIDO = DEBITOPAGO.CODREDUZIDO AND "
        Sql = Sql & "DEBITOPARCELA.AnoExercicio = DEBITOPAGO.AnoExercicio AND DEBITOPARCELA.CodLancamento = DEBITOPAGO.CodLancamento AND DEBITOPARCELA.SeqLancamento = DEBITOPAGO.SeqLancamento AND "
        Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOPAGO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOPAGO.CODCOMPLEMENTO "
        'Sql = Sql & "WHERE VALORPAGOREAL>0 AND RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' AND (CODBANCO=" & aBanco(nContador)
        Sql = Sql & "WHERE RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' and "
        If chkSimples.value = vbChecked Then
            Sql = Sql & " (CODBANCO=91 OR CODBANCO=92 OR CODBANCO=93 OR CODBANCO=94 OR CODBANCO=95 OR CODBANCO=96 OR CODBANCO=97 OR CODBANCO=98) "
        Else
            Sql = Sql & "CodBanco = " & aBanco(nContador)
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
             If Not IsNull(!VALORPAGOTOTAL) Then
                 nSomaPago = !VALORPAGOTOTAL
             Else
                nSomaPago = 0
             End If
            .Close
        End With
                           
        Sql = "SELECT DISTINCT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,"
        Sql = Sql & "DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.DATAINSCRICAO,DEBITOPARCELA.DATAAJUIZA,DEBITOPAGO.DATAPAGAMENTO,DEBITOPAGO.DATARECEBIMENTO,DEBITOPAGO.VALORPAGO,DEBITOPAGO.VALORPAGOREAL, DEBITOPAGO.CODBANCO,"
        Sql = Sql & "DEBITOPAGO.CodAgencia,DEBITOPAGO.RESTITUIDO FROM DEBITOPARCELA INNER JOIN DEBITOPAGO ON DEBITOPARCELA.CODREDUZIDO = DEBITOPAGO.CODREDUZIDO AND "
        Sql = Sql & "DEBITOPARCELA.AnoExercicio = DEBITOPAGO.AnoExercicio AND DEBITOPARCELA.CodLancamento = DEBITOPAGO.CodLancamento AND DEBITOPARCELA.SeqLancamento = DEBITOPAGO.SeqLancamento AND "
        Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOPAGO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOPAGO.CODCOMPLEMENTO "
        Sql = Sql & "WHERE VALORPAGOREAL>0 AND RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' AND "
        If chkSimples.value = vbChecked Then
            Sql = Sql & " (CODBANCO=91 OR CODBANCO=92 OR CODBANCO=93 OR CODBANCO=94 OR CODBANCO=95 OR CODBANCO=96 OR CODBANCO=97 OR CODBANCO=98) "
        Else
            Sql = Sql & "(CodBanco = " & aBanco(nContador) & ")"
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nNumRec = .RowCount
            nDif = 0
            Do Until .EOF
                If xId Mod 10 = 0 Then
                    CallPb xId, nNumRec
                End If
                nValorPago = !valorpagoreal
                nCodBanco = aBanco(nContador)
                
                
                'If !CODREDUZIDO = 502113 Then MsgBox "TESTE"
                
                If !CodLancamento = 20 Then
                    '***  parcelamentos *****
                    Sql = "SELECT NUMPROCESSO FROM DESTINOREPARC WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND NUMSEQUENCIA=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux2
                        If .RowCount = 0 Then
                            bDA = True: bAj = True
                            GoTo continua
                        End If
                        sNumProc = !numprocesso
                       .Close
                    End With
                    
                    Sql = "SELECT * FROM ORIGEMREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux2
                        Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                        Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !numsequencia & " AND NUMPARCELA=" & !NumParcela & " AND "
                        Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                        With RdoAux3
                            If Not IsNull(!datainscricao) Then
                               bDA = True
                            Else
                               bDA = False
                            End If
                            If Not IsNull(!dataajuiza) Then
                               bAj = True
                            Else
                               bAj = False
                            End If
                            .Close
                        End With
                        .Close
                    End With
                    '************************
                Else
                    If Not IsNull(!datainscricao) Then
                       bDA = True
                    Else
                       bDA = False
                    End If
                    If Not IsNull(!dataajuiza) Then
                       bAj = True
                    Else
                       bAj = False
                    End If
                End If
                
continua:
                
                Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If Not IsNull(!soma) Then
                       nSomaTributo = !soma
                    Else
                       nSomaTributo = 0
                    End If
                   .Close
                End With
                
                nValorTaxa = 0
                Sql = "SELECT parceladocumento.codreduzido, NumDocumento.valortaxadoc FROM parceladocumento INNER JOIN numdocumento ON "
                Sql = Sql & "parceladocumento.numdocumento = numdocumento.numdocumento where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and "
                Sql = Sql & "codlancamento=" & !CodLancamento & " and seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    nValorTaxa = RdoAux2!ValorTaxaDoc
                End If
                RdoAux2.Close
                
                nSomaTributo = nSomaTributo + nValorTaxa
                If nValorPago >= nSomaTributo Then
                    nDif = nValorPago - nSomaTributo
                Else
                    nDif = nValorPago - nSomaTributo
                End If
                Sql = "SELECT DEBITOTRIBUTO.*, DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,VALORJUROS,VALORMULTA,VALORCORRECAO,TRIBUTO.FICHA,"
                Sql = Sql & "TRIBUTO.FICHAJRMULTA, TRIBUTO.FICHADIVIDA,TRIBUTO.FICHADAJRMUL,TRIBUTO.FICHADAENCA,TRIBUTO.FICHAAJUIZA,TRIBUTO.FICHAAJJRMUL,FICHAAJENCA "
                Sql = Sql & "FROM DEBITOTRIBUTO LEFT OUTER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
                Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    nCodFicha = 0: nCodFichaJM = 0: nCodFichaC = 0
                    Do Until .EOF
                        'If !CodTributo = 1 Then MsgBox "TESTE"
                    
                        If nSomaTributo = 0 Then GoTo proximo
                        nPerc = !ValorTributo / nSomaTributo 'PRINCIPAL / SOMA DOS TRIBUTOS
                        nValorJMC = nDif * nPerc 'DIFERENCA * PERCENTUAL (JUROS,MULTA E CORRECAO)
                        If Not bAj Then
                            If Not bDA Then
                                nCodFicha = Val(SubNull(!Ficha))
                                nCodFichaJM = Val(SubNull(!FichaJrMulta))
                            Else
                                nCodFicha = Val(SubNull(!FichaDivida))
                                nCodFichaJM = Val(SubNull(!FichaDaJrMul))
                                nCodFichaC = Val(SubNull(!FichaDaEnca))
                            End If
                        Else
                            nCodFicha = Val(SubNull(!FichaAjuiza))
                            nCodFichaJM = Val(SubNull(!FichaAjJrMul))
                            nCodFichaC = Val(SubNull(!FichaAjEnca))
                        End If
                        
'                        If nCodFicha = 0 And !CodTributo <> 26 And !CodTributo <> 587 And !CodTributo <> 609 And !CodTributo <> 552 Then
'                            MsgBox !CodTributo
'                        End If
                                
                        
                        If nCodFicha > 0 Then
                           bAchou = False
                           For x = 1 To UBound(Matrix)
                               If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = nCodFicha And Matrix(x).CodBanco = nCodBanco Then
                                  bAchou = True
                                  Exit For
                               End If
                           Next
                           If bAchou Then
                              Matrix(x).ValorTotal = Matrix(x).ValorTotal + !ValorTributo
                           Else
                              ReDim Preserve Matrix(UBound(Matrix) + 1)
                              Matrix(UBound(Matrix)).DataReceita = cData
                              Matrix(UBound(Matrix)).CodBanco = nCodBanco
                              Matrix(UBound(Matrix)).CodTributo = !CodTributo
                              Matrix(UBound(Matrix)).NumFicha = nCodFicha
                              Matrix(UBound(Matrix)).ValorTotal = !ValorTributo
                           End If
                        End If
                        
                        If nValorTaxa > 0 Then
                           bAchou = False
                           For x = 1 To UBound(Matrix)
                               If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = 15 And Matrix(x).CodBanco = nCodBanco Then
                                  bAchou = True
                                  Exit For
                               End If
                           Next
                           If bAchou Then
                              Matrix(x).ValorTotal = Matrix(x).ValorTotal + nValorTaxa
                           Else
                              ReDim Preserve Matrix(UBound(Matrix) + 1)
                              Matrix(UBound(Matrix)).DataReceita = cData
                              Matrix(UBound(Matrix)).CodBanco = nCodBanco
                              Matrix(UBound(Matrix)).CodTributo = 3
                              Matrix(UBound(Matrix)).NumFicha = 15
                              Matrix(UBound(Matrix)).ValorTotal = nValorTaxa
                           End If
                        End If
                        
                        If nCodFichaJM > 0 Then
                           bAchou = False
                           For x = 1 To UBound(Matrix)
                               If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = nCodFichaJM And Matrix(x).CodBanco = nCodBanco Then
                                  bAchou = True
                                  Exit For
                               End If
                           Next
                           If bAchou Then
                              If ((nValorJMC / 3) * 2) > 0 Then
                                    Matrix(x).ValorTotal = Matrix(x).ValorTotal + ((nValorJMC / 3) * 2)
                              End If
                              
                           Else
                              ReDim Preserve Matrix(UBound(Matrix) + 1)
                              Matrix(UBound(Matrix)).DataReceita = cData
                              Matrix(UBound(Matrix)).CodBanco = nCodBanco
                              Matrix(UBound(Matrix)).CodTributo = !CodTributo
                              Matrix(UBound(Matrix)).NumFicha = nCodFichaJM
                              Matrix(UBound(Matrix)).ValorTotal = ((nValorJMC / 3) * 2)
                              If Matrix(UBound(Matrix)).ValorTotal < 0 Then
                                    Matrix(UBound(Matrix)).ValorTotal = 0
                              End If
                           End If
                        End If
                        If nCodFichaC > 0 Then
                           bAchou = False
                           For x = 1 To UBound(Matrix)
                               If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = nCodFichaC And Matrix(x).CodBanco = nCodBanco Then
                                  bAchou = True
                                  Exit For
                               End If
                           Next
                           If bAchou Then
                              Matrix(x).ValorTotal = Matrix(x).ValorTotal + (nValorJMC / 3)
                           Else
                              ReDim Preserve Matrix(UBound(Matrix) + 1)
                              Matrix(UBound(Matrix)).DataReceita = cData
                              Matrix(UBound(Matrix)).CodBanco = nCodBanco
                              Matrix(UBound(Matrix)).CodTributo = !CodTributo
                              Matrix(UBound(Matrix)).NumFicha = nCodFichaC
                              Matrix(UBound(Matrix)).ValorTotal = (nValorJMC / 3)
                           End If
                        End If
'                        If nCodFicha = 157 Then MsgBox "TESTE"
                       .MoveNext
                    Loop
                   .Close
                End With
proximo:
                DoEvents
                xId = xId + 1
               .MoveNext
            Loop
        End With
        CallPb xId, nNumRec
        'DIFERENCA
        nSomaDif = 0
        For x = 1 To UBound(Matrix)
            nSomaDif = nSomaDif + Matrix(x).ValorTotal
        Next
        nSomaTributo = nSomaPago - nSomaDif
        'resto
           If nSomaTributo < 0 Then nSomaTributo = 0
        
           nSomaRC = 0
           Sql = "SELECT * FROM RECEITACLASSIFICAR  WHERE DATARECEITA='" & Format(cData, "mm/dd/yyyy") & "' "
           Sql = Sql & " AND CODBANCO=" & aBanco(nContador)
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux2
                Do Until .EOF
                    nSomaRC = nSomaRC + !ValorTotal
                    nSomaTributo = nSomaTributo + !ValorTotal
                   .MoveNext
                Loop
               .Close
           End With
           
           nSomaDC = 0
           Sql = "SELECT DISTINCT CODBANCO, VALORCLASS FROM DEBITOCLASSIFICAR WHERE DATARECEITA='" & Format(cData, "mm/dd/yyyy") & "' "
           Sql = Sql & " AND CODBANCO=" & aBanco(nContador)
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux2
                Do Until .EOF
                    If Not IsNull(!VALORCLASS) Then
                        nSomaDC = nSomaDC + !VALORCLASS
                        nSomaTributo = nSomaTributo + !VALORCLASS
                    End If
                   .MoveNext
                Loop
           End With
           
           If nSomaTributo > 0 Then
'              ReDim Preserve Matrix(UBound(Matrix) + 1)
'              Matrix(UBound(Matrix)).DataReceita = cData
'              Matrix(UBound(Matrix)).CodBanco = aBanco(nContador)
'              Matrix(UBound(Matrix)).CodTributo = 500
'              Matrix(UBound(Matrix)).NumFicha = 50416
'              Matrix(UBound(Matrix)).ValorTotal = nSomaTributo
                bAchou = False
                For x = 1 To UBound(Matrix)
                    If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = 50416 And Matrix(x).CodBanco = aBanco(nContador) Then
                       bAchou = True
                       Exit For
                    End If
                Next
                If bAchou Then
                   Matrix(x).ValorTotal = Matrix(x).ValorTotal + nSomaTributo
                Else
                    'joga a 50416 no maior indice
                    nIndMaior = 0
                    For x = 1 To UBound(Matrix)
                        If Matrix(x).NumFicha = 2 Then
                            nIndMaior = x
                            Matrix(x).ValorTotal = Matrix(x).ValorTotal + nSomaTributo
                            Exit For
                        End If
                    Next
                    
                    If nIndMaior = 0 Then
                        Matrix(UBound(Matrix) - 1).ValorTotal = Matrix(UBound(Matrix) - 1).ValorTotal + nSomaTributo
                    End If
                    
'                   ReDim Preserve Matrix(UBound(Matrix) + 1)
'                   Matrix(UBound(Matrix)).DataReceita = cData
'                   Matrix(UBound(Matrix)).CodBanco = aBanco(nContador)
'                   Matrix(UBound(Matrix)).CodTributo = 500
'                   Matrix(UBound(Matrix)).NumFicha = 50416
'                   Matrix(UBound(Matrix)).ValorTotal = FormatNumber(nSomaTributo, 2)
                End If
           End If
           
           nTotalMatriz = 0
           If nSomaDif > nSomaPago Then
            nSomaMatriz = nSomaDif - nSomaPago
            nMaiorValor = 0
            For x = 1 To UBound(Matrix)
               nTotalMatriz = nTotalMatriz + Matrix(x).ValorTotal
               If Matrix(x).ValorTotal > nMaiorValor Then
                  nMaiorValor = Matrix(x).ValorTotal
                  nIndMaior = x
               End If
            Next
            
            If nSomaDC + nSomaRC + nSomaPago = nTotalMatriz Then
            Else
            'If nSomaMatriz < 1 Then
               Matrix(nIndMaior).ValorTotal = Matrix(nIndMaior).ValorTotal - nSomaMatriz
            End If
           End If
                      
  '          nSomaMatriz = 0
  '          nMaiorValor = 0
  '          For x = 1 To UBound(Matrix)
  '             nSomaMatriz = nSomaMatriz + Matrix(x).ValorTotal
  '             If Matrix(x).ValorTotal > nMaiorValor Then
  '                nMaiorValor = Matrix(x).ValorTotal
  '                nIndMaior = x
  '             End If
  '          Next
           
'            nSomaMatriz = 0
'            For x = 1 To UBound(Matrix)
'               If Matrix(x).ValorTotal < 0 Then
'                  nSomaMatriz = nSomaMatriz + Abs(Matrix(x).ValorTotal)
 '                 Matrix(x).ValorTotal = 0
  '             End If
 '           Next
'            If UBound(Matrix) > 0 Then
'                Matrix(nIndMaior).ValorTotal = Matrix(nIndMaior).ValorTotal - nSomaMatriz
 '           End If
            
        If cGetInputState() <> 0 Then DoEvents
        nSomaDif = 0
        For x = 1 To UBound(Matrix)
'            If Matrix(x).NumFicha = 143 Then MsgBox "TESTE"
            nSomaFicha = 0
            If CDbl(Matrix(x).ValorTotal) > 0 Then
               'BUSCA NATUREZA E VINCULO PARA CADA FICHA
               
                Sql = "SELECT NATUREZA,VINCULO,PERC FROM FICHACONTABIL WHERE FICHA=" & Matrix(x).NumFicha
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    Do Until .EOF
                        sNatureza = !Natureza
                        sVinculo = !Vinculo
                        If IsNull(!Perc) Then
                            MsgBox "% não cadastrado para ficha " & Matrix(x).NumFicha
                            nPerc = 100
                        Else
                            'nPerc = !Perc
                            nPerc = !Perc
                        End If
                        nValorTotal = Matrix(x).ValorTotal * nPerc / 100
                        nSomaFicha = nSomaFicha + nValorTotal
                       'GRAVA NA TABELA
                        Sql = "INSERT ANALISERECEITA (COMPUTER,DATARECEITA,CODBANCO,CODTRIBUTO,NUMFICHA,VALORTOTAL,NATUREZA,VINCULO) VALUES('" & NomeDeLogin & "','"
                        Sql = Sql & Format(Matrix(x).DataReceita, "mm/dd/yyyy") & "'," & Matrix(x).CodBanco & "," & Matrix(x).CodTributo & ","
                        Sql = Sql & Matrix(x).NumFicha & "," & Virg2Ponto(CStr(nValorTotal)) & ",'" & Trim$(sNatureza) & "','" & Trim$(sVinculo) & "')"
                        cn.Execute Sql, rdExecDirect
                        'ax = FillSpace(sNatureza, 20) & FillSpace(sVinculo, 20) & Year(Matrix(x).DataReceita) & Format(Month(Matrix(x).DataReceita), "00") & Format(Day(Matrix(x).DataReceita), "00") & Format(Matrix(x).CodBanco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(nValorTotal, 2)), ",", "")), "0000000000000")
                        'Print #1, ax
                       .MoveNext
                    Loop
                    
                   'ARREDONDA DENTRO DA TABELA PARA 2 DECIMAIS
                    Sql = "UPDATE ANALISERECEITA Set ValorTotal = Round(ValorTotal, 2) Where CODBANCO=" & Matrix(x).CodBanco & " AND  NumFicha = " & Matrix(x).NumFicha
                     'cn.Execute Sql, rdExecDirect
                   'APAGA AS QUE FOREM ZERO
                    Sql = "DELETE FROM ANALISERECEITA Where CODBANCO=" & Matrix(x).CodBanco & " AND NumFicha = " & Matrix(x).NumFicha & " AND VALORTOTAL=0"
                    cn.Execute Sql, rdExecDirect
                   'SOMA O VALOR TOTAL DA FICHA
                    Sql = "SELECT SUM(VALORTOTAL) AS VALORTOTAL from analisereceita Where CODBANCO=" & Matrix(x).CodBanco & " AND  NumFicha = " & Matrix(x).NumFicha
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If IsNull(!ValorTotal) Then
                            nValorTotal = 0
                        Else
                            nValorTotal = !ValorTotal
                        End If
                       .Close
                    End With
                   'GUARDA O NUMERO DO PRIMEIRO VINCULO QUE POSSUE VALOR, PARA JOGAR NELE A DIFERENÇA
                    Sql = "SELECT NATUREZA,VINCULO from analisereceita Where CODBANCO=" & Matrix(x).CodBanco & " AND NumFicha = " & Matrix(x).NumFicha
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            sNatureza = !Natureza
                            sVinculo = !Vinculo
                        Else
                            sNatureza = "9999"
                            sVinculo = "9999"
                            MsgBox "Ficha não encontrada " & Matrix(x).NumFicha & " Tributo: " & Matrix(x).CodTributo
                        End If
                       .Close
                    End With
                   'VERIFICA A DIFERENÇA ENTRE O QUE ESTA NA BASE E O TOTAL DA FICHA
                    'nSomaFicha = nSomaFicha
                    If nValorTotal > nSomaFicha Then
                        nDif = nValorTotal - nSomaFicha
                        Sql = "UPDATE ANALISERECEITA SET VALORTOTAL=VALORTOTAL - (" & Virg2Ponto(CStr(nDif)) & ") WHERE CODBANCO=" & Matrix(x).CodBanco & " AND  NUMFICHA=" & Matrix(x).NumFicha
                        Sql = Sql & " AND NATUREZA='" & sNatureza & "' AND VINCULO='" & sVinculo & "'"
                    ElseIf nValorTotal < nSomaFicha Then
                        nDif = Round(nSomaFicha - nValorTotal, 2)
                        Sql = "UPDATE ANALISERECEITA SET VALORTOTAL=VALORTOTAL + (" & Virg2Ponto(CStr(nDif)) & ") WHERE CODBANCO=" & Matrix(x).CodBanco & " AND NUMFICHA=" & Matrix(x).NumFicha
                        Sql = Sql & " AND NATUREZA='" & sNatureza & "' AND VINCULO='" & sVinculo & "'"
                    End If
                    cn.Execute Sql, rdExecDirect
                   .Close
                End With
            End If
        Next
        
    Next
Next

'Sql = "SELECT * FROM ANALISERECEITA WHERE COMPUTER='" & NomeDeLogin & "'"
Sql = "SELECT DISTINCT analisereceita.computer, analisereceita.datareceita, analisereceita.codbanco, analisereceita.codtributo, analisereceita.numficha,"
Sql = Sql & "analisereceita.valortotal, analisereceita.natureza, analisereceita.vinculo, analisereceita.codreduzido, analisereceita.anoexercicio,"
Sql = Sql & "analisereceita.CodLancamento , analisereceita.SeqLancamento, analisereceita.NumParcela, analisereceita.CODCOMPLEMENTO, fichacontabil.arq "
Sql = Sql & "FROM analisereceita INNER JOIN fichacontabil ON analisereceita.numficha = fichacontabil.ficha WHERE COMPUTER='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ax = FillSpace(!Natureza, 20) & FillSpace(!Vinculo, 20) & Year(!DataReceita) & Format(Month(!DataReceita), "00") & Format(Day(!DataReceita), "00") & Format(!CodBanco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(!ValorTotal, 2)), ",", "")), "0000000000000")
 '       If !ARQ = 1 Then
            Print #1, ax
 '       Else
 '           Print #2, ax
 '       End If
       .MoveNext
    Loop
   .Close
End With

Close #1
'Close #2
Pb.value = 100
lblMsg.Caption = "Analise concluida..."
lblMsg.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Sub cmdGerar_Click()

GeraAnalise2
Ocupado
frmReport.ShowReport "Analise", frmMdi.HWND, Me.HWND
Sql = "DELETE FROM ANALISERECEITA WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Liberado

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Centraliza Me

'Sql = "SELECT CODBANCO,NOMEBANCO FROM BANCO WHERE CODBANCO<>0 AND CODBANCO<>91 AND CODBANCO<>92 AND CODBANCO<>93 AND CODBANCO<>94 AND CODBANCO<>95 AND CODBANCO<>96 AND CODBANCO<>97 AND CODBANCO<>98  "
Sql = "SELECT CODBANCO,NOMEBANCO FROM BANCO WHERE CODBANCO<>0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbBanco.AddItem !NomeBanco
        cmbBanco.ItemData(cmbBanco.NewIndex) = !CodBanco
       .MoveNext
    Loop
End With
cmbBanco.ListIndex = 0
cmbBanco.Enabled = False

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = Int(Pb.value) & " %"

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
Resume Next
End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub GeraAnalise2()

Dim cData As Date, xId As Long, nNumRec As Long, nSomaFicha As Double, i As Integer, Matrix2() As Analise, nNumFicha As Long
Dim bDA As Boolean, bAj As Boolean, Matrix() As Analise, nCodFicha As Long, nValorPago As Double
Dim nCodFichaJM As Long, nCodFichaC As Long, x As Integer, bAchou As Boolean, ax As String, nPos As Integer
Dim nSomaTributo As Double, nDif As Double, nCodBanco As Integer, nPerc As Double, nValorJMC As Double
Dim nSomaPago As Double, nSomaDif As Double, nSomaMatriz As Double, nMaiorValor As Double, nIndMaior As Integer
Dim aBanco() As Integer, nContador As Integer, nContaBanco As Integer, t As Integer, sNumProc As String
Dim nSomaDC As Double, nSomaRC As Double, nTotalMatriz As Double, sNatureza As String, sVinculo As String, nValorTotal As Double


nSomaPago = 0: nSomaTotal = 0
cData = CDate(mskDataIni.Text)
cn.QueryTimeout = 0
Pb.value = 0
Sql = "DELETE FROM ANALISERECEITA WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

If chk1.value = 1 Then
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") For Output Shared As #1
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & "SD" For Output Shared As #2
Else
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") For Output Shared As #1
   Open sPathBin & "\REC" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(cmbBanco.ItemData(cmbBanco.ListIndex), "000") & "SD" For Output Shared As #2
End If

lblMsg.Caption = "Gerando analise do dia.: " & cData
lblMsg.Refresh
If cGetInputState() <> 0 Then DoEvents

If chk1.value = 1 Then
    Sql = "SELECT DISTINCT CODBANCO FROM DEBITOPAGO WHERE DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' AND RESTITUIDO IS NULL AND "
    Sql = Sql & "CODBANCO<>91 AND CODBANCO<>92 AND CODBANCO<>93 AND CODBANCO<>94 AND CODBANCO<>95 AND CODBANCO<>96 AND CODBANCO<>97 AND CODBANCO<>98 AND CODBANCO<>99"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        ReDim aBanco(0)
        Do Until .EOF
            ReDim Preserve aBanco(UBound(aBanco) + 1)
            aBanco(UBound(aBanco)) = !CodBanco
           .MoveNext
        Loop
       .Close
    End With
    nContaBanco = UBound(aBanco)

    Sql = "SELECT DISTINCT CODBANCO FROM RECEITACLASSIFICAR WHERE DATARECEITA ='" & Format(cData, "mm/dd/yyyy") & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            bAchou = False
            For t = 1 To UBound(aBanco)
                If aBanco(t) = !CodBanco Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                ReDim Preserve aBanco(UBound(aBanco) + 1)
                aBanco(UBound(aBanco)) = !CodBanco
            End If
           .MoveNext
        Loop
       .Close
    End With
    nContaBanco = UBound(aBanco)
Else
    ReDim aBanco(1)
    aBanco(1) = cmbBanco.ItemData(cmbBanco.ListIndex)
    nContaBanco = 1
End If

For nContador = 1 To nContaBanco
    ReDim Matrix(0): ReDim Matrix2(0)
    If chk1.value = 1 Then
    Else
        ReDim aBanco(1)
        aBanco(1) = cmbBanco.ItemData(cmbBanco.ListIndex)
    End If
    
    Sql = "DELETE FROM ANALISERECEITADETALHE WHERE COMPUTER='" & NomeDeLogin & "' AND CODBANCO=" & aBanco(nContador) & " AND DATARECEITA= '" & Format(cData, "mm/dd/yyyy") & "'"
    cn.Execute Sql, rdExecDirect
        
    Sql = "SELECT SUM(DEBITOPAGO.VALORPAGOREAL) AS VALORPAGOTOTAL FROM DEBITOPARCELA RIGHT OUTER JOIN DEBITOPAGO ON DEBITOPARCELA.CODREDUZIDO = DEBITOPAGO.CODREDUZIDO AND "
    Sql = Sql & "DEBITOPARCELA.AnoExercicio = DEBITOPAGO.AnoExercicio AND DEBITOPARCELA.CodLancamento = DEBITOPAGO.CodLancamento AND DEBITOPARCELA.SeqLancamento = DEBITOPAGO.SeqLancamento AND "
    Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOPAGO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOPAGO.CODCOMPLEMENTO "
    Sql = Sql & "WHERE VALORPAGOREAL>0 AND RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' AND (CODBANCO=" & aBanco(nContador)
    If aBanco(nContador) = 1 Then
        Sql = Sql & " OR CODBANCO=91 OR CODBANCO=92 OR CODBANCO=93 OR CODBANCO=94 OR CODBANCO=95 OR CODBANCO=96 OR CODBANCO=97 OR CODBANCO=98) "
    Else
        Sql = Sql & ")"
    End If
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         If Not IsNull(!VALORPAGOTOTAL) Then
             nSomaPago = !VALORPAGOTOTAL
         Else
            nSomaPago = 0
         End If
        .Close
    End With
                       
    Sql = "SELECT DISTINCT DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,"
    Sql = Sql & "DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.DATAINSCRICAO,DEBITOPARCELA.DATAAJUIZA,DEBITOPAGO.DATAPAGAMENTO,DEBITOPAGO.DATARECEBIMENTO,DEBITOPAGO.VALORPAGO,DEBITOPAGO.VALORPAGOREAL, DEBITOPAGO.CODBANCO,"
    Sql = Sql & "DEBITOPAGO.CodAgencia,DEBITOPAGO.RESTITUIDO FROM DEBITOPARCELA INNER JOIN DEBITOPAGO ON DEBITOPARCELA.CODREDUZIDO = DEBITOPAGO.CODREDUZIDO AND "
    Sql = Sql & "DEBITOPARCELA.AnoExercicio = DEBITOPAGO.AnoExercicio AND DEBITOPARCELA.CodLancamento = DEBITOPAGO.CodLancamento AND DEBITOPARCELA.SeqLancamento = DEBITOPAGO.SeqLancamento AND "
    Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOPAGO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOPAGO.CODCOMPLEMENTO "
    Sql = Sql & "WHERE VALORPAGOREAL>0 AND RESTITUIDO IS NULL AND DATARECEBIMENTO = '" & Format(cData, "mm/dd/yyyy") & "' AND (CODBANCO=" & aBanco(nContador)
    If aBanco(nContador) = 1 Then
        Sql = Sql & " OR CODBANCO=91 OR CODBANCO=92 OR CODBANCO=93 OR CODBANCO=94 OR CODBANCO=95 OR CODBANCO=96 OR CODBANCO=97 OR CODBANCO=98) "
    Else
        Sql = Sql & ")"
    End If
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nNumRec = .RowCount
        nDif = 0
        Do Until .EOF
            If xId Mod 10 = 0 Then
                CallPb xId, nNumRec
            End If
            nValorPago = FormatNumber(!valorpagoreal, 2)
            nCodBanco = aBanco(nContador)
            
                                
            
            
            If !CodLancamento = 20 Then
                '***  parcelamentos *****
                Sql = "SELECT NUMPROCESSO FROM DESTINOREPARC WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND NUMSEQUENCIA=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                With RdoAux2
                    If .RowCount = 0 Then
                        bDA = True: bAj = True
                        GoTo continua
                    End If
                    sNumProc = !numprocesso
                   .Close
                End With
                
                Sql = "SELECT * FROM ORIGEMREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                With RdoAux2
                    Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !numsequencia & " AND NUMPARCELA=" & !NumParcela & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                    Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    With RdoAux3
                        If Not IsNull(!datainscricao) Then
                           bDA = True
                        Else
                           bDA = False
                        End If
                        If Not IsNull(!dataajuiza) Then
                           bAj = True
                        Else
                           bAj = False
                        End If
                        .Close
                    End With
                    .Close
                End With
                '************************
            Else
                If Not IsNull(!datainscricao) Then
                   bDA = True
                Else
                   bDA = False
                End If
                If Not IsNull(!dataajuiza) Then
                   bAj = True
                Else
                   bAj = False
                End If
            End If
            
continua:
            
            Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                If Not IsNull(!soma) Then
                   nSomaTributo = FormatNumber(!soma, 2)
                Else
                   nSomaTributo = 0
                End If
               .Close
            End With
            
            If nValorPago >= nSomaTributo Then
                nDif = nValorPago - nSomaTributo
            Else
                nDif = nValorPago - nSomaTributo
            End If
            Sql = "SELECT DEBITOTRIBUTO.*, DEBITOTRIBUTO.CODTRIBUTO,DEBITOTRIBUTO.VALORTRIBUTO,VALORJUROS,VALORMULTA,VALORCORRECAO,TRIBUTO.FICHA,"
            Sql = Sql & "TRIBUTO.FICHAJRMULTA, TRIBUTO.FICHADIVIDA,TRIBUTO.FICHADAJRMUL,TRIBUTO.FICHADAENCA,TRIBUTO.FICHAAJUIZA,TRIBUTO.FICHAAJJRMUL,FICHAAJENCA "
            Sql = Sql & "FROM DEBITOTRIBUTO LEFT OUTER JOIN TRIBUTO ON DEBITOTRIBUTO.CODTRIBUTO = TRIBUTO.CODTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                nCodFicha = 0: nCodFichaJM = 0: nCodFichaC = 0
                Do Until .EOF
                    If nSomaTributo = 0 Then GoTo proximo
                    nPerc = !ValorTributo / nSomaTributo 'PRINCIPAL / SOMA DOS TRIBUTOS
                    If nPerc < 0 Then MsgBox "aqui"
                    nValorJMC = nDif * nPerc 'DIFERENCA * PERCENTUAL (JUROS,MULTA E CORRECAO)
                    If Not bAj Then
                        If Not bDA Then
                            nCodFicha = Val(SubNull(!Ficha))
                            nCodFichaJM = Val(SubNull(!FichaJrMulta))
                        Else
                            nCodFicha = Val(SubNull(!FichaDivida))
                            nCodFichaJM = Val(SubNull(!FichaDaJrMul))
                            nCodFichaC = Val(SubNull(!FichaDaEnca))
                        End If
                    Else
                        nCodFicha = Val(SubNull(!FichaAjuiza))
                        nCodFichaJM = Val(SubNull(!FichaAjJrMul))
                        nCodFichaC = Val(SubNull(!FichaAjEnca))
                    End If
                    
                    
                    If nCodFicha > 0 Then
                       bAchou = False
                       For x = 1 To UBound(Matrix)
                           If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = nCodFicha And Matrix(x).CodBanco = nCodBanco Then
                              bAchou = True
                              Exit For
                           End If
                       Next
                       If bAchou Then
                          Matrix(x).ValorTotal = Matrix(x).ValorTotal + FormatNumber(!ValorTributo, 2)
                       Else
                          ReDim Preserve Matrix(UBound(Matrix) + 1)
                          Matrix(UBound(Matrix)).DataReceita = cData
                          Matrix(UBound(Matrix)).CodBanco = nCodBanco
                          Matrix(UBound(Matrix)).CodTributo = !CodTributo
                          Matrix(UBound(Matrix)).NumFicha = nCodFicha
                          Matrix(UBound(Matrix)).ValorTotal = FormatNumber(!ValorTributo, 2)
                          Matrix(UBound(Matrix)).nCodReduz = !CODREDUZIDO
                          Matrix(UBound(Matrix)).nAno = !AnoExercicio
                          Matrix(UBound(Matrix)).nLanc = !CodLancamento
                          Matrix(UBound(Matrix)).nSeq = !SeqLancamento
                          Matrix(UBound(Matrix)).nParc = !NumParcela
                          Matrix(UBound(Matrix)).nCompl = !CODCOMPLEMENTO
                       End If
                    End If
                    If nCodFichaJM > 0 Then
                       bAchou = False
                       For x = 1 To UBound(Matrix)
                           If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = nCodFichaJM And Matrix(x).CodBanco = nCodBanco Then
                              bAchou = True
                              Exit For
                           End If
                       Next
                       If bAchou Then
                          If ((nValorJMC / 3) * 2) > 0 Then
                                Matrix(x).ValorTotal = FormatNumber(Matrix(x).ValorTotal + ((nValorJMC / 3) * 2), 2)
                          End If
                       Else
                          nNumFicha = nCodFichaJM
                          ReDim Preserve Matrix(UBound(Matrix) + 1)
                          Matrix(UBound(Matrix)).DataReceita = cData
                          Matrix(UBound(Matrix)).CodBanco = nCodBanco
                          Matrix(UBound(Matrix)).CodTributo = !CodTributo
                          Matrix(UBound(Matrix)).NumFicha = nCodFichaJM
                          Matrix(UBound(Matrix)).ValorTotal = FormatNumber(((nValorJMC / 3) * 2), 2)
                          If Matrix(UBound(Matrix)).ValorTotal < 0 Then
                                Matrix(UBound(Matrix)).ValorTotal = 0
                          End If
                          Matrix(UBound(Matrix)).nCodReduz = !CODREDUZIDO
                          Matrix(UBound(Matrix)).nAno = !AnoExercicio
                          Matrix(UBound(Matrix)).nLanc = !CodLancamento
                          Matrix(UBound(Matrix)).nSeq = !SeqLancamento
                          Matrix(UBound(Matrix)).nParc = !NumParcela
                          Matrix(UBound(Matrix)).nCompl = !CODCOMPLEMENTO
                       End If
                    End If
                    If nCodFichaC > 0 Then
                       bAchou = False
                       For x = 1 To UBound(Matrix)
                           If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = nCodFichaC And Matrix(x).CodBanco = nCodBanco Then
                              bAchou = True
                              Exit For
                           End If
                       Next
                       If bAchou Then
                          Matrix(x).ValorTotal = Matrix(x).ValorTotal + FormatNumber((nValorJMC / 3), 2)
                       Else
                          nNumFicha = nCodFichaC
                          ReDim Preserve Matrix(UBound(Matrix) + 1)
                          Matrix(UBound(Matrix)).DataReceita = cData
                          Matrix(UBound(Matrix)).CodBanco = nCodBanco
                          Matrix(UBound(Matrix)).CodTributo = !CodTributo
                          Matrix(UBound(Matrix)).NumFicha = nCodFichaC
                          Matrix(UBound(Matrix)).ValorTotal = FormatNumber((nValorJMC / 3), 2)
                          Matrix(UBound(Matrix)).nCodReduz = !CODREDUZIDO
                          Matrix(UBound(Matrix)).nAno = !AnoExercicio
                          Matrix(UBound(Matrix)).nLanc = !CodLancamento
                          Matrix(UBound(Matrix)).nSeq = !SeqLancamento
                          Matrix(UBound(Matrix)).nParc = !NumParcela
                          Matrix(UBound(Matrix)).nCompl = !CODCOMPLEMENTO
                       End If
                    End If
                   .MoveNext
                Loop
               .Close
            End With
proximo:
            xId = xId + 1
           .MoveNext
        Loop
    End With
    CallPb xId, nNumRec
    'DIFERENCA
    nSomaDif = 0
    For x = 1 To UBound(Matrix)
        nSomaDif = nSomaDif + Matrix(x).ValorTotal
    Next
    nSomaTributo = FormatNumber(nSomaPago, 2) - FormatNumber(nSomaDif, 2)
    'resto
       If nSomaTributo < 0 Then nSomaTributo = 0
    
       nSomaRC = 0
       Sql = "SELECT * FROM RECEITACLASSIFICAR  WHERE DATARECEITA='" & Format(cData, "mm/dd/yyyy") & "' "
       Sql = Sql & " AND CODBANCO=" & aBanco(nContador)
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            Do Until .EOF
                nSomaRC = nSomaRC + !ValorTotal
                nSomaTributo = nSomaTributo + !ValorTotal
               .MoveNext
            Loop
           .Close
       End With
       
       nSomaDC = 0
       Sql = "SELECT DISTINCT CODBANCO, VALORCLASS FROM DEBITOCLASSIFICAR WHERE DATARECEITA='" & Format(cData, "mm/dd/yyyy") & "' "
       Sql = Sql & " AND CODBANCO=" & aBanco(nContador)
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            Do Until .EOF
                If Not IsNull(!VALORCLASS) Then
                    nSomaDC = nSomaDC + !VALORCLASS
                    nSomaTributo = FormatNumber(nSomaTributo + !VALORCLASS, 2)
                End If
               .MoveNext
            Loop
       End With
       
       If nSomaTributo > 0 Then
'              ReDim Preserve Matrix(UBound(Matrix) + 1)
'              Matrix(UBound(Matrix)).DataReceita = cData
'              Matrix(UBound(Matrix)).CodBanco = aBanco(nContador)
'              Matrix(UBound(Matrix)).CodTributo = 500
'              Matrix(UBound(Matrix)).NumFicha = 50416
'              Matrix(UBound(Matrix)).ValorTotal = nSomaTributo
            bAchou = False
            For x = 1 To UBound(Matrix)
                If Matrix(x).DataReceita = cData And Matrix(x).NumFicha = 50416 And Matrix(x).CodBanco = aBanco(nContador) Then
                   bAchou = True
                   Exit For
                End If
            Next
            If bAchou Then
               Matrix(x).ValorTotal = Matrix(x).ValorTotal + nSomaTributo
            Else
                nNumFicha = 50416
                ReDim Preserve Matrix(UBound(Matrix) + 1)
                Matrix(UBound(Matrix)).DataReceita = cData
                Matrix(UBound(Matrix)).CodBanco = aBanco(nContador)
                Matrix(UBound(Matrix)).CodTributo = 500
                Matrix(UBound(Matrix)).NumFicha = 50416
                Matrix(UBound(Matrix)).ValorTotal = FormatNumber(nSomaTributo, 2)
            End If
       End If
       
       nTotalMatriz = 0
       If nSomaDif > nSomaPago Then
        nSomaMatriz = nSomaDif - nSomaPago
        nMaiorValor = 0
        For x = 1 To UBound(Matrix)
           nTotalMatriz = nTotalMatriz + Matrix(x).ValorTotal
           If Matrix(x).ValorTotal > nMaiorValor Then
              nMaiorValor = Matrix(x).ValorTotal
              nIndMaior = x
           End If
        Next
        
        If Round(nSomaDC + nSomaRC + nSomaPago, 2) = Round(nTotalMatriz, 2) Then
        Else
        'If nSomaMatriz < 1 Then
           Matrix(nIndMaior).ValorTotal = Matrix(nIndMaior).ValorTotal - nSomaMatriz
        End If
       End If
                  
        nSomaMatriz = 0
        nMaiorValor = 0
        For x = 1 To UBound(Matrix)
           nSomaMatriz = nSomaMatriz + Matrix(x).ValorTotal
           If Matrix(x).ValorTotal > nMaiorValor Then
              nMaiorValor = Matrix(x).ValorTotal
              nIndMaior = x
           End If
        Next
       
        nSomaMatriz = 0
        For x = 1 To UBound(Matrix)
           If Matrix(x).ValorTotal < 0 Then
              nSomaMatriz = nSomaMatriz + Abs(Matrix(x).ValorTotal)
              Matrix(x).ValorTotal = 0
           End If
        Next
        Matrix(nIndMaior).ValorTotal = Matrix(nIndMaior).ValorTotal - nSomaMatriz
        
    nSomaDif = 0
    For x = 1 To UBound(Matrix)
        nSomaFicha = 0
        If CDbl(Matrix(x).ValorTotal) > 0 Then
           'BUSCA NATUREZA E VINCULO PARA CADA FICHA
            Sql = "SELECT NATUREZA,VINCULO,PERC FROM FICHACONTABIL WHERE FICHA=" & Matrix(x).NumFicha
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                    sNatureza = !Natureza
                    sVinculo = !Vinculo
                    If IsNull(!Perc) Then
                        MsgBox "% não cadastrado para ficha " & Matrix(x).NumFicha
                        nPerc = 100
                    Else
                        nPerc = !Perc
                    End If
                    nValorTotal = Matrix(x).ValorTotal * nPerc / 100
                    nSomaFicha = nSomaFicha + nValorTotal
                   'GRAVA NA TABELA
                    Sql = "INSERT ANALISERECEITA (COMPUTER,DATARECEITA,CODBANCO,CODTRIBUTO,NUMFICHA,VALORTOTAL,NATUREZA,VINCULO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO) VALUES('" & NomeDeLogin & "','"
                    Sql = Sql & Format(Matrix(x).DataReceita, "mm/dd/yyyy") & "'," & Matrix(x).CodBanco & "," & Matrix(x).CodTributo & ","
                    Sql = Sql & Matrix(x).NumFicha & "," & Virg2Ponto(CStr(nValorTotal)) & ",'" & Trim$(sNatureza) & "','" & Trim$(sVinculo) & "'," & Matrix(x).nCodReduz & "," & Matrix(x).nAno & ","
                    Sql = Sql & Matrix(x).nLanc & "," & Matrix(x).nSeq & "," & Matrix(x).nParc & "," & Matrix(x).nCompl & ")"
                    cn.Execute Sql, rdExecDirect
                    'ax = FillSpace(sNatureza, 20) & FillSpace(sVinculo, 20) & Year(Matrix(x).DataReceita) & Format(Month(Matrix(x).DataReceita), "00") & Format(Day(Matrix(x).DataReceita), "00") & Format(Matrix(x).CodBanco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(nValorTotal, 2)), ",", "")), "0000000000000")
                    'Print #1, ax
                   .MoveNext
                Loop
                
               'ARREDONDA DENTRO DA TABELA PARA 2 DECIMAIS
                Sql = "UPDATE ANALISERECEITA Set ValorTotal = Round(ValorTotal, 2) Where CODBANCO=" & Matrix(x).CodBanco & " AND  NumFicha = " & Matrix(x).NumFicha
                cn.Execute Sql, rdExecDirect
               'APAGA AS QUE FOREM ZERO
                Sql = "DELETE FROM ANALISERECEITA Where CODBANCO=" & Matrix(x).CodBanco & " AND NumFicha = " & Matrix(x).NumFicha & " AND VALORTOTAL=0"
                cn.Execute Sql, rdExecDirect
               'SOMA O VALOR TOTAL DA FICHA
                Sql = "SELECT SUM(VALORTOTAL) AS VALORTOTAL from analisereceita Where CODBANCO=" & Matrix(x).CodBanco & " AND  NumFicha = " & Matrix(x).NumFicha
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If IsNull(!ValorTotal) Then
                        nValorTotal = 0
                    Else
                        nValorTotal = Round(!ValorTotal, 2)
                    End If
                   .Close
                End With
               'GUARDA O NUMERO DO PRIMEIRO VINCULO QUE POSSUE VALOR, PARA JOGAR NELE A DIFERENÇA
                Sql = "SELECT NATUREZA,VINCULO from analisereceita Where CODBANCO=" & Matrix(x).CodBanco & " AND NumFicha = " & Matrix(x).NumFicha
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        sNatureza = !Natureza
                        sVinculo = !Vinculo
                    Else
                        sNatureza = "9999"
                        sVinculo = "9999"
                        MsgBox "Ficha não encontrada " & Matrix(x).NumFicha & " Tributo: " & Matrix(x).CodTributo
                    End If
                   .Close
                End With
               'VERIFICA A DIFERENÇA ENTRE O QUE ESTA NA BASE E O TOTAL DA FICHA
                nSomaFicha = Round(nSomaFicha, 2)
                If nValorTotal > nSomaFicha Then
                    nDif = Round(nValorTotal - nSomaFicha, 2)
                    Sql = "UPDATE ANALISERECEITA SET VALORTOTAL=VALORTOTAL - (" & Virg2Ponto(CStr(nDif)) & ") WHERE CODBANCO=" & Matrix(x).CodBanco & " AND  NUMFICHA=" & Matrix(x).NumFicha
                    Sql = Sql & " AND NATUREZA='" & sNatureza & "' AND VINCULO='" & sVinculo & "'"
                ElseIf nValorTotal < nSomaFicha Then
                    nDif = Round(nSomaFicha - nValorTotal, 2)
                    Sql = "UPDATE ANALISERECEITA SET VALORTOTAL=VALORTOTAL + (" & Virg2Ponto(CStr(nDif)) & ") WHERE CODBANCO=" & Matrix(x).CodBanco & " AND NUMFICHA=" & Matrix(x).NumFicha
                    Sql = Sql & " AND NATUREZA='" & sNatureza & "' AND VINCULO='" & sVinculo & "'"
                End If
                cn.Execute Sql, rdExecDirect
               .Close
            End With
        End If
    Next
    
Next

Sql = "SELECT DISTINCT analisereceita.computer, analisereceita.datareceita, analisereceita.codbanco, analisereceita.codtributo, analisereceita.numficha,"
Sql = Sql & "analisereceita.valortotal, analisereceita.natureza, analisereceita.vinculo, analisereceita.codreduzido, analisereceita.anoexercicio,"
Sql = Sql & "analisereceita.CodLancamento , analisereceita.SeqLancamento, analisereceita.NumParcela, analisereceita.CODCOMPLEMENTO, fichacontabil.arq "
Sql = Sql & "FROM analisereceita INNER JOIN fichacontabil ON analisereceita.numficha = fichacontabil.ficha WHERE COMPUTER='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ax = FillSpace(!Natureza, 20) & FillSpace(!Vinculo, 20) & Year(!DataReceita) & Format(Month(!DataReceita), "00") & Format(Day(!DataReceita), "00") & Format(!CodBanco, "0000") & "00000000" & Format(RemovePonto(Replace(CStr(FormatNumber(!ValorTotal, 2)), ",", "")), "0000000000000") & "0000000000" & Format(Now, "ddmmhhmm")
        If !ARQ = 1 Then
            Print #1, ax
        Else
            Print #2, ax
        End If
       .MoveNext
    Loop
   .Close
End With

For x = 1 To UBound(Matrix2)
    With Matrix2(x)
        Sql = "INSERT ANALISERECEITADETALHE (COMPUTER,DATARECEITA,CODBANCO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,NUMFICHA,VALORTRIBUTO,VALORTOTAL) VALUES('" & NomeDeLogin & "','"
        Sql = Sql & Format(.DataReceita, "mm/dd/yyyy") & "'," & .CodBanco & "," & .nCodReduz & "," & .nAno & "," & .nLanc & "," & .nSeq & "," & .nParc & "," & .nCompl & "," & .CodTributo & "," & .NumFicha & "," & Virg2Ponto(CStr(.ValorTributo)) & ","
        Sql = Sql & Virg2Ponto(CStr(.ValorTotal)) & ")"
'        cn.Execute Sql, rdExecDirect
    End With
Next

Close #1
Close #2
Pb.value = 100
lblMsg.Caption = "Analise concluida..."
lblMsg.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

