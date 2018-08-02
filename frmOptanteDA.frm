VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmOptanteDA 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Optantes por Débito Automático"
   ClientHeight    =   5370
   ClientLeft      =   11280
   ClientTop       =   3780
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   8835
   Begin VB.ComboBox cmbDia 
      Height          =   315
      ItemData        =   "frmOptanteDA.frx":0000
      Left            =   4050
      List            =   "frmOptanteDA.frx":0061
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4950
      Width           =   780
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmOptanteDA.frx":00D8
      Left            =   7515
      List            =   "frmOptanteDA.frx":00DF
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4950
      Width           =   1005
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmOptanteDA.frx":00E9
      Left            =   5625
      List            =   "frmOptanteDA.frx":0111
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   4950
      Width           =   780
   End
   Begin VB.CommandButton cmdCorrige 
      Caption         =   "Corrige"
      Height          =   315
      Left            =   3450
      TabIndex        =   9
      Top             =   5625
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox txtParcela 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2175
      TabIndex        =   7
      Top             =   5610
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Atualiza Vencimento"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   5595
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ComboBox cmbBanco 
      Height          =   315
      ItemData        =   "frmOptanteDA.frx":013C
      Left            =   720
      List            =   "frmOptanteDA.frx":0152
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   90
      Width           =   3465
   End
   Begin MSFlexGridLib.MSFlexGrid grdOpt 
      Height          =   4215
      Left            =   30
      TabIndex        =   0
      Top             =   570
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorSel    =   12582912
      ForeColorSel    =   16777215
      FocusRect       =   0
      SelectionMode   =   1
      BorderStyle     =   0
      FormatString    =   $"frmOptanteDA.frx":01CA
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   240
      Left            =   645
      TabIndex        =   3
      Top             =   5010
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   7650
      TabIndex        =   10
      ToolTipText     =   "Sair da Tela"
      Top             =   120
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
      MICON           =   "frmOptanteDA.frx":0291
      PICN            =   "frmOptanteDA.frx":02AD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Height          =   315
      Left            =   5895
      TabIndex        =   11
      ToolTipText     =   "Cancelar Edição"
      Top             =   120
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gerar Arquivo"
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
      MICON           =   "frmOptanteDA.frx":031B
      PICN            =   "frmOptanteDA.frx":0337
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton chameleonButton1 
      Height          =   315
      Left            =   7515
      TabIndex        =   17
      ToolTipText     =   "Gerar Arquivo"
      Top             =   5895
      Visible         =   0   'False
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Gerar Arquivo"
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
      MICON           =   "frmOptanteDA.frx":03D6
      PICN            =   "frmOptanteDA.frx":03F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblSoma 
      Caption         =   "0 Contribuintes"
      Height          =   195
      Left            =   4230
      TabIndex        =   19
      Top             =   180
      Width           =   1545
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dia..:"
      Height          =   195
      Index           =   1
      Left            =   3465
      TabIndex        =   18
      Top             =   4995
      Width           =   510
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   195
      Left            =   6795
      TabIndex        =   15
      Top             =   4995
      Width           =   510
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mês..:"
      Height          =   195
      Index           =   0
      Left            =   5040
      TabIndex        =   13
      Top             =   4995
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcela"
      Enabled         =   0   'False
      Height          =   225
      Left            =   1455
      TabIndex        =   8
      Top             =   5640
      Width           =   705
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   3015
      TabIndex        =   5
      Top             =   5025
      Width           =   630
   End
   Begin VB.Label lblPF 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   60
      TabIndex        =   4
      Top             =   5025
      Width           =   390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco: "
      Height          =   225
      Left            =   90
      TabIndex        =   1
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "frmOptanteDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String

Private Sub chameleonButton1_Click()
Dim x As Integer, ax As String, RegA As String, RegZ As String, nCodReduz As Long
Dim nSoma As Double, sNomeArq As String, nNumParc As Integer, nSoma2 As Double, sSoma As String
Dim nSeq237 As Integer, nSeq341 As Integer, nSeq409 As Integer, nSeq399 As Integer, nSeq033 As Integer
Dim nContaLinha As Integer, sInsc As String, nValorTxExp As Double, nSeq104 As Integer
Dim xImovel As New clsImovel

If MsgBox("Gerar o arquivo de dEbito automAtico para o Banco?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

nNumParc = Val(txtParcela.Text)

'If nNumParc = 0 Then
'    MsgBox "PARCELA INVALIDA."
'    Exit Sub
'End If

If Val(Left$(cmbBanco.Text, 3)) = 237 Then 'BRADESCO
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ237'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq237 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("00538", "20") & "PREF.MUN.JABOTICABAL" & "237" & FillSpace("BRADESCO", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq237, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "DA" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Right$(Year(Now), 2), "00") & ".REM"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 341 Then 'ITAU
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ341'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq341 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("0905003200094", 20) & "PREF.MUN.JABOTICABAL" & "341" & FillSpace("ITAU", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq341, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "T341" & Format(nNumParc, "00") & Format(nSeq341, "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 409 Then 'UNIBANCO
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ409'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq409 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("20071987118022402", 20) & "PREF.MUN.JABOTICABAL" & "409" & FillSpace("UNIBANCO", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq409, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "T409" & Format(nNumParc, "00") & Format(nSeq409, "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 399 Then 'HSBC
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ399'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq399 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("00005012210", 20) & "PREF.MUN.JABOTICABAL" & "399" & FillSpace("HSBC", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq399, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "H399" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 33 Then 'BANESPA
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ033'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq033 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("00330023008000007463", 20) & "PREF.MUN.JABOTICABAL" & "033" & FillSpace("SANTANDER BANESPA", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq033, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "B033" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 104 Then 'CAIXA FEDERAL
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ104'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq104 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace(" ", 20) & "PREF.MUN.JABOTICABAL" & "104" & FillSpace("CAIXA ECON FEDERAL", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq104, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "C104" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
End If

Sql = "SELECT valorparcela From expediente WHERE (codlancamento = 1) AND (anoexped = " & Year(Now) & ")"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValorTxExp = FormatNumber(RdoAux!VALORPARCELA, 2)

nContaLinha = 0
nSoma = 0
Ocupado
'sPathArqDA = "c:\tmp"
Open sPathArqDA & "\" & sNomeArq For Output As #1
With grdOpt
   .SetFocus
    Print #1, RegA
    For x = 1 To .Rows - 1
        If x Mod 5 = 0 Then
           CallPb CLng(x), CLng(.Rows - 1)
        End If
'        nCodReduz = .TextMatrix(x, 0)
'        If nCodReduz = 1399 Then MsgBox "teste"
        Sql = "SELECT STATUSLANC FROM DEBITOPARCELA WHERE CODREDUZIDO = " & .TextMatrix(x, 0) & " AND year(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND MONTH(DATAVENCIMENTO)=" & Val(cmbMes.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If !statuslanc = 1 Then
                    GoTo proximo
                End If
            Else
                GoTo proximo
            End If
           .Close
        End With
        
'        Sql = "SELECT SUM(DEBITOTRIBUTO.VALORTRIBUTO) AS TOTAL, DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.DATAVENCIMENTO FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO AND "
'        Sql = Sql & "DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND "
'        Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOTRIBUTO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO "
'        Sql = Sql & " where DEBITOPARCELA.CODREDUZIDO = " & Val(.TextMatrix(x, 0)) & " and DEBITOPARCELA.CODLANCAMENTO=1 AND YEAR(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND MONTH(DATAVENCIMENTO)=" & Val(cmbMes.Text) & " AND DEBITOPARCELA.STATUSLANC=3 AND DEBITOTRIBUTO.CODTRIBUTO<>3 GROUP BY DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.DATAVENCIMENTO"
        
        Sql = "SELECT SUM(DEBITOTRIBUTO.VALORTRIBUTO) AS TOTAL, DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.CODCOMPLEMENTO,"
        Sql = Sql & "DEBITOPARCELA.DATAVENCIMENTO FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO AND "
        Sql = Sql & "DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND "
        Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOTRIBUTO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO "
        Sql = Sql & " where DEBITOPARCELA.CODREDUZIDO = " & Val(.TextMatrix(x, 0)) & " and DEBITOPARCELA.CODLANCAMENTO=1 AND YEAR(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND MONTH(DATAVENCIMENTO)=" & Val(cmbMes.Text) & " AND DAY(DATAVENCIMENTO)=" & Val(cmbDia.Text) & " AND DEBITOPARCELA.STATUSLANC=3 AND DEBITOTRIBUTO.CODTRIBUTO<>3 AND NUMPARCELA>0  "
        Sql = Sql & "GROUP BY DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.ANOEXERCICIO,DEBITOPARCELA.CODLANCAMENTO,DEBITOPARCELA.SEQLANCAMENTO,DEBITOPARCELA.NUMPARCELA,DEBITOPARCELA.CODCOMPLEMENTO,DEBITOPARCELA.DATAVENCIMENTO"
        
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               sInsc = Format(!AnoExercicio, "0000") & Format(!CodLancamento, "00") & Format(!SeqLancamento, "00") & Format(!NumParcela, "00") & Format(!CODCOMPLEMENTO, "0")
               grdOpt.TextMatrix(x, 5) = Format(!DataVencimento, "dd/mm/yyyy")
               grdOpt.TextMatrix(x, 6) = FormatNumber(!Total + nValorTxExp, 2)
               nSoma = nSoma + !Total + nValorTxExp
                                                      
'               With xImovel
'                    .CarregaImovel Val(grdOpt.TextMatrix(x, 0))
'                    sInsc = Format$(RemovePonto(Left$(.Inscricao, 18)), "000000000000000")
'               End With
                                                      
               'ax = "E" & FillSpace(grdOpt.TextMatrix(x, 0), 25) & Format(grdOpt.TextMatrix(x, 2), "0000") & FillSpace(grdOpt.TextMatrix(x, 3), 14)
               ax = "E" & FillSpace(sInsc, 25) & Format(grdOpt.TextMatrix(x, 2), "0000") & FillSpace(grdOpt.TextMatrix(x, 3), 14)
               ax = ax & Right$(grdOpt.TextMatrix(x, 5), 4) & Mid$(grdOpt.TextMatrix(x, 5), 4, 2) & Left$(grdOpt.TextMatrix(x, 5), 2)
               ax = ax & Format(RetornaNumero(grdOpt.TextMatrix(x, 6)), "000000000000000") & "03" & FillSpace(grdOpt.TextMatrix(x, 0), 60) & FillSpace(" ", 20) & "0"
               Print #1, ax
               nContaLinha = nContaLinha + 1
            End If
           .Close
        End With
proximo:
    Next
End With
sSoma = CStr(nSoma)
sSoma = FormatNumber(sSoma, 2)
If Left(Right(sSoma, 2), 1) = "," Then
    nSoma = RetornaNumero(sSoma) * 10
Else
    nSoma = RetornaNumero(sSoma)
End If

nContaLinha = nContaLinha + 2

RegZ = "Z" & Format(nContaLinha, "000000") & Format(nSoma, "00000000000000000") & FillSpace(" ", 126)
Print #1, RegZ
Close #1

If Val(Left$(cmbBanco.Text, 3)) = 237 Then 'BRADESCO
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ237'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 341 Then 'ITAU
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ341'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 409 Then 'UNIBANCO
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ409'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 399 Then 'HSBC
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ399'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 33 Then 'BANESPA
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ033'"
    cn.Execute Sql, rdExecDirect
End If

Liberado
MsgBox "FIM", vbInformation, "Atenção"

End Sub

Private Sub cmbBanco_Click()

If cmbBanco.ListIndex > -1 Then
   CarregaLista Val(Left$(cmbBanco.Text, 3)), 1
End If

lblSoma.Caption = CStr(grdOpt.Rows - 1) & " Contribuintes"

End Sub



Private Sub cmdConsultar_Click()
Unload Me
End Sub

Private Sub CarregaLista(nBanco As Integer, nOrder As Integer)

Ocupado
grdOpt.Rows = 1
Sql = "SELECT CODREDUZ,NOMECIDADAO,CODAGENCIA,NUMEROCONTA,DATAOPCAO,CODIGOPREF FROM VWDEBITOAUTOMATICO "
Sql = Sql & "WHERE CODBANCO=" & nBanco & " ORDER BY "
Select Case nOrder
    Case 0
        Sql = Sql & "CODREDUZ"
    Case 1
        Sql = Sql & "NOMECIDADAO"
    Case 2
        Sql = Sql & "CODAGENCIA"
    Case 3
        Sql = Sql & "NUMEROCONTA"
    Case 4
        Sql = Sql & "DATAOPCAO"
End Select

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         grdOpt.AddItem Format(!CodReduz, "0000000") & Chr(9) & !nomecidadao & Chr(9) & !CodAgencia & Chr(9) & !NumeroConta & Chr(9) & Format(!DataOpcao, "dd/mm/yyyy") & Chr(9) & Chr(9) & Chr(9) & !codigopref
        .MoveNext
    Loop
End With

Liberado

End Sub

Private Sub cmdCorrige_Click()
Dim cnAc As rdoConnection, RdoAux As rdoResultset
'CONECTA
Screen.MousePointer = vbHourglass
Exit Sub
Set en = rdoEngine.rdoEnvironments(0)
en.CursorDriver = rdUseServer
With en
    .CursorDriver = rdUseOdbc
    .LoginTimeout = 30

    Set cnAc = en.OpenConnection(dsname:="BANCO", _
        Prompt:=rdDriverNoPrompt, _
        Connect:="uid=Admin;")
End With

Sql = "SELECT DISTINCT CODIGO FROM UNIBANCO"
Set RdoAux = cnAc.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format("22/07/2004", "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & !Codigo
        Sql = Sql & " AND ANOEXERCICIO=2004 AND CODLANCAMENTO=1 AND SEQLANCAMENTO=0 AND NUMPARCELA=6 AND CODCOMPLEMENTO=0 AND SEQPAG>0"
'        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
End With

Sql = "SELECT DISTINCT CODIGO FROM ITAU"
Set RdoAux = cnAc.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format("22/07/2004", "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & !Codigo
        Sql = Sql & " AND ANOEXERCICIO=2004 AND CODLANCAMENTO=1 AND SEQLANCAMENTO=0 AND NUMPARCELA=6 AND CODCOMPLEMENTO=0 AND SEQPAG>0"
 '       cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
End With

Sql = "SELECT DISTINCT CODIGO FROM BRADESCO"
Set RdoAux = cnAc.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format("22/07/2004", "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & !Codigo
        Sql = Sql & " AND ANOEXERCICIO=2004 AND CODLANCAMENTO=1 AND SEQLANCAMENTO=0 AND NUMPARCELA=6 AND CODCOMPLEMENTO=0 AND SEQPAG>0"
'        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
End With


MsgBox "FIM"

'DESCONECTA
Screen.MousePointer = vbDefault
cnAc.Close
End Sub

Private Sub cmdGerar_Click()
Dim x As Integer, ax As String, RegA As String, RegZ As String, nCodReduz As Long
Dim nSoma As Double, sNomeArq As String, nNumParc As Integer, nSoma2 As Double, sSoma As String
Dim nSeq237 As Integer, nSeq341 As Integer, nSeq409 As Integer, nSeq399 As Integer, nSeq033 As Integer, nSeq001 As Integer, nSeq104 As Integer
Dim nContaLinha As Integer, sInsc As String, nValorTxExp As Double
Dim xImovel As New clsImovel, nCodPref As Long

If MsgBox("Gerar o arquivo de Débito automático para o Banco?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

nNumParc = Val(txtParcela.Text)

'If nNumParc = 0 Then
'    MsgBox "PARCELA INVALIDA."
'    Exit Sub
'End If

If Val(Left$(cmbBanco.Text, 3)) = 237 Then 'BRADESCO
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ237'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq237 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("00538", "20") & "PREF.MUN.JABOTICABAL" & "237" & FillSpace("BRADESCO", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq237, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "DA" & Format(Day(Now), "00") & Format(Month(Now), "00") & Format(Right$(Year(Now), 2), "00") & ".REM"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 341 Then 'ITAU
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ341'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq341 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("0905003200094", 20) & "PREF.MUN.JABOTICABAL" & "341" & FillSpace("ITAU", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq341, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "T341" & Format(nNumParc, "00") & Format(nSeq341, "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 409 Then 'UNIBANCO
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ409'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq409 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("20071987118022402", 20) & "PREF.MUN.JABOTICABAL" & "409" & FillSpace("UNIBANCO", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq409, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "T409" & Format(nNumParc, "00") & Format(nSeq409, "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 399 Then 'HSBC
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ399'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq399 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("00005012210", 20) & "PREF.MUN.JABOTICABAL" & "399" & FillSpace("HSBC", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq399, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "H399" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 33 Then 'BANESPA
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ033'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq033 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace("00330023008000007463", 20) & "PREF.MUN.JABOTICABAL" & "033" & FillSpace("SANTANDER BANESPA", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq033, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "B033" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 1 Then 'BB
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ001'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq001 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace(" ", 20) & "PREF.MUN.JABOTICABAL" & "001" & FillSpace("BANCO DO BRASIL", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq001, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "B001" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
ElseIf Val(Left$(cmbBanco.Text, 3)) = 104 Then 'CAIXA FED
   Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQ104'"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   nSeq104 = RdoAux!VALPARAM
   RegA = "A1" & FillSpace(" ", 20) & "PREF.MUN.JABOTICABAL" & "104" & FillSpace("CAIXA ECON FEDERAL", 20) & Format(Year(Now), "0000") & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(nSeq104, "000000") & "04" & "DEBITO AUTOMATICO" & FillSpace(" ", 52)
   sNomeArq = "C104" & Format(Day(Now), "00") & Format(Month(Now), "00") & ".TXT"
End If

'Sql = "SELECT valorparcela From expediente WHERE (codlancamento = 1) AND (anoexped = " & Year(Now) & ")"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'nValorTxExp = FormatNumber(RdoAux!VALORPARCELA, 2)
nValorTxExp = 0
nContaLinha = 0
nSoma = 0
Ocupado
'sPathArqDA = "c:\tmp"
Open sPathArqDA & "\" & sNomeArq For Output As #1
With grdOpt
   .SetFocus
   
    Print #1, RegA
    For x = 1 To .Rows - 1
        nCodPref = .TextMatrix(x, 7)
        If x Mod 5 = 0 Then
           CallPb CLng(x), CLng(.Rows - 1)
        End If
'        nCodReduz = .TextMatrix(x, 0)
'        If nCodReduz = 1399 Then MsgBox "teste"
        'Sql = "SELECT STATUSLANC FROM DEBITOPARCELA WHERE CODREDUZIDO = " & .TextMatrix(x, 0) & " AND year(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND MONTH(DATAVENCIMENTO)=" & Val(cmbMes.Text) & " AND SEQLANCAMENTO=0"
        Sql = "SELECT STATUSLANC FROM DEBITOPARCELA WHERE CODREDUZIDO = " & nCodPref & " AND NUMPARCELA>0 AND  year(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND MONTH(DATAVENCIMENTO)=" & Val(cmbMes.Text) & " AND SEQLANCAMENTO=0"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If !statuslanc = 1 Then
                    GoTo proximo
                End If
            Else
                GoTo proximo
            End If
           .Close
        End With
        
        Sql = "SELECT SUM(DEBITOTRIBUTO.VALORTRIBUTO) AS TOTAL, DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.DATAVENCIMENTO FROM DEBITOPARCELA INNER JOIN DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO AND "
        Sql = Sql & "DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND "
        Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOTRIBUTO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO "
        Sql = Sql & " where DEBITOPARCELA.CODREDUZIDO = " & nCodPref & " and DEBITOPARCELA.CODLANCAMENTO=1 AND DEBITOPARCELA.NUMPARCELA>0 AND YEAR(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND MONTH(DATAVENCIMENTO)=" & Val(cmbMes.Text) & " AND DAY(DATAVENCIMENTO)=" & Val(cmbDia.Text) & " AND DEBITOPARCELA.STATUSLANC=3 AND DEBITOTRIBUTO.CODTRIBUTO<>3 GROUP BY DEBITOPARCELA.CODREDUZIDO,DEBITOPARCELA.DATAVENCIMENTO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               grdOpt.TextMatrix(x, 5) = Format(!DataVencimento, "dd/mm/yyyy")
               grdOpt.TextMatrix(x, 6) = FormatNumber(!Total + nValorTxExp, 2)
               nSoma = nSoma + !Total + nValorTxExp
                                                      
               With xImovel
                    '.CarregaImovel Val(grdOpt.TextMatrix(x, 0))
                    .CarregaImovel nCodPref
'                    sInsc = Format$(RemovePonto(Left$(.Inscricao, 18)), "000000000000000")
                    
                    sInsc = Format$(RemovePonto(Left$(.CodigoImovel, 18)), "000000000000000")
               End With
                                                      
               'ax = "E" & FillSpace(grdOpt.TextMatrix(x, 0), 25) & Format(grdOpt.TextMatrix(x, 2), "0000") & FillSpace(grdOpt.TextMatrix(x, 3), 14)
               ax = "E" & FillSpace(sInsc, 25) & Format(grdOpt.TextMatrix(x, 2), "0000") & FillSpace(grdOpt.TextMatrix(x, 3), 14)
               ax = ax & Right$(grdOpt.TextMatrix(x, 5), 4) & Mid$(grdOpt.TextMatrix(x, 5), 4, 2) & Left$(grdOpt.TextMatrix(x, 5), 2)
               ax = ax & Format(RetornaNumero(grdOpt.TextMatrix(x, 6)), "000000000000000") & "03" & FillSpace(grdOpt.TextMatrix(x, 0), 60) & FillSpace(" ", 20) & "0"
               Print #1, ax
               nContaLinha = nContaLinha + 1
            End If
           .Close
        End With
proximo:
    Next
End With
sSoma = CStr(nSoma)
sSoma = FormatNumber(sSoma, 2)
If Left(Right(sSoma, 2), 1) = "," Then
    nSoma = RetornaNumero(sSoma) * 10
Else
    nSoma = RetornaNumero(sSoma)
End If

nContaLinha = nContaLinha + 2

RegZ = "Z" & Format(nContaLinha, "000000") & Format(nSoma, "00000000000000000") & FillSpace(" ", 126)
Print #1, RegZ
Close #1

If Val(Left$(cmbBanco.Text, 3)) = 237 Then 'BRADESCO
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ237'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 341 Then 'ITAU
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ341'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 409 Then 'UNIBANCO
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ409'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 399 Then 'HSBC
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ399'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 33 Then 'BANESPA
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ033'"
    cn.Execute Sql, rdExecDirect
ElseIf Val(Left$(cmbBanco.Text, 3)) = 1 Then 'BB
    Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM+1 WHERE NOMEPARAM='SEQ033'"
    cn.Execute Sql, rdExecDirect
End If

Liberado
MsgBox "FIM", vbInformation, "Atenção"
End Sub


Private Sub Command1_Click()
Dim nCodReduz As Long, sDataVencimento As String

sDataVencimento = "16/03/2005"
With grdOpt
    For x = 1 To .Rows - 1
       nCodReduz = .TextMatrix(x, 0)
       .TextMatrix(x, 5) = sDataVencimento
       Sql = "UPDATE DEBITOPARCELA SET DATAVENCIMENTO='" & Format(CDate(sDataVencimento), "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & nCodReduz
       Sql = Sql & " AND ANOEXERCICIO=2005 AND CODLANCAMENTO=1 AND NUMPARCELA=3"
       cn.Execute Sql, rdExecDirect
    Next
End With
MsgBox "FIM"
End Sub

Private Sub Command2_Click()
Dim sData As String, xId As Long, nNumRec As Long

xId = 1
nNumRec = 1
Sql = "SELECT * From DEBITOPARCELA Where CODREDUZIDO > 500000 AND  (AnoExercicio > 2004) And (Year(DATAVENCIMENTO) = 2000)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nNumRec = .RowCount
    Do Until .EOF
        If x Mod 100 = 0 Then
           CallPb xId, nNumRec
        End If
        
        sData = Day(!DataVencimento) & "/" & Month(!DataVencimento) & "/" & !AnoExercicio
        Sql = "UPDATE DEBITOPARCELA SET DATAVENCIMENTO='" & Format(sData, "mm/dd/yyyy") & "' WHERE "
        Sql = Sql & "CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
        Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO
        cn.Execute Sql, rdExecDirect
        
        xId = xId + 1
       .MoveNext
    Loop
End With
 
 
End Sub

Private Sub Form_Load()
Ocupado
Set xImovel = New clsImovel
Centraliza Me
cmbBanco.ListIndex = 0
Liberado
cmbAno.Clear
cmbAno.AddItem Year(Now)
cmbMes.Text = Month(Now)
cmbAno.Text = Year(Now)
cmbDia.Text = Day(Now)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub grdOpt_Click()
If grdOpt.Rows = 1 Then Exit Sub
If grdOpt.MouseRow > 0 Then Exit Sub

grdOpt.Row = 1
grdOpt.RowSel = grdOpt.Rows - 1
grdOpt.col = grdOpt.MouseCol
grdOpt.Sort = flexSortStringAscending

End Sub

Private Sub grdOpt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'If grdOpt.Row = 1 Then
'   CarregaLista Val(Left$(cmbBanco.text, 3)), grdOpt.Col
'End If

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPF.Caption = FormatNumber(Pb.value, 2)

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub


Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Format(sTexto, String(nTamanho, "0"))

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function


