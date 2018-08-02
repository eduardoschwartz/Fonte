VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmListaSN 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de pagamento dos optantes do Simples Nacional"
   ClientHeight    =   1680
   ClientLeft      =   5610
   ClientTop       =   4965
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6540
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   180
      TabIndex        =   10
      Top             =   1215
      Width           =   4740
      _ExtentX        =   8361
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
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmDevedorSimples.frx":0000
      Left            =   1665
      List            =   "frmDevedorSimples.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   585
      Width           =   3210
   End
   Begin VB.ComboBox cmbAte 
      Height          =   315
      ItemData        =   "frmDevedorSimples.frx":003D
      Left            =   5130
      List            =   "frmDevedorSimples.frx":0065
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   1140
   End
   Begin VB.ComboBox cmbDe 
      Height          =   315
      ItemData        =   "frmDevedorSimples.frx":00A5
      Left            =   3330
      List            =   "frmDevedorSimples.frx":00CD
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   1140
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmDevedorSimples.frx":010D
      Left            =   1665
      List            =   "frmDevedorSimples.frx":010F
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1140
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5190
      TabIndex        =   5
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
      MICON           =   "frmDevedorSimples.frx":0111
      PICN            =   "frmDevedorSimples.frx":012D
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5190
      TabIndex        =   4
      ToolTipText     =   "Cancelar Edição"
      Top             =   810
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDevedorSimples.frx":019B
      PICN            =   "frmDevedorSimples.frx":01B7
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
      Caption         =   "Tipo de Seleção...:"
      Height          =   225
      Index           =   3
      Left            =   225
      TabIndex        =   9
      Top             =   645
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Até..:"
      Height          =   225
      Index           =   2
      Left            =   4680
      TabIndex        =   8
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "De..:"
      Height          =   225
      Index           =   1
      Left            =   2970
      TabIndex        =   7
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de apuração..:"
      Height          =   225
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   240
      Width           =   1500
   End
End
Attribute VB_Name = "frmListaSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Type SIMPLES
    nCodReduz As Long
    sRazao As String
    sCNPJ As String
    sData As String
    nMes1 As Double
    nMes2 As Double
    nMes3 As Double
    nMes4 As Double
    nMes5 As Double
    nMes6 As Double
    nMes7 As Double
    nMes8 As Double
    nMes9 As Double
    nMes10 As Double
    nMes11 As Double
    nMes12 As Double
    sAtividade As String
End Type

Private Type REFS
    nAnoVT As Integer
    nMesVT As Integer
    nAnoRL As Integer
    nMesRL As Integer
End Type

Private Sub Form_Load()
Dim x As Integer

For x = 2004 To Year(Now)
    cmbAno.AddItem CStr(x)
Next

Centraliza Me
cmbAno.Text = Year(Now)
cmbDe.ListIndex = 0
cmbAte.ListIndex = 11
cmbTipo.ListIndex = 0
End Sub

Private Sub cmdGerar_Click()
Dim Sql As String, Rdoaux As rdoResultset, nCodReduz As Long, dDataVencto As Date, RdoAux2 As rdoResultset, t As Integer, nMesRef As Integer, nAnoRef As Integer
Dim nNumRec As Long, xId As Long, sMes As String, nMesDe As Integer, nMesAte As Integer, nValor As Double
Dim aSimples() As SIMPLES, x As Integer, bAchou As Boolean, aMatriz(1 To 12) As Double, Y As Integer, aRef() As REFS, sAtividade As String

ReDim aSimples(0): ReDim aRef(0)
nMesDe = cmbDe.ListIndex + 1: nMesAte = cmbAte.ListIndex + 1

If nMesDe > nMesAte Then
    MsgBox "Período inválido.", vbCritical, "Erro"
    Exit Sub
End If

Sql = "SELECT * FROM SIMPLESREF ORDER BY ANOVT,MESVT,ANORL,MESRL"
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With Rdoaux
    Do Until .EOF
        ReDim Preserve aRef(UBound(aRef) + 1)
        aRef(UBound(aRef)).nAnoVT = !ANOVT
        aRef(UBound(aRef)).nMesVT = !MESVT
        aRef(UBound(aRef)).nAnoRL = !ANORL
        aRef(UBound(aRef)).nMesRL = !MESRL
       .MoveNext
    Loop
   .Close
End With

Sql = "DELETE FROM LISTAPAGTOSN"
cn.Execute Sql, rdExecDirect

Sql = "SELECT debitoparcela.datavencimento, SUM(debitotributo.valortributo) AS Total, mobiliario.codigomob, mobiliario.razaosocial, mobiliario.cnpj, "
Sql = Sql & "mobiliario.dataabertura FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
Sql = Sql & "mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob WHERE   (debitotributo.codtributo <> 3) AND (debitoparcela.codlancamento = 5) AND (mobiliario.simples = 1) AND "
Sql = Sql & "(debitoparcela.statuslanc = 1 OR debitoparcela.statuslanc = 2 OR debitoparcela.statuslanc = 3 OR debitoparcela.statuslanc = 7) "
Sql = Sql & "GROUP BY debitoparcela.codreduzido, debitoparcela.datavencimento, mobiliario.codigomob, mobiliario.razaosocial, mobiliario.cnpj, "
Sql = Sql & "mobiliario.dataabertura HAVING (YEAR(debitoparcela.datavencimento) BETWEEN  " & Val(cmbAno.Text) & "  AND " & Val(cmbAno.Text) + 1 & " ) ORDER BY debitoparcela.codreduzido, debitoparcela.datavencimento "
Set Rdoaux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With Rdoaux
    nNumRec = .RowCount: xId = 1
    Do Until .EOF
        If xId Mod 20 = 0 Then
            CallPb xId, nNumRec
        End If
        nCodReduz = !CODIGOMOB
        nValor = !Total
        
        sAtividade = ""
        Sql = "SELECT CODATIVIDADE FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sAtividade = sAtividade & CStr(!codatividade) & ","
               .MoveNext
            Loop
           .Close
        End With
        If sAtividade <> "" Then
            sAtividade = Left(sAtividade, Len(sAtividade) - 1)
        End If
        
        bAchou = False
        For x = 0 To UBound(aSimples)
            If aSimples(x).nCodReduz = nCodReduz Then
                bAchou = True: Exit For
            End If
        Next
                
        If Not bAchou Then
            ReDim Preserve aSimples(UBound(aSimples) + 1)
            aSimples(UBound(aSimples)).nCodReduz = nCodReduz
            aSimples(UBound(aSimples)).sRazao = !RAZAOSOCIAL
            aSimples(UBound(aSimples)).sCNPJ = !Cnpj
            aSimples(UBound(aSimples)).sData = Format(!DATAABERTURA, "dd/mm/yyyy")
            aSimples(UBound(aSimples)).sAtividade = sAtividade
            nMesRef = 0: nAnoRef = 0
            For t = 1 To UBound(aRef)
                If Year(!DataVencimento) = aRef(t).nAnoVT And Month(!DataVencimento) = aRef(t).nMesVT Then
                    nAnoRef = aRef(t).nAnoRL: nMesRef = aRef(t).nMesRL
                    Exit For
                End If
            Next
            If nAnoRef <> Val(cmbAno.Text) Then GoTo PROXIMO2
            Select Case nMesRef
                Case 0
                    GoTo PROXIMO2
                Case 1
                    aSimples(UBound(aSimples)).nMes1 = nValor
                Case 2
                    aSimples(UBound(aSimples)).nMes2 = nValor
                Case 3
                    aSimples(UBound(aSimples)).nMes3 = nValor
                Case 4
                    aSimples(UBound(aSimples)).nMes4 = nValor
                Case 5
                    aSimples(UBound(aSimples)).nMes5 = nValor
                Case 6
                    aSimples(UBound(aSimples)).nMes6 = nValor
                Case 7
                    aSimples(UBound(aSimples)).nMes7 = nValor
                Case 8
                    aSimples(UBound(aSimples)).nMes8 = nValor
                Case 9
                    aSimples(UBound(aSimples)).nMes9 = nValor
                Case 10
                    aSimples(UBound(aSimples)).nMes10 = nValor
                Case 11
                    aSimples(UBound(aSimples)).nMes11 = nValor
                Case 12
                    aSimples(UBound(aSimples)).nMes12 = nValor
            End Select
            
        Else
            
            nMesRef = 0: nAnoRef = 0
            For t = 1 To UBound(aRef)
                If Year(!DataVencimento) = aRef(t).nAnoVT And Month(!DataVencimento) = aRef(t).nMesVT Then
                    nAnoRef = aRef(t).nAnoRL: nMesRef = aRef(t).nMesRL
                    Exit For
                End If
            Next
            If nAnoRef <> Val(cmbAno.Text) Then GoTo PROXIMO2
            Select Case nMesRef
                Case 0
                    GoTo PROXIMO2
                Case 1
                    aSimples(x).nMes1 = nValor
                Case 2
                    aSimples(x).nMes2 = nValor
                Case 3
                    aSimples(x).nMes3 = nValor
                Case 4
                    aSimples(x).nMes4 = nValor
                Case 5
                    aSimples(x).nMes5 = nValor
                Case 6
                    aSimples(x).nMes6 = nValor
                Case 7
                    aSimples(x).nMes7 = nValor
                Case 8
                    aSimples(x).nMes8 = nValor
                Case 9
                    aSimples(x).nMes9 = nValor
                Case 10
                    aSimples(x).nMes10 = nValor
                Case 11
                    aSimples(x).nMes11 = nValor
                Case 12
                    aSimples(x).nMes12 = nValor
            End Select
            
            
        End If
PROXIMO2:
        xId = xId + 1
       .MoveNext
    Loop
   .Close
End With

PBar.Value = 0: nNumRec = UBound(aSimples)
For x = 1 To UBound(aSimples)
    If x Mod 20 = 0 Then
        CallPb CLng(x), nNumRec
    End If
    With aSimples(x)
        aMatriz(1) = .nMes1: aMatriz(2) = .nMes2: aMatriz(3) = .nMes3: aMatriz(4) = .nMes4: aMatriz(5) = .nMes5: aMatriz(6) = .nMes6
        aMatriz(7) = .nMes7: aMatriz(8) = .nMes8: aMatriz(9) = .nMes9: aMatriz(10) = .nMes10: aMatriz(11) = .nMes11: aMatriz(12) = .nMes12
    End With
    
    If cmbTipo.ListIndex = 0 Then 'NENHUMA PARCELA PAGA NO PERIODO
        For Y = nMesDe To nMesAte
            If aMatriz(Y) > 0 Then
                GoTo PROXIMO
            End If
        Next
    ElseIf cmbTipo.ListIndex = 1 Then 'ALGUMA PARCELA NÃO PAGA NO PERIODO
        bAchou = False
        For Y = nMesDe To nMesAte
            If aMatriz(Y) = 0 Then
                bAchou = True: Exit For
            End If
        Next
        If Not bAchou Then GoTo PROXIMO
    End If
    
    Sql = "INSERT LISTAPAGTOSN (CODIGO,RAZAO,CNPJ,DATAABERTURA,JAN,FEV,MAR,ABR,MAI,JUN,JUL,AGO,SETB,OUT,NOV,DEZ,ATIVIDADE) VALUES (" & aSimples(x).nCodReduz & ",'" & Mask(aSimples(x).sRazao) & "','" & aSimples(x).sCNPJ & "','" & Format(aSimples(x).sData, "mm/dd/yyyy") & "',"
    Sql = Sql & Virg2Ponto(CStr(aSimples(x).nMes1)) & "," & Virg2Ponto(CStr(aSimples(x).nMes2)) & "," & Virg2Ponto(CStr(aSimples(x).nMes3)) & "," & Virg2Ponto(CStr(aSimples(x).nMes4)) & ","
    Sql = Sql & Virg2Ponto(CStr(aSimples(x).nMes5)) & "," & Virg2Ponto(CStr(aSimples(x).nMes6)) & "," & Virg2Ponto(CStr(aSimples(x).nMes7)) & "," & Virg2Ponto(CStr(aSimples(x).nMes8)) & ","
    Sql = Sql & Virg2Ponto(CStr(aSimples(x).nMes9)) & "," & Virg2Ponto(CStr(aSimples(x).nMes10)) & "," & Virg2Ponto(CStr(aSimples(x).nMes11)) & "," & Virg2Ponto(CStr(aSimples(x).nMes12)) & ",'" & aSimples(x).sAtividade & "')"
    cn.Execute Sql, rdExecDirect
PROXIMO:
Next

PBar.Value = 100

frmReport.ShowReport "LISTAPAGTOSN", frmMdi.hwnd, Me.hwnd
Sql = "DELETE FROM LISTAPAGTOSN"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub


Private Sub CallPb(nVal As Long, nTot As Long)
If ((nVal * 100) / nTot) <= 100 Then
   PBar.Value = (nVal * 100) / nTot
Else
   PBar.Value = 100
End If
Me.Refresh
If cGetInputState() <> 0 Then DoEvents
End Sub

