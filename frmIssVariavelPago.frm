VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmIssVariavelPago 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Iss variável pago entre 2001 e 2006"
   ClientHeight    =   5040
   ClientLeft      =   3180
   ClientTop       =   3180
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   11115
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   240
      Left            =   135
      TabIndex        =   5
      Top             =   4635
      Width           =   5145
      _ExtentX        =   9075
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
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5340
      TabIndex        =   4
      ToolTipText     =   "Parar Execução"
      Top             =   4650
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Parar"
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
      MICON           =   "frmIssVariavelPago.frx":0000
      PICN            =   "frmIssVariavelPago.frx":001C
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
      Height          =   315
      Left            =   8190
      TabIndex        =   3
      ToolTipText     =   "Enviar para a impressora"
      Top             =   4650
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmIssVariavelPago.frx":0176
      PICN            =   "frmIssVariavelPago.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   9600
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   4650
      Width           =   1350
      _ExtentX        =   2381
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
      MICON           =   "frmIssVariavelPago.frx":02EC
      PICN            =   "frmIssVariavelPago.frx":0308
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
      Left            =   6780
      TabIndex        =   2
      ToolTipText     =   "Cálcula lancamentos pagos"
      Top             =   4650
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Calcular"
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
      MICON           =   "frmIssVariavelPago.frx":0376
      PICN            =   "frmIssVariavelPago.frx":0392
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   4515
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7964
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   12582912
      ForeColorFixed  =   16777215
      BackColorSel    =   192
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmIssVariavelPago.frx":0431
   End
End
Attribute VB_Name = "frmIssVariavelPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, bStop As Boolean

Private Sub cmdCalculo_Click()
Dim x As Integer, nCodReduz As Long, nAno As Integer, nValor As Double, nSoma As Double
Dim nTot As Long, nPos As Long

Ocupado
grdMain.Rows = 1: Pb.Value = 0: bStop = False

Sql = "SELECT debitoparcela.codreduzido, mobiliario.razaosocial FROM debitoparcela INNER JOIN debitotributo ON "
Sql = Sql & "debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN "
Sql = Sql & "mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob WHERE (debitoparcela.anoexercicio BETWEEN 2001 AND 2006) AND "
Sql = Sql & "(debitoparcela.codlancamento = 5) AND (debitoparcela.statuslanc = 1 OR  debitoparcela.statuslanc = 2) AND "
Sql = Sql & "(debitotributo.codtributo = 13) AND (mobiliario.dataencerramento IS NULL) GROUP BY debitoparcela.codreduzido, "
Sql = Sql & "mobiliario.razaosocial HAVING (debitoparcela.codreduzido BETWEEN 100000 AND 300000) AND (SUM(debitotributo.valortributo) > 0) "
Sql = Sql & "ORDER BY SUM(debitotributo.valortributo) DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddItem Format(!CODREDUZIDO, "000000") & "-" & !razaosocial
       .MoveNext
    Loop
   .Close
End With

With grdMain
    For x = 1 To .Rows - 1
        If x > 12 Then .TopRow = x - 10
        nSoma = 0: nTot = .Rows - 1
        nPos = nPos + 1
        CallPb nPos, CLng(nTot)
        For nAno = 2001 To 2006
            nCodReduz = Val(Left$(.TextMatrix(x, 0), 6))
            Sql = "SELECT debitoparcela.codreduzido, mobiliario.razaosocial, SUM(debitotributo.valortributo) AS valortotal, debitoparcela.anoexercicio "
            Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND "
            Sql = Sql & "debitoparcela.anoexercicio = debitotributo.anoexercicio AND debitoparcela.codlancamento = debitotributo.codlancamento AND "
            Sql = Sql & "debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND "
            Sql = Sql & "debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob "
            Sql = Sql & "WHERE (debitoparcela.codlancamento = 5) AND (debitoparcela.statuslanc = 1 OR  debitoparcela.statuslanc = 2) AND "
            Sql = Sql & "(debitotributo.codtributo = 13) AND (mobiliario.dataencerramento IS NULL) GROUP BY debitoparcela.codreduzido, mobiliario.razaosocial, "
            Sql = Sql & "debitoparcela.anoexercicio Having debitoparcela.CODREDUZIDO = " & nCodReduz & " And debitoparcela.AnoExercicio = " & nAno & " And Sum(debitotributo.valortributo) > 0"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount = 0 Then
                    nValor = 0
                Else
                    nValor = !ValorTotal
                End If
               .Close
            End With
            Select Case nAno
                Case 2001
                    .TextMatrix(x, 1) = FormatNumber(nValor, 2)
                Case 2002
                    .TextMatrix(x, 2) = FormatNumber(nValor, 2)
                Case 2003
                    .TextMatrix(x, 3) = FormatNumber(nValor, 2)
                Case 2004
                    .TextMatrix(x, 4) = FormatNumber(nValor, 2)
                Case 2005
                    .TextMatrix(x, 5) = FormatNumber(nValor, 2)
                Case 2006
                    .TextMatrix(x, 6) = FormatNumber(nValor, 2)
            End Select
            nSoma = nSoma + nValor
        Next
        .TextMatrix(x, 7) = FormatNumber(nSoma, 2)
        If cGetInputState() <> 0 Then DoEvents
        If bStop Then
            MsgBox "Terminado pelo usuário.", vbExclamation, "Atenção"
            Exit For
        End If
    Next
    .TopRow = 1
End With

Liberado
End Sub

Private Sub cmdCancel_Click()
bStop = True
End Sub

Private Sub cmdPrint_Click()
Dim x As Integer, nCodReduz As Long, sRazao As String

Ocupado
If grdMain.Rows = 1 Then
    MsgBox "Cálculo não gerado.", vbExclamation, "Atenção"
    Exit Sub
End If

Sql = "DELETE FROM ISSVARIAVELTMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

With grdMain
    For x = 1 To .Rows - 1
        nCodReduz = Val(Left$(.TextMatrix(x, 0), 6))
        sRazao = Mid$(.TextMatrix(x, 0), 8, Len(.TextMatrix(x, 0)) - 7)
        If .TextMatrix(x, 1) = "" Then Exit For
        Sql = "INSERT ISSVARIAVELTMP(COMPUTER,CODREDUZIDO,RAZAOSOCIAL,V1,V2,V3,V4,V5,V6,TOTAL) VALUES('" & NomeDoUsuario & "',"
        Sql = Sql & nCodReduz & ",'" & Left$(sRazao, 50) & "'," & Virg2Ponto(RemovePonto(.TextMatrix(x, 1))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 2))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(.TextMatrix(x, 3))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 4))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 5))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(.TextMatrix(x, 6))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 7))) & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

Liberado

frmReport.ShowReport "ISSVARIAVELPAGO", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ISSVARIAVELTMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect


End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.Value = (nPosF * 100) / nTotal
Else
   Pb.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

