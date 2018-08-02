VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDesdobroCarne 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desdobro de Carnê"
   ClientHeight    =   3510
   ClientLeft      =   4815
   ClientTop       =   4440
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   7905
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   3120
      Width           =   1935
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   6630
      TabIndex        =   7
      ToolTipText     =   "Sair da Tela"
      Top             =   3120
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
      MICON           =   "frmDesdobroCarne.frx":0000
      PICN            =   "frmDesdobroCarne.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   5400
      TabIndex        =   8
      ToolTipText     =   "Executar o desdobro de carnê"
      Top             =   3120
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Executar"
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
      MICON           =   "frmDesdobroCarne.frx":008A
      PICN            =   "frmDesdobroCarne.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CommandButton cmdDesconto 
      Caption         =   "&Desconto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   3
      Top             =   150
      Width           =   1155
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Carregar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5190
      TabIndex        =   2
      Top             =   150
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdAno 
      Height          =   2475
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   4366
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "^Ano            |>Qtde Parc. |>Vl. Parcela      |>Vl. Total        |>Desconto     |>Vl. Parc.Atual   |>Vl. Parc.Novo   "
   End
   Begin VB.TextBox txtCodNovo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3930
      TabIndex        =   1
      Top             =   150
      Width           =   975
   End
   Begin VB.TextBox txtCodAnt 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "No Processo:"
      Height          =   225
      Index           =   2
      Left            =   150
      TabIndex        =   10
      Top             =   3150
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Código Novo:"
      Height          =   225
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Código Antigo:"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   1125
   End
End
Attribute VB_Name = "frmDesdobroCarne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, RdoAux3 As rdoResultset

Private Sub cmdDesconto_Click()
Dim z As Variant
With grdAno
    If .Rows = 1 Then Exit Sub
    If .Row > 0 Then
        z = InputBox("Digite o valor do desconto.", "Novo Valor")
        If z = "" Or Val(z) = 0 Or Not IsNumeric(z) Then
            .TextMatrix(.Row, 4) = "0,00"
            .TextMatrix(.Row, 5) = "0,00"
            .TextMatrix(.Row, 6) = "0,00"
        Else
            .TextMatrix(.Row, 4) = Format(CDbl(z), "#0.00")
            .TextMatrix(.Row, 5) = Format((.TextMatrix(.Row, 3) - z) / .TextMatrix(.Row, 1), "#0.00")
            .TextMatrix(.Row, 6) = Format(z / .TextMatrix(.Row, 1), "#0.00")
        End If
    End If
End With

End Sub

Private Sub cmdExec_Click()
If grdAno.Rows = 1 Then
    MsgBox "Selecione os débitos", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtCodNovo.Text) = 0 Then
    MsgBox "Digite o código novo.", vbCritical, "Atenção"
    Exit Sub
End If

If txtNumProc.Text = "" Then
    MsgBox "Digite o numero do processo.", vbCritical, "Atenção"
    Exit Sub
End If

Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE CODREDUZIDO=" & Val(txtCodNovo.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Imóvel de destino não cadastrado.", vbCritical, "Atenção"
        Exit Sub
    End If
   .Close
End With

Sql = "SELECT CODREDUZIDO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & Val(txtCodNovo.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Imóvel de destino possue débitos.", vbCritical, "Atenção"
        Exit Sub
    End If
   .Close
End With

If MsgBox("Deseja executar o desdobro de carnês?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
    ExecutaDesdobro
End If

End Sub

Private Sub ExecutaDesdobro()
Dim x As Integer, nCodAnt As Long, nCodNovo As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, dVencto As Date
Dim nValorParcAnt As Double, nValorParcNovo As Double, nCodTributo As Integer, sObsNovo As String, sObsAnt As String, sData As String

nCodAnt = Val(txtCodAnt.Text)
nCodNovo = Val(txtCodNovo.Text)

sObsAnt = "Conforme processo " & txtNumProc.Text & " ocorreu o desdobro de carne e os valores foram alterados como segue:" & vbCrLf
sObsNovo = "Conforme processo " & txtNumProc.Text & " ocorreu o desdobro de carne e os novos valores são como segue:" & vbCrLf

With grdAno
    For x = 1 To .Rows - 1
        nAno = .TextMatrix(x, 0)
        nValorParcAnt = .TextMatrix(x, 5)
        nValorParcNovo = .TextMatrix(x, 6)
        Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodAnt & " AND ANOEXERCICIO=" & nAno
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nLanc = !CodLancamento
                nSeq = !SeqLancamento
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                dVencto = Format(!DataVencimento, "dd/mm/yyyy")
               'GRAVA NA TABELA DEBITOPARCELA
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'                Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",3,'"
'                Sql = Sql & Format(dVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',1,'GTI')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
                Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",3,'"
                Sql = Sql & Format(dVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',1,236)"
                cn.Execute Sql, rdExecDirect
               'GRAVA NA TABELA DEBITO TRIBUTO
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",1," & Virg2Ponto(CStr(nValorParcNovo)) & ")"
                cn.Execute Sql, rdExecDirect
               'APAGA TRIBUTOS ANTIGOS
                Sql = "DELETE FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & nCodAnt & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
                Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
               'GRAVA NOVOS TRIBUTOS PARA O CÓDIGO ANTIGO
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodAnt & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",1," & Virg2Ponto(CStr(nValorParcAnt)) & ")"
                cn.Execute Sql, rdExecDirect
               .MoveNext
            Loop
           .Close
        End With
        sObsAnt = sObsAnt & "Ano: " & nAno & " de " & .TextMatrix(x, 2) & " para " & FormatNumber(nValorParcAnt, 2) & vbCrLf
        sObsNovo = sObsNovo & "Ano: " & nAno & " Valor: " & FormatNumber(nValorParcNovo, 2) & vbCrLf
    Next
End With

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & nCodAnt
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With
sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USUARIO,DATAOBS,OBS) VALUES(" & nCodAnt & "," & nSeq & ",'"
Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "','" & Mask(sObsAnt) & "')"
cn.Execute Sql, rdExecDirect

Sql = "SELECT MAX(SEQ) AS MAXIMO FROM DEBITOOBSERVACAO WHERE CODREDUZIDO=" & nCodNovo
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeq = 1
    Else
        nSeq = !maximo + 1
    End If
   .Close
End With
sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
Sql = "INSERT DEBITOOBSERVACAO(CODREDUZIDO,SEQ,USUARIO,DATAOBS,OBS) VALUES(" & nCodNovo & "," & nSeq & ",'"
Sql = Sql & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "','" & Mask(sObsNovo) & "')"
cn.Execute Sql, rdExecDirect

MsgBox "Desdobro efetuado com sucesso.", vbInformation, "Atenção"


End Sub

Private Sub cmdLoad_Click()
Dim nAno As Integer, x As Integer

grdAno.Rows = 1: nAno = 0
Sql = "SELECT anoexercicio, COUNT(anoexercicio) AS contador From debitoparcela Where CODREDUZIDO = " & Val(txtCodAnt.Text)
Sql = Sql & " And (statuslanc = 3) and numparcela>0  GROUP BY anoexercicio ORDER BY ANOEXERCICIO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdAno.AddItem !AnoExercicio & Chr(9) & !contador
       .MoveNext
    Loop
   .Close
End With

With grdAno
    For x = 1 To .Rows - 1
        Sql = "SELECT numparcela From debitoparcela Where CODREDUZIDO = " & Val(txtCodAnt.Text) & " AND ANOEXERCICIO=" & .TextMatrix(x, 0)
        Sql = Sql & " And (statuslanc = 3) and numparcela>0 "
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        Sql = "SELECT SUM(VALORTRIBUTO) AS VALORPARCELA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & Val(txtCodAnt.Text) & " AND ANOEXERCICIO="
        Sql = Sql & .TextMatrix(x, 0) & " AND NUMPARCELA=" & RdoAux2!NumParcela & " AND CODTRIBUTO<>3"
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        grdAno.TextMatrix(x, 2) = Format(RdoAux3!VALORPARCELA, "#0.00")
        grdAno.TextMatrix(x, 3) = Format(RdoAux3!VALORPARCELA * CDbl(.TextMatrix(x, 1)), "#0.00")
        grdAno.TextMatrix(x, 4) = "0,00"
        grdAno.TextMatrix(x, 5) = "0,00"
        grdAno.TextMatrix(x, 6) = "0,00"
        RdoAux2.Close: RdoAux3.Close
    Next
End With

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub
