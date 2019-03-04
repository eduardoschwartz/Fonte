VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmDividaAtiva 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encerramento do Livro de Divida Ativa"
   ClientHeight    =   3525
   ClientLeft      =   2925
   ClientTop       =   2700
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6570
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmDividaAtiva.frx":0000
      Left            =   2025
      List            =   "frmDividaAtiva.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   2265
   End
   Begin VB.ComboBox cmbAno 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmDividaAtiva.frx":0004
      Left            =   1980
      List            =   "frmDividaAtiva.frx":0065
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   915
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   4995
      TabIndex        =   3
      Top             =   1230
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmDividaAtiva.frx":0123
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
   Begin prjChameleon.chameleonButton cmdRetorna 
      Height          =   345
      Left            =   2850
      TabIndex        =   11
      ToolTipText     =   "Encerramento do Livro"
      Top             =   2670
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Encerrar"
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
      MICON           =   "frmDividaAtiva.frx":013F
      PICN            =   "frmDividaAtiva.frx":015B
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
      Height          =   345
      Left            =   5175
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   2670
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
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
      MICON           =   "frmDividaAtiva.frx":02B5
      PICN            =   "frmDividaAtiva.frx":02D1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   345
      Left            =   1680
      TabIndex        =   13
      ToolTipText     =   "Impressão do Livro de Divida Ativa"
      Top             =   2670
      Visible         =   0   'False
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
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
      MCOL            =   65280
      MPTR            =   1
      MICON           =   "frmDividaAtiva.frx":033F
      PICN            =   "frmDividaAtiva.frx":035B
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
      Left            =   3135
      TabIndex        =   14
      Top             =   2130
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Contribuintes à Analisar..:"
      Height          =   225
      Index           =   2
      Left            =   90
      TabIndex        =   22
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1980
      TabIndex        =   21
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2805
      TabIndex        =   20
      Top             =   2130
      Width           =   270
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   5595
      TabIndex        =   19
      Top             =   2145
      Width           =   480
   End
   Begin VB.Label lblFalta 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   5025
      TabIndex        =   18
      Top             =   1800
      Width           =   810
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Contribuintes Restantes.:"
      Height          =   225
      Index           =   3
      Left            =   3135
      TabIndex        =   17
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label lblPag 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   1980
      TabIndex        =   16
      Top             =   2130
      Width           =   705
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Número da Página.........:"
      Height          =   225
      Index           =   4
      Left            =   105
      TabIndex        =   15
      Top             =   2130
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Livro................:"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   10
      Top             =   615
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano do Livro.................:"
      Height          =   210
      Left            =   150
      TabIndex        =   9
      Top             =   255
      Width           =   1755
   End
   Begin VB.Label lblNumero 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   1980
      TabIndex        =   8
      Top             =   945
      Width           =   4335
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Número do Livro...........:"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   7
      Top             =   945
      Width           =   1755
   End
   Begin VB.Label lblDataIni 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   1980
      TabIndex        =   6
      Top             =   1305
      Width           =   960
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Encerramento..:"
      Height          =   225
      Index           =   0
      Left            =   3150
      TabIndex        =   5
      Top             =   1290
      Width           =   1755
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Abertura..........:"
      Height          =   225
      Index           =   5
      Left            =   150
      TabIndex        =   4
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label lblLog 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Aguardando."
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   15
      TabIndex        =   0
      Top             =   3210
      Width           =   6315
   End
End
Attribute VB_Name = "frmDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, bEnc As Boolean
Dim nTotal As Integer

Private Sub cmbAno_Click()
If cmbAno.ListIndex = -1 Then Exit Sub
Limpa
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
cmbTipo_Click
End Sub

Private Sub cmbTipo_Click()
Dim nLast As Integer, sNum As String
If cmbTipo.ListIndex = -1 Then Exit Sub
Limpa
Sql = "SELECT NUMERO,DATAABERTURA,DATAENCERRAMENTO FROM LIVRO WHERE ANO=" & Val(cmbAno.Text)
Sql = Sql & " AND CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
       If IsNull(!dataencerramento) Then
          bEnc = False
       Else
          bEnc = True
       End If
       Sql = "SELECT NUMEROOLD FROM GRADELIVRO WHERE ANO=" & Val(cmbAno.Text) & " AND "
       Sql = Sql & "CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " ORDER BY NUMEROOLD"
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            If .RowCount > 0 Then
                Do Until .EOF
                    sNum = sNum & CStr(!NUMEROOLD) & ", "
                   .MoveNext
                Loop
                sNum = Chomp(sNum, chomp_righT, 1)
                lblNumero.Caption = sNum
            Else
                lblNumero.Caption = RdoAux!Numero
            End If
           .Close
       End With
       If Not IsNull(!DATAABERTURA) Then lblDataIni.Caption = Format(!DATAABERTURA, "dd/mm/yyyy")
       If Not IsNull(!dataencerramento) Then
          mskDataFim.Text = Format(!dataencerramento, "dd/mm/yyyy")
          mskDataFim.Locked = True
          mskDataFim.BackColor = Kde
       Else
          mskDataFim.Locked = False
          mskDataFim.BackColor = Branco
       End If
    Else
       bEnc = False
       Sql = "SELECT MAX(NUMERO) AS MAXIMO FROM LIVRO "
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       nLast = RdoAux2!maximo
       RdoAux2.Close
       lblNumero.Caption = nLast + 1
    End If
   .Close
End With

End Sub

Private Sub cmdBaixa_Click()

If Not bEnc Then
    MsgBox "Este livro ainda não foi encerrado.", vbExclamation, "Atenção"
    Exit Sub
End If

End Sub

Private Sub cmdRetorna_Click()

If bEnc Then
    'MsgBox "Este livro já foi encerrado.", vbExclamation, "Atenção"
    'Exit Sub
Else
    If lblDataIni.Caption = "" Then
        MsgBox "Este livro não esta aberto.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If Not IsDate(mskDataFim.Text) Then
        MsgBox "Data de Encerramento inválida", vbExclamation, "Atenção"
        Exit Sub
    End If
    If Year(mskDataFim.Text) < Val(cmbAno.Text) Then
        MsgBox "Data de encerramento inválida", vbExclamation, "atenção"
        Exit Sub
    End If
End If

If MsgBox("Voce deseja encerrar o Livro de " & cmbTipo.Text & " nº " & lblNumero.Caption & " de " & cmbAno.Text & " ?", vbQuestion + vbYesNo, "CONFIRMAÇÃO !!!") = vbNo Then Exit Sub
Ocupado
CloseBook cmbTipo.ItemData(cmbTipo.ListIndex)
Liberado
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me

'cmbAno.Text = Year(Now)
Sql = "SELECT CODTIPO, DESCTIPO FROM TIPOLIVRO "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !DESCTIPO
        cmbTipo.ItemData(cmbTipo.NewIndex) = !CodTipo
       .MoveNext
    Loop
   .Close
End With
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0

End Sub

Private Sub Limpa()
lblNumero.Caption = ""
LimpaMascara mskDataFim
lblDataIni.Caption = ""
End Sub

Private Sub CloseBook(nTipo As Integer)
Dim sTributosDA As String, sLancamentoDA As String, xId As Long, sTypeBook As String, sLanc As String
Dim nPos As Integer, nPagina As Integer
Pb.value = 0

Sql = "select codlancamento from lancamento where tipolivro=" & nTipo & " order by codlancamento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sLanc = sLanc & !CodLancamento & ","
       .MoveNext
    Loop
   .Close
End With
sLanc = Chomp(sLanc, chomp_righT, 1)

sLancamentoDA = "CODLANCAMENTO in (" & sLanc & ")"

lblLog.Caption = "Analisando Tributos...."
lblLog.Refresh
'Sql = "SELECT CODTRIBUTO FROM TRIBUTO WHERE DA=1 ORDER BY CODTRIBUTO"
Sql = "SELECT CODTRIBUTO FROM TRIBUTO WHERE livro=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " ORDER BY CODTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sTributosDA = sTributosDA & !CodTributo & ","
       .MoveNext
    Loop
End With
sTributosDA = Chomp(sTributosDA, chomp_righT, 1)

lblLog.Caption = "Analisando Livro...."
lblLog.Refresh
Sql = "SELECT MAX(PAGINALIVRO) AS MAXIMO FROM DEBITOPARCELA WHERE numerolivro=" & Val(lblNumero.Caption) & " and anoexercicio=" & cmbAno.Text & " and (statuslanc=3)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nPagina = 1
    Else
        nPagina = !maximo + 1
    End If
    lblPag.Caption = nPagina
End With

lblLog.Caption = "Analisando Lançamentos...."
lblLog.Refresh


Sql = "SELECT MAX(numcertidao) AS MAXIMO FROM DEBITOPARCELA WHERE numerolivro=" & Val(lblNumero.Caption) & " and anoexercicio=" & Val(cmbAno.Text) & " and (statuslanc=3)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If Not IsNull(RdoAux!maximo) Then
    xId = RdoAux!maximo + 1
Else
    xId = 1
End If

'Sql = "SELECT DISTINCT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO FROM vwDIVIDAATIVA WHERE " & sTypeBook & " AND YEAR(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND (" & sLancamentoDA & ") AND NUMPARCELA>0 AND (statuslanc=3) "
Sql = "SELECT DISTINCT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO FROM vwDIVIDAATIVA WHERE YEAR(DATAVENCIMENTO)=" & Val(cmbAno.Text) & " AND (" & sLancamentoDA & ") AND NUMPARCELA>0 AND (statuslanc=3) "
Sql = Sql & " AND DATAINSCRICAO IS NULL AND CODTRIBUTO IN (SELECT CODTRIBUTO FROM TRIBUTO WHERE DA=1) ORDER BY CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTotal = .RowCount
    lblTotal.Caption = nTotal
    lblFalta.Caption = nTotal
    lblLog.Caption = "Iniciando Atualização...."
    lblLog.Refresh
    Do Until .EOF
        'If !CODREDUZIDO = 102755 Then MsgBox "TESTE"
        If xId Mod 100 = 0 Then
            CallPb xId, CLng(lblTotal.Caption)
        End If
        Sql = "UPDATE DEBITOPARCELA SET NUMEROLIVRO=" & Val(lblNumero.Caption) & " ,PAGINALIVRO=" & nPagina
        Sql = Sql & " ,DATAINSCRICAO='" & Format(mskDataFim.Text, "mm/dd/yyyy") & "',NUMCERTIDAO=" & xId & " WHERE CODREDUZIDO=" & !CODREDUZIDO
        Sql = Sql & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND (STATUSLANC=3) "
        Sql = Sql & " AND DATAINSCRICAO IS NULL"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
'        If nPos > 1 Then
'            nPos = 1
            nPagina = nPagina + 1
            lblPag.Caption = nPagina
        'End If
        nTotal = nTotal - 1
        lblFalta.Caption = nTotal
        lblFalta.Refresh
proximo:
        xId = xId + 1
       .MoveNext
       DoEvents
    Loop
   .Close
End With

Sql = "UPDATE LIVRO SET DATAENCERRAMENTO='" & Format(mskDataFim.Text, "mm/dd/yyyy") & " ' WHERE NUMERO=" & lblNumero.Caption
cn.Execute Sql, rdExecDirect

lblPB.Caption = "100,00"
Pb.value = 100
lblLog.Caption = "Livro Encerrado com Sucesso...."
lblLog.Refresh

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

