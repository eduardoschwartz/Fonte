VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmProdutividadeSaldo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Produtividade - Saldo"
   ClientHeight    =   5700
   ClientLeft      =   4500
   ClientTop       =   2535
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   5115
   Begin prjChameleon.chameleonButton cmdCalc 
      Height          =   315
      Left            =   180
      TabIndex        =   4
      ToolTipText     =   "Editar Registro"
      Top             =   6345
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmProdutividadeSaldo.frx":0000
      PICN            =   "frmProdutividadeSaldo.frx":001C
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
      Left            =   3960
      TabIndex        =   6
      ToolTipText     =   "Imprimir extrato"
      Top             =   5310
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmProdutividadeSaldo.frx":036E
      PICN            =   "frmProdutividadeSaldo.frx":038A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtEdit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   12
      Top             =   6390
      Visible         =   0   'False
      Width           =   555
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   2850
      TabIndex        =   5
      ToolTipText     =   "Editar Registro"
      Top             =   6255
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmProdutividadeSaldo.frx":04E4
      PICN            =   "frmProdutividadeSaldo.frx":0500
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
      Height          =   4245
      Left            =   60
      TabIndex        =   3
      Top             =   990
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   7488
      _Version        =   393216
      Rows            =   14
      Cols            =   8
      RowHeightMin    =   300
      BackColorBkg    =   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      MergeCells      =   1
      Appearance      =   0
      FormatString    =   "MÊS/ANO           |^Pontos Mês |                  |^Negativos |^Utilizados  |^Saldo          |Mes  |Ano      "
   End
   Begin VB.ComboBox cmbFiscal 
      Height          =   315
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   4125
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmProdutividadeSaldo.frx":065A
      Left            =   840
      List            =   "frmProdutividadeSaldo.frx":0685
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   540
      Width           =   2175
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmProdutividadeSaldo.frx":06EE
      Left            =   3750
      List            =   "frmProdutividadeSaldo.frx":06F0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   540
      Width           =   1215
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   315
      Left            =   2835
      TabIndex        =   10
      ToolTipText     =   "Cancelar Edição"
      Top             =   6255
      Visible         =   0   'False
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmProdutividadeSaldo.frx":06F2
      PICN            =   "frmProdutividadeSaldo.frx":070E
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
      Left            =   4050
      TabIndex        =   11
      ToolTipText     =   "Gravar os Dados"
      Top             =   6255
      Visible         =   0   'False
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
      MICON           =   "frmProdutividadeSaldo.frx":0868
      PICN            =   "frmProdutividadeSaldo.frx":0884
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEndMounth 
      Height          =   315
      Left            =   180
      TabIndex        =   13
      ToolTipText     =   "Fim de Mês"
      Top             =   5310
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Fim de Mês"
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
      MICON           =   "frmProdutividadeSaldo.frx":0C29
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
      Caption         =   "Fiscal....:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mês......:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano....:"
      Height          =   225
      Index           =   2
      Left            =   3150
      TabIndex        =   7
      Top             =   630
      Width           =   615
   End
End
Attribute VB_Name = "frmProdutividadeSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ExtratoType
    nDia As Integer
    sDia As String
    sEvento As String
    nPontos As Integer
End Type

Private Type TabelaType
    nCodFiscal As Integer
    nMes As Integer
    nAno As Integer
    nPontos As Integer
    nNegativo As Integer
    nUtilizado As Integer
    nSaldo As Integer
End Type

Private Type aEventoType
    sNome As String
    nPontos As Integer
End Type

Dim bExec As Boolean, aPontos() As Integer, aEvento() As aEventoType

Private Sub cmbAno_Click()
CarregaMes
End Sub

Private Sub cmbFiscal_Click()
CarregaMes
End Sub

Private Sub cmbMes_Click()
CarregaMes
End Sub

Private Sub cmdAlterar_Click()
If cmbFiscal.ListIndex = -1 Then
    MsgBox "Selecione um fiscal.", vbCritical, "Atenção"
    Exit Sub
End If
ControlBehaviour False
grdMain.col = 1
grdMain.Row = 1
grdMain_EnterCell
End Sub

Private Sub cmdCalc_Click()
Dim nRow As Integer, nCol As Integer, nValor As Integer, x As Integer

For nRow = 0 To 12
    nValor = Val(grdMain.TextMatrix(nRow, 1))
    grdMain.TextMatrix(nRow, 2) = ""
    If nValor > 900 Then
       nValor = 900
    End If
    If nValor > 0 Then
        For x = 2 To 5
            grdMain.TextMatrix(nRow, x) = "0"
        Next
        If nValor < 200 Then
            grdMain.TextMatrix(nRow, 5) = nValor
        Else
            If nValor > 600 Then
'                grdMain.TextMatrix(nRow, 2) = nValor - 600
            Else
                grdMain.TextMatrix(nRow, 3) = 600 - nValor
            End If
                    
            If nValor >= 600 Then
                grdMain.TextMatrix(nRow, 4) = 600
            Else
                'abate os saldos anteriores
            
            
                grdMain.TextMatrix(nRow, 4) = nValor
            End If
        
            grdMain.TextMatrix(nRow, 5) = nValor - Val(grdMain.TextMatrix(nRow, 4))
        
        End If
                
                
    End If
Next

CalculaTotal
End Sub

Private Sub cmdCancel_Click()
ControlBehaviour True
txtEdit.Visible = False
End Sub

Private Sub cmdEndMounth_Click()
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

ProdutividadeFinalizarMes
End Sub

Private Sub cmdGravar_Click()

If MsgBox("Salvar alterações?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

SaveRecord
ControlBehaviour True
End Sub


Private Sub cmdPrint_Click()
Dim sFileName As String, x As Integer, ax As String

sFileName = "Saldo.txt"
FF1 = FreeFile()
Open sPathBin & "\" & sFileName For Output As FF1

'******* HEADER ********
Print #FF1, "SALDO DE PONTOS - PRODUTIVIDADE"
Print #FF1, "MÊS REFERÊNCIA: " & cmbMes.Text & "/" & cmbAno.Text
Print #FF1, "FISCAL: " & cmbFiscal.Text
Print #FF1, " "
Print #FF1, "====================================================="
Print #FF1, "MÊS/ANO        PONTOS  NEGATIVOS  UTILIZADOS  SALDO"
Print #FF1, "====================================================="

'******* BODY ********
With grdMain
    For x = 1 To .Rows - 1
        ax = FillSpace(.TextMatrix(x, 0), 17) & FillSpace(.TextMatrix(x, 1), 9) & FillSpace(Format(.TextMatrix(x, 3), "000"), 11) & FillSpace(Format(.TextMatrix(x, 4), "000"), 10) & FillSpace(Format(.TextMatrix(x, 5), "000"), 5)
        Print #FF1, ax
        Print #FF1, " "
    Next
End With

Close #FF1
Liberado

z = Shell("NOTEPAD" & " " & sPathBin & "\" & sFileName, vbNormalFocus)
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset, x As Integer
On Error GoTo Erro:
Centraliza Me
ControlBehaviour True
bExec = True
'If UCase(NomeDeLogin) = "SCHWARTZ" Then cmdEndMounth.Visible = True
grdMain.ColWidth(2) = 0
grdMain.ColWidth(6) = 0
grdMain.ColWidth(7) = 0

Sql = "select codigo,nome,nomecompleto from produtividadefiscal inner join "
Sql = Sql & "usuario on produtividadefiscal.nome = usuario.nomelogin where calculo=1 order by nomecompleto "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFiscal.AddItem !NomeCompleto
        cmbFiscal.ItemData(cmbFiscal.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

If NomeDeLogin <> "SCHWARTZ" Then
    If Not ProdIsBossLogin() Then
        cmbFiscal.Text = RetornaUsuarioFullName()
        cmbFiscal.Enabled = False
        cmdAlterar.Enabled = False
    End If
End If

bExec = False
For x = 2011 To Year(Now) + 1
    cmbAno.AddItem (CStr(x))
Next

cmbMes.ListIndex = Month(Now) - 1
cmbAno.Text = Year(Now)
bExec = True
CarregaMes
CarregaEvento

Exit Sub
Erro:
MsgBox "Erro Fatal!"
cmbFiscal.Enabled = False
cmdAlterar.Enabled = False
cmdPrint.Enabled = False
End Sub

Private Sub ControlBehaviour(bStart As Boolean)
cmdAlterar.Visible = False
cmdPrint.Visible = bStart
cmdGravar.Visible = False
cmdCancel.Visible = False
cmbFiscal.Enabled = bStart
cmbMes.Enabled = bStart
cmbAno.Enabled = bStart

End Sub

Private Sub grdMain_EnterCell()
Exit Sub
If grdMain.col = 2 Then Exit Sub
If cmdAlterar.Visible Then Exit Sub
If grdMain.Row = grdMain.Rows - 1 Then
    txtEdit.Visible = False
    Exit Sub
End If
If txtEdit.Visible = False Then txtEdit.Visible = True
txtEdit.Text = grdMain.TextMatrix(grdMain.Row, grdMain.col)
txtEdit.Left = grdMain.Left + grdMain.CellLeft
txtEdit.Top = grdMain.Top + grdMain.CellTop
txtEdit.Width = grdMain.CellWidth
txtEdit.Height = grdMain.CellHeight
txtEdit.SetFocus
End Sub

Private Sub grdMain_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If txtEdit.Visible = False Then grdMain_EnterCell
End If
End Sub

Private Sub grdMain_LeaveCell()
If grdMain.col = 2 Then Exit Sub

If grdMain.Row = grdMain.Rows - 1 Then Exit Sub
If cmdGravar.Visible = True Then
    If txtEdit.Text = "" Then txtEdit.Text = "0"
    grdMain.TextMatrix(grdMain.Row, grdMain.col) = txtEdit.Text
End If
End Sub

Private Sub txtEdit_GotFocus()
txtEdit.SelStart = 0
txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
'If Not bExec Then Exit Sub
If KeyCode = 38 Then 'up
    grdMain_LeaveCell
    If grdMain.Row > 1 Then
        grdMain.Row = grdMain.Row - 1
    End If
    grdMain_EnterCell
ElseIf KeyCode = 40 Then 'down
    grdMain_LeaveCell
    If grdMain.Row < grdMain.Rows - 2 Then
        grdMain.Row = grdMain.Row + 1
    End If
    grdMain_EnterCell
ElseIf KeyCode = 37 Then 'left
    grdMain_LeaveCell
    If grdMain.col > 1 Then
        grdMain.col = grdMain.col - 1
    End If
    grdMain_EnterCell
ElseIf KeyCode = 39 Then 'right
    grdMain_LeaveCell
    If grdMain.col < 5 Then
        grdMain.col = grdMain.col + 1
    End If
    grdMain_EnterCell
ElseIf KeyCode = 27 Then 'ESCAPE
    txtEdit.Visible = False
End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
Dim nValor As Integer, nCol As Integer, nRow As Integer

nCol = grdMain.col
nRow = grdMain.Row

If KeyAscii = vbKeyReturn Then
    If txtEdit.Visible = False Then txtEdit.Visible = True
    KeyAscii = 0
    grdMain_LeaveCell
    If nRow < grdMain.Rows - 2 Then
        grdMain.Row = grdMain.Row + 1
    Else
        grdMain.Row = grdMain.Rows - 2
    End If
    
    grdMain_EnterCell
Else
    Tweak txtEdit, KeyAscii, IntegerPositive
End If

End Sub

Private Sub CarregaMes()
Dim x As Integer, nMes As Integer, nAno As Integer, sMesAno As String
If cmbMes.ListIndex = -1 Then Exit Sub
If Not bExec Then Exit Sub
Clear
nMes = cmbMes.ItemData(cmbMes.ListIndex)
nAno = Val(cmbAno.Text)

For x = 12 To 1 Step -1
    sMesAno = MonthName(nMes) & "/" & CStr(nAno)
    grdMain.TextMatrix(x, 0) = sMesAno
    grdMain.TextMatrix(x, 6) = CStr(nMes)
    grdMain.TextMatrix(x, 7) = CStr(nAno)
    nMes = nMes - 1
    If nMes = 0 Then
        nMes = 12
        nAno = nAno - 1
    End If
Next

grdMain.TextMatrix(grdMain.Rows - 1, 0) = "TOTAL -->"
Le
End Sub

Private Sub SaveRecord()
Dim Sql As String, x As Integer, nMes As Integer, nAno As Integer, nCodFiscal As Integer
Dim nPontos As Integer, nNegativos As Integer, nUtilizados As Integer, nSaldo As Integer
Dim nMesRef As Integer, nAnoRef As Integer

nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
nMesRef = cmbMes.ItemData(cmbMes.ListIndex)
nAnoRef = Val(cmbAno.Text)

Sql = "DELETE FROM PRODUTIVIDADESALDO WHERE CODFISCAL=" & nCodFiscal & " AND ANOREF=" & nAnoRef & " AND MESREF=" & nMesRef
cn.Execute Sql, rdExecDirect

With grdMain
    For x = 1 To .Rows - 2
        nPontos = Val(.TextMatrix(x, 1))
        nNegativos = Val(.TextMatrix(x, 3))
        nUtilizados = Val(.TextMatrix(x, 4))
        nSaldo = Val(.TextMatrix(x, 5))
        nMes = Val(.TextMatrix(x, 6))
        nAno = Val(.TextMatrix(x, 7))
        
        Sql = "INSERT PRODUTIVIDADESALDO(CODFISCAL,ANOREF,MESREF,ANO,MES,PONTOS,NEGATIVOS,UTILIZADOS,SALDO) VALUES("
        Sql = Sql & nCodFiscal & "," & nAnoRef & "," & nMesRef & "," & nAno & "," & nMes & "," & nPontos & "," & nNegativos & ","
        Sql = Sql & nUtilizados & "," & nSaldo & ")"
        cn.Execute Sql, rdExecDirect
        
    Next
End With

txtEdit.Visible = False
End Sub

Private Sub Clear()
Dim x As Integer, Y As Integer

With grdMain
    For x = 1 To .Rows - 1
        For Y = 1 To .Cols - 1
            If Y <> 2 Then
                .TextMatrix(x, Y) = "0"
            End If
        Next
    Next
End With

End Sub

Private Sub CalculaTotal()
Dim nSoma As Integer


For Y = 1 To grdMain.Cols - 1
    nSoma = 0
    For x = 1 To grdMain.Rows - 2
        nSoma = nSoma + Val(grdMain.TextMatrix(x, Y))
    Next
    If Y <> 2 Then
        grdMain.TextMatrix(grdMain.Rows - 1, Y) = nSoma
    End If
Next

End Sub

Private Sub Le()
Dim x As Integer, nMes As Integer, nAno As Integer, nCodFiscal As Integer
Dim Sql As String, RdoAux As rdoResultset, nAnoRef As Integer, nMesRef As Integer

If cmbFiscal.ListIndex = -1 Then Exit Sub
nCodFiscal = cmbFiscal.ItemData(cmbFiscal.ListIndex)
nMes = cmbMes.ItemData(cmbMes.ListIndex)
nAno = Val(cmbAno.Text)

Sql = "SELECT * FROM PRODUTIVIDADESALDO WHERE CODFISCAL=" & nCodFiscal & " AND ANOREF=" & nAno
Sql = Sql & " AND MESREF=" & nMes & " ORDER BY ANO DESC,MES"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nAnoRef = !Ano
        nMesRef = !Mes
        For x = 1 To grdMain.Rows - 2
            If nMesRef = Val(grdMain.TextMatrix(x, 6)) And nAnoRef = Val(grdMain.TextMatrix(x, 7)) Then
                Exit For
            End If
        Next
        grdMain.TextMatrix(x, 1) = !Pontos
        grdMain.TextMatrix(x, 3) = !negativos
        grdMain.TextMatrix(x, 4) = !utilizados
        grdMain.TextMatrix(x, 5) = !saldo
       .MoveNext
    Loop
   .Close
End With

CalculaTotal
End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub ProdutividadeFinalizarMes()
Dim Sql As String, RdoAux As rdoResultset, nCodFiscal As Integer, nMesRef As Integer, nAnoRef As Integer
Dim x As Integer, nMes As Integer, nAno As Integer, aTabela() As TabelaType, RdoAux2 As rdoResultset
Dim nMesRefOld As Integer, nAnoRefOld As Integer, nPos As Integer, nPontosAtual As Integer, nSomaSaldo As Integer
Dim nValorACompensar As Integer, nSomaCompensado As Integer, nUtilizou As Integer, nFirstYear As Integer, nFirstMonth As Integer

nFirstMonth = 0
nFirstYear = 0
ReDim aTabela(0)
'define os meses para cálculo
If Month(Now) = 1 Then
    nMesRef = 12
    nAnoRef = Year(Now) - 1
Else
    nMesRef = Month(Now) - 1
    nAnoRef = Year(Now)
End If

If nMesRef = 1 Then
    nMesRefOld = 12
    nAnoRefOld = nAnoRef - 1
Else
    nMesRefOld = nMesRef - 1
    nAnoRefOld = nAnoRef
End If

'executa o cálculo para cada um dos fiscais
'Sql = "SELECT codigo FROM produtividadefiscal WHERE codigo=8 and  calculo=1 ORDER BY codigo"
Sql = "SELECT codigo FROM produtividadefiscal where  calculo=1 ORDER BY codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
Do Until RdoAux.EOF
    nCodFiscal = RdoAux!Codigo
    'nCodFiscal = 9
    'busca saldo do mes de fechamento
    nPontosAtual = FillArray(nCodFiscal, nMesRef, nAnoRef)
    
    'pontos atuais virtuais
    'apenas no primeiro mes de funcionamento, depois apagar estas linhas
'    If nAnoRef = 2012 And nMesRef = 4 Then
'        If nCodFiscal = 1 Then 'eduardo
'            nPontosAtual = 570
'        ElseIf nCodFiscal = 3 Then 'marta
'            nPontosAtual = 560
'        ElseIf nCodFiscal = 4 Then 'daniela
'            nPontosAtual = 581
'        ElseIf nCodFiscal = 5 Then 'ana
'            nPontosAtual = 554
'        ElseIf nCodFiscal = 6 Then 'rosangela
'            nPontosAtual = 492
'        ElseIf nCodFiscal = 7 Then 'luiz
'            nPontosAtual = 595
'        ElseIf nCodFiscal = 8 Then 'carmesciano
'            nPontosAtual = 250
'        ElseIf nCodFiscal = 9 Then 'rita
'            nPontosAtual = 377
'
'        End If
'    End If
    
    '**********************
    
    nSomaSaldo = 0
    'busca extrato anterior para sabermos os saldos disponiveis
    Sql = "select * from produtividadesaldo where codfiscal=" & nCodFiscal & " and anoref=" & nAnoRefOld
    Sql = Sql & " and mesref=" & nMesRefOld & " order by ano,mes"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Do Until .EOF
            If nFirstMonth = 0 Then
                nFirstMonth = !Mes
                nFirstYear = !Ano
            End If
            nPos = UBound(aTabela) + 1
            ReDim Preserve aTabela(nPos)
            aTabela(nPos).nCodFiscal = nCodFiscal
            aTabela(nPos).nAno = !Ano
            aTabela(nPos).nMes = !Mes
            aTabela(nPos).nSaldo = !saldo
            aTabela(nPos).nPontos = !Pontos
            aTabela(nPos).nUtilizado = !utilizados
            aTabela(nPos).nNegativo = !negativos
            nSomaSaldo = nSomaSaldo + !saldo
           .MoveNext
        Loop
       .Close
    End With
    
    'cria uma posição a mais na matriz para o saldo de pontos atual
    ReDim Preserve aTabela(UBound(aTabela) + 1)
    nPos = UBound(aTabela)
    aTabela(nPos).nCodFiscal = nCodFiscal
    aTabela(nPos).nAno = nAnoRef
    aTabela(nPos).nMes = nMesRef
    
    'se os pontos atuais forem maiores que 600
    'última posição da matriz conterá:
    '-pontos=pontosatuais(<=900)  / utilizado=600 / negativos=0 / saldo=pontosatuais-600
    If nPontosAtual > 900 Then nPontosAtual = 900
    If nPontosAtual >= 600 And nPontosAtual <= 900 Then
        aTabela(nPos).nPontos = nPontosAtual
        aTabela(nPos).nUtilizado = 600
        aTabela(nPos).nNegativo = 0
        aTabela(nPos).nSaldo = nPontosAtual - 600
    End If
    
    'se os pontos atuais forem maiores que 200 e menores que 600
    'última posição da matriz conterá:
    '-pontos=pontosatuais  / utilizado=pontosatuais + saldo (<=600) / negativos=600-pontosatuais / saldo=0
    nUtilizou = 0
    If nPontosAtual >= 200 And nPontosAtual < 600 Then
        'rotina para compesação dos saldos anteriores (se houver saldo)
        If nSomaSaldo > 0 Then
            'o valor a compensar será:
            nValorACompensar = 600 - nPontosAtual
            nSomaCompensado = 0
            For x = 1 To UBound(aTabela) - 1
                If aTabela(x).nCodFiscal = nCodFiscal Then
                    If aTabela(x).nSaldo > 0 Then
                        If nPontosAtual + nSomaCompensado + aTabela(x).nSaldo <= 600 Then
                            nSomaCompensado = nSomaCompensado + aTabela(x).nSaldo
                            nUtilizou = nUtilizou + nSomaCompensado
                            aTabela(x).nSaldo = 0
                        ElseIf nPontosAtual + nSomaCompensado + aTabela(x).nSaldo > 600 Then
                            nSomaCompensado = (nPontosAtual + nSomaCompensado + aTabela(x).nSaldo) - 600
                            nUtilizou = nUtilizou + nSomaCompensado
                            aTabela(x).nSaldo = nSomaCompensado
                            Exit For
                        End If
                        
                    End If
                End If
            Next
        End If
        '********************************************
        aTabela(nPos).nPontos = nPontosAtual
        aTabela(nPos).nUtilizado = nPontosAtual + nUtilizou 'x será a nova rotina de baixa de saldo
        aTabela(nPos).nNegativo = 600 - nPontosAtual
        aTabela(nPos).nSaldo = 0
    End If
    
    
    'se os pontos atuais forem menores que 200
    'última posição da matriz conterá:
    '-pontos=pontosatuais  / utilizado=0 / negativos=0 / saldo=npontos
    If nPontosAtual >= 200 And nPontosAtual < 600 Then
        aTabela(nPos).nPontos = nPontosAtual
        aTabela(nPos).nUtilizado = 0
        aTabela(nPos).nNegativo = 0
        aTabela(nPos).nSaldo = nPontos
    End If
    
    
    RdoAux.MoveNext
Loop
RdoAux.Close

'grava a matriz na base de dados
'criando o novo mes de produtividade

For x = 2 To UBound(aTabela)
    If aTabela(x).nAno = nFirstYear And aTabela(x).nMes = nFirstMonth Then
    Else
        Sql = "insert produtividadesaldo(codfiscal,anoref,mesref,ano,mes,pontos,negativos,utilizados,saldo) values("
        Sql = Sql & aTabela(x).nCodFiscal & "," & nAnoRef & "," & nMesRef & "," & aTabela(x).nAno & "," & aTabela(x).nMes & ","
        Sql = Sql & aTabela(x).nPontos & "," & aTabela(x).nNegativo & "," & aTabela(x).nUtilizado & "," & aTabela(x).nSaldo & ")"
        cn.Execute Sql, rdExecDirect
    End If
Next

MsgBox "fim"
End Sub

Private Function FillArray(nCodFiscal As Integer, nMes As Integer, nAno As Integer) As Integer
Dim sData As String, nDay As Long, nWeekDay As Long, nSomaPontos As Integer
Dim nCodEvento As Integer, sEvento As String, nPontos As Integer
Dim bIsBoss As Boolean, x As Integer, aExtrato() As ExtratoType

FillPontosMes nCodFiscal, nMes, nAno
ReDim aExtrato(0)
nSaldoAnterior = 423 'CORRIGIR
bIsBoss = ProdIsBoss(nCodFiscal)
nLastDay = Val(Left(Format$(DateSerial(nAno, Val(nMes) + 1, 0), "dd/mm/yyyy"), 2))
'ReDim aPontos(nLastDay)

For nDay = 1 To nLastDay
    ReDim Preserve aExtrato(UBound(aExtrato) + 1)
    aExtrato(nDay).nDia = nDay
    sData = Format(nDay, "00") & "/" & Format(nMes, "00") & "/" & cmbAno.Text
    nWeekDay = Weekday(CDate(sData))
    aExtrato(nDay).sDia = WeekdayName(nWeekDay, True)

    nCodEvento = ProdEventoDia(nCodFiscal, CDate(sData))
    If nCodEvento > 0 Then
        If nWeekDay = 1 Or nWeekDay = 7 Then
            sEvento = "SÁB/DOM/FER."
            nPontos = 0
        Else
            If nCodEvento = 3 Then 'LICENÇA PREMIO NÃO TEM PONTOS
                nPontos = 0
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
            nPontos = 0
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
    
Next

nSomaPontos = 0
For x = 1 To UBound(aExtrato)
    nSomaPontos = nSomaPontos + aExtrato(x).nPontos
Next

FillArray = nSomaPontos
End Function

Private Sub FillPontosMes(nCodFiscal As Integer, nMes As Integer, nAno As Integer)
Dim Sql As String, RdoAux As rdoResultset, nSomaTarefa As Integer, nDia As Integer
Dim bAchou As Boolean, nPos As Integer, sNome As String

nLastDay = Val(Left(Format$(DateSerial(nAno, Val(nMes) + 1, 0), "dd/mm/yyyy"), 2))
ReDim aPontos(nLastDay)
ReDim aExtratoItem(0)

Sql = "SELECT produtividadetarefa.data, produtividadetarefa.item, produtividadetarefa.qtde, produtividadetarefa.valor, produtividadetarefa.processo, "
Sql = Sql & "produtividadedesc.descricao FROM produtividadetarefa INNER JOIN produtividadedesc ON produtividadetarefa.item = produtividadedesc.item "
Sql = Sql & "where year(data)=" & nAno & " and month(data)=" & nMes & " and fiscal=" & nCodFiscal
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nDia = Day(!Data)
        nSomaTarefa = !valor * !QTDE
        nSomaMes = nSomaMes + nSomaTarefa
        aPontos(nDia) = aPontos(nDia) + nSomaTarefa
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaEvento()
Dim Sql As String, RdoAux As rdoResultset

ReDim aEvento(0)
Sql = "select codigo,nome,pontodia from produtividadeevento order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aEvento(UBound(aEvento) + 1)
        aEvento(UBound(aEvento)).sNome = !NOME
        aEvento(UBound(aEvento)).nPontos = !pontodia
       .MoveNext
    Loop
   .Close
End With

End Sub

