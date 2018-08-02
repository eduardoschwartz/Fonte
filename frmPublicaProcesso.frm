VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmPublicaProcesso 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Publicação de Processos"
   ClientHeight    =   4485
   ClientLeft      =   6015
   ClientTop       =   3480
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   8310
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1890
      TabIndex        =   5
      Top             =   2910
      Width           =   1365
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5130
      TabIndex        =   4
      Top             =   2880
      Width           =   1365
   End
   Begin VB.ComboBox cmbDespacho 
      Height          =   315
      ItemData        =   "frmPublicaProcesso.frx":0000
      Left            =   1890
      List            =   "frmPublicaProcesso.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3570
      Width           =   4605
   End
   Begin VB.TextBox txtEdital 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2220
      TabIndex        =   1
      Top             =   120
      Width           =   1755
   End
   Begin MSFlexGridLib.MSFlexGrid grdProc 
      Height          =   2235
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   3942
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   $"frmPublicaProcesso.frx":0036
   End
   Begin prjChameleon.chameleonButton cmdRefresh 
      Height          =   345
      Left            =   4080
      TabIndex        =   2
      ToolTipText     =   "Carregar Edital"
      Top             =   90
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "!"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   192
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPublicaProcesso.frx":00C1
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
      Left            =   5970
      TabIndex        =   11
      ToolTipText     =   "Gravar os Dados"
      Top             =   4050
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPublicaProcesso.frx":00DD
      PICN            =   "frmPublicaProcesso.frx":00F9
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
      Left            =   7050
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   4050
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPublicaProcesso.frx":049E
      PICN            =   "frmPublicaProcesso.frx":04BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   2250
      TabIndex        =   13
      ToolTipText     =   "Excluir Registro"
      Top             =   4050
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmPublicaProcesso.frx":0528
      PICN            =   "frmPublicaProcesso.frx":0544
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1200
      TabIndex        =   14
      ToolTipText     =   "Editar Registro"
      Top             =   4050
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
      MICON           =   "frmPublicaProcesso.frx":05E6
      PICN            =   "frmPublicaProcesso.frx":0602
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   150
      TabIndex        =   15
      ToolTipText     =   "Novo Registro"
      Top             =   4050
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmPublicaProcesso.frx":075C
      PICN            =   "frmPublicaProcesso.frx":0778
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
      Left            =   3330
      TabIndex        =   16
      ToolTipText     =   "Publicar"
      Top             =   4050
      Width           =   1050
      _ExtentX        =   1852
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
      MICON           =   "frmPublicaProcesso.frx":08D2
      PICN            =   "frmPublicaProcesso.frx":08EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7050
      TabIndex        =   17
      ToolTipText     =   "Cancelar Edição"
      Top             =   4020
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
      MICON           =   "frmPublicaProcesso.frx":0A48
      PICN            =   "frmPublicaProcesso.frx":0A64
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
      Caption         =   "Número do processo..:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   2970
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Ano do processo..:"
      Height          =   255
      Index           =   1
      Left            =   3660
      TabIndex        =   9
      Top             =   2940
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do requerente..:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   3300
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de despacho.....:"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   7
      Top             =   3630
      Width           =   1635
   End
   Begin VB.Label lblReq 
      ForeColor       =   &H00400000&
      Height          =   225
      Left            =   1890
      TabIndex        =   6
      Top             =   3300
      Width           =   5985
   End
End
Attribute VB_Name = "frmPublicaProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset
Dim Evento As String

Private Sub cmdAlterar_Click()
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
End Sub

Private Sub cmdExcluir_Click()

If grdProc.Rows = 1 Then Exit Sub
Sql = "DELETE FROM PUBLICACAO WHERE EDITAL='" & txtEdital.text & "' AND ANO=" & Val(txtAno.text) & " AND NUMERO=" & Val(txtNumero.text)
cn.Execute Sql, rdExecDirect
If grdProc.Rows > 2 Then
    grdProc.RemoveItem (grdProc.Row)
    Limpa
    grdProc.Row = 1
Else
    grdProc.Rows = 1
    Limpa
End If
grdProc.ColSel = 2
End Sub

Private Sub cmdGravar_Click()
Dim x As Integer, sNum As String, bAchou As Boolean
If lblReq.Caption = "" Then
    MsgBox "Processo inválido.", vbExclamation, "Atenção"
    Exit Sub
End If
If cmbDespacho.ListIndex = -1 Then
    MsgBox "Selecione o despacho.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
sNum = txtNumero.text & "/" & txtAno.text
With grdProc
    For x = 1 To .Rows - 1
        If sNum = .TextMatrix(x, 1) Then
            bAchou = True
            Exit For
        End If
    Next
End With
If bAchou Then
    MsgBox "Processo já incluido.", vbExclamation, "Atenção"
    Exit Sub
Else
    If Evento = "Novo" Then
        grdProc.AddItem sNum & Chr(9) & lblReq.Caption & Chr(9) & cmbDespacho.text
        Sql = "INSERT PUBLICACAO (EDITAL,ANO,NUMERO,REQUERENTE,DESPACHO) VALUES('" & Mask(txtEdital.text) & "',"
        Sql = Sql & Val(txtAno.text) & "," & Val(txtNumero.text) & ",'" & Mask(Left$(lblReq.Caption, 40)) & "','" & cmbDespacho.text & "')"
    Else
        grdProc.TextMatrix(grdProc.Row, 2) = cmbDespacho.text
        Sql = "UPDATE PUBLICACAO SET DESPACHO='" & cmbDespacho.text & "' WHERE EDITAL='" & txtEdital.text & "' AND ANO=" & Val(txtAno.text) & " AND NUMERO=" & Val(txtNumero.text)
    End If
    cn.Execute Sql, rdExecDirect
End If

Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub cmdPrint_Click()
Dim sNomeArq As String

If grdProc.Rows = 1 Then
    MsgBox "Nada a imprimir.", vbCritical, "Atenção"
    Exit Sub
End If

sNomeArq = sPathBin & "\EDITAL.TXT"
Open sNomeArq For Output As #1
    Print #1, "**************************************************************"
    Print #1, "EDITAL DE PUBLICAÇÃO No " & txtEdital.text
    Print #1, "**************************************************************"
    Print #1, ""
    
    With grdProc
        For x = 1 To .Rows - 1
            ax = FillSpace(.TextMatrix(x, 0), 12) & " - " & FillSpace(.TextMatrix(x, 1), 40) & " - " & FillSpace(.TextMatrix(x, 2), 30)
            Print #1, ax
        Next
    End With
Close #1
x = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)
End Sub

Private Sub cmdRefresh_Click()
Sql = "SELECT * FROM PUBLICACAO WHERE EDITAL='" & txtEdital.text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        If MsgBox("Deseja abrir o Edital No " & txtEdital.text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            grdProc.Rows = 1
            Limpa
            Evento = "Novo"
            Eventos "INCLUIR"
            txtNumero.SetFocus
        End If
    Else
        grdProc.Rows = 1
        Do Until .EOF
            grdProc.AddItem !Numero & "/" & !Ano & Chr(9) & !REQUERENTE & Chr(9) & !DESPACHO
           .MoveNext
        Loop
    End If
   .Close
End With
If grdProc.Rows > 1 Then
    grdProc_RowColChange
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Eventos "INICIAR"
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdPrint.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmPublicaProcesso
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = Kde
           Ct.Enabled = False
       End If
   Next
   txtEdital.Enabled = True
   txtEdital.BackColor = Branco
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdPrint.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmPublicaProcesso
       If TypeOf Ct Is TextBox Or TypeOf Ct Is ComboBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   txtEdital.Enabled = False
   txtEdital.BackColor = Kde
   If Evento = "Alterar" Then
        txtAno.Enabled = False
        txtAno.BackColor = Kde
        txtNumero.Enabled = False
        txtNumero.BackColor = Kde
   End If
End If

End Sub

Private Sub grdProc_RowColChange()
With grdProc
    If grdProc.Row > 0 Then
        txtAno.text = Val(Right$(.TextMatrix(.Row, 0), 4))
        txtNumero.text = Val(Left$(.TextMatrix(.Row, 0), Len(.TextMatrix(.Row, 0)) - 5))
        lblReq.Caption = .TextMatrix(.Row, 1)
        cmbDespacho.text = .TextMatrix(.Row, 2)
    End If
End With

End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtAno_LostFocus()
CarregaRequerente
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
Tweak txtNumero, KeyAscii, IntegerPositive
End Sub

Private Sub CarregaRequerente()
Dim nCodCidadao As Long
If Val(txtAno.text) = 0 Or Val(txtNumero.text) = 0 Then Exit Sub
lblReq.Caption = ""

Sql = "SELECT CODCIDADAO FROM PROCESSOGTI WHERE ANO=" & Val(txtAno.text) & " AND NUMERO=" & Val(Left$(txtNumero.text, Len(txtNumero.text) - 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Processo não cadastrado ", vbExclamation, "Atenção"
        Exit Sub
    Else
        nCodCidadao = !CodCidadao
    End If
   .Close
End With
Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblReq.Caption = !NOMECIDADAO
    End If
   .Close
End With

End Sub

Private Sub Limpa()
txtNumero.text = ""
txtAno.text = ""
lblReq.Caption = ""
cmbDespacho.ListIndex = -1
End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

