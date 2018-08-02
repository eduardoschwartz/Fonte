VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDebBanco 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa de Débito Automático"
   ClientHeight    =   5640
   ClientLeft      =   1380
   ClientTop       =   2235
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8805
   Begin prjChameleon.chameleonButton cmdShow 
      Height          =   315
      Left            =   5400
      TabIndex        =   20
      ToolTipText     =   "Visualizar Arquivo na Origem"
      Top             =   1560
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Visualizar"
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
      MICON           =   "frmDebBanco.frx":0000
      PICN            =   "frmDebBanco.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdF 
      Height          =   3165
      Left            =   30
      TabIndex        =   14
      Top             =   1950
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   5583
      _Version        =   393216
      Rows            =   20
      Cols            =   13
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDebBanco.frx":0176
   End
   Begin VB.PictureBox ImgLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   0
      ScaleHeight     =   1170
      ScaleWidth      =   2685
      TabIndex        =   0
      Top             =   0
      Width           =   2715
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   1845
      TabIndex        =   17
      ToolTipText     =   "Reativação do Arquivo"
      Top             =   5250
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Reativação"
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
      MICON           =   "frmDebBanco.frx":023D
      PICN            =   "frmDebBanco.frx":0259
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
      Height          =   315
      Left            =   60
      TabIndex        =   13
      ToolTipText     =   "Efetuar Baixa"
      Top             =   5250
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&fetuar Baixa"
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
      MICON           =   "frmDebBanco.frx":03B3
      PICN            =   "frmDebBanco.frx":03CF
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
      Left            =   3210
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   5250
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
      MICON           =   "frmDebBanco.frx":046E
      PICN            =   "frmDebBanco.frx":048A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdArq 
      Height          =   1005
      Left            =   2790
      TabIndex        =   1
      Top             =   120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1773
      _Version        =   393216
      Rows            =   3
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      AllowBigSelection=   0   'False
      GridLines       =   0
      GridLinesFixed  =   0
      BorderStyle     =   0
      Appearance      =   0
      FormatString    =   " |Arquivos Disponíveis                      "
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   105
      Left            =   5790
      TabIndex        =   11
      Top             =   5430
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   185
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grdParc 
      Height          =   1155
      Left            =   30
      TabIndex        =   15
      Top             =   5940
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   2037
      _Version        =   393216
      Rows            =   1
      Cols            =   17
      FixedCols       =   0
      BackColorSel    =   12582912
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDebBanco.frx":04F8
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   1155
      Left            =   -45
      TabIndex        =   16
      Top             =   6750
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   2037
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   12582912
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Codigo        |>Vl.Lançado  |>Vl.Multa     |>Vl.Juros       |>Vl.Correção    |>Vl.Total                |>Linha "
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Remessa:"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   30
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Convênio:"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   29
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Geração:"
      Height          =   255
      Index           =   2
      Left            =   2490
      TabIndex        =   28
      Top             =   1290
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Sequencial....:"
      Height          =   255
      Index           =   3
      Left            =   2490
      TabIndex        =   27
      Top             =   1590
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Versão Layout.....:"
      Height          =   255
      Index           =   4
      Left            =   5430
      TabIndex        =   26
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblCR 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1590
      TabIndex        =   25
      Top             =   1290
      Width           =   675
   End
   Begin VB.Label lblCC 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1590
      TabIndex        =   24
      Top             =   1590
      Width           =   735
   End
   Begin VB.Label lblDG 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3870
      TabIndex        =   23
      Top             =   1290
      Width           =   1125
   End
   Begin VB.Label lblNS 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   3870
      TabIndex        =   22
      Top             =   1590
      Width           =   675
   End
   Begin VB.Label lblVL 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   6810
      TabIndex        =   21
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label lblEfetuado 
      BackStyle       =   0  'Transparent
      Caption         =   "R$ "
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6930
      TabIndex        =   19
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Efetuado..:"
      Height          =   255
      Left            =   5700
      TabIndex        =   18
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "]"
      Height          =   225
      Index           =   1
      Left            =   8670
      TabIndex        =   10
      Top             =   5370
      Width           =   45
   End
   Begin VB.Label Label5 
      Caption         =   "["
      Height          =   225
      Index           =   0
      Left            =   5700
      TabIndex        =   9
      Top             =   5370
      Width           =   45
   End
   Begin VB.Label lblPb 
      BackColor       =   &H00EEEEEE&
      Height          =   225
      Left            =   5850
      TabIndex        =   8
      Top             =   5190
      Width           =   2655
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   165
      Left            =   6930
      TabIndex        =   7
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco..............:"
      Height          =   225
      Left            =   5700
      TabIndex        =   6
      Top             =   90
      Width           =   1185
   End
   Begin VB.Label lblRegTot 
      BackStyle       =   0  'Transparent
      Caption         =   "R$ "
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6930
      TabIndex        =   5
      Top             =   630
      Width           =   1215
   End
   Begin VB.Label lblNumReg 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6930
      TabIndex        =   4
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total........:"
      Height          =   225
      Left            =   5700
      TabIndex        =   3
      Top             =   630
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Registros:"
      Height          =   225
      Left            =   5700
      TabIndex        =   2
      Top             =   360
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1200
      Left            =   2700
      Top             =   0
      Width           =   6075
   End
End
Attribute VB_Name = "frmDebBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ControleFebraban
    Registro As String * 150
End Type

Private Type FebrabanA   'HEADER DO ARQUIVO
    CodigoRegistro As String * 1
    CodigoRemessa As String * 1
    CodigoConvenio As String * 20
    NomeEmpresa As String * 20
    CodigoBanco As String * 3
    NomeBanco As String * 20
    DataGeracao As String * 8
    NumeroSeq As String * 6
    VersaoLayout As String * 2
    Filler As String * 69
End Type

Private Type FebrabanB 'CADASTRAMENTO DE DEBITO AUTOMATICO
   CodigoRegistro As String * 1
   Distrito As String * 2
   Setor As String * 2
   Quadra As String * 4
   Lote As String * 5
   Seq As String * 2
   FillerID As String * 10
   CodAgencia As String * 4
   ContaCliente As String * 14
   DataOpcao As String * 8
   Filler As String * 97
   CodMovimento As String * 1
End Type

Private Type FebrabanF 'RETORNO DO DEBITO AUTOMATICO
   CodigoRegistro As String * 1
   Distrito As String * 2
   Setor As String * 2
   Quadra As String * 4
   Lote As String * 5
   Seq As String * 2
   FillerID As String * 10
   CodAgencia As String * 4
   ContaCliente As String * 14
   DataVencto As String * 8
   ValorDebito As String * 15
   CodRetorno As String * 2
   NumDoc As String * 9
   Filler1 As String * 51
   Filler2 As String * 20
   CodMovimento As String * 1
End Type

Private Type FebrabanH 'OCORRENCIA DE ALTERACAO CONTROLE EMPRESA
   CodigoRegistro As String * 1
   DistritoOld As String * 2
   SetorOld As String * 2
   QuadraOld As String * 4
   LoteOld As String * 5
   SeqOld As String * 2
   FillerIDOld As String * 10
   CodAgencia As String * 4
   ContaCliente As String * 14
   Distrito As String * 2
   Setor As String * 2
   Quadra As String * 4
   Lote As String * 5
   Seq As String * 2
   FillerID As String * 10
   Ocorrencia As String * 58
   Filler As String * 22
   CodMovimento As String * 1
End Type

Private Type FebrabanJ 'CONFIRMA O PROCESSAMENTO DE ARQUIVOS
   CodigoRegistro As String * 1
   CodigoNSA As String * 6
   DataGeracao As String * 8
   TotalRegistros As String * 6
   ValorTotal As String * 17
   DataProcessamento As String * 8
   Filler As String * 104
End Type

Private Type FebrabanT 'TOTAL CLIENTES DEBITADOS
   CodigoRegistro As String * 1
   TotalRegistros As String * 6
   ValorTotal As String * 17
   Filler As String * 126
End Type

Private Type FebrabanX 'RELACAO DE AGENCIAS
   CodigoRegistro As String * 1
   CodigoAgencia As String * 4
   NomeAgencia As String * 30
   EnderecoAgencia As String * 30
   NumAgencia As String * 5
   CepAgencia As String * 5
   SufixoCep As String * 3
   Cidade As String * 20
   SiglaUF As String * 2
   SituacaoAgencia As String * 1
   Filler As String * 49
End Type

Private Type FebrabanZ 'RODAPE DO ARQUIVO
    CodigoRegistro As String * 1
    TotalRegistro As String * 6
    ValorTotal As String * 17
    Filler As String * 126
End Type

Dim aFebrabanA() As FebrabanA 'HEADER DO ARQUIVO
Dim aFebrabanB() As FebrabanB 'CADASTRAMENTO DE DEBITO AUTOMATICO
Dim aFebrabanF() As FebrabanF 'RETORNO DO DEBITO AUTOMATICO
Dim aFebrabanH() As FebrabanH 'OCORRENCIA DE ALTERACAO CONTROLE EMPRESA
Dim aFebrabanJ() As FebrabanJ 'CONFIRMA O PROCESSAMENTO DE ARQUIVOS
Dim aFebrabanT() As FebrabanT 'TOTAL CLIENTES DEBITADOS
Dim aFebrabanX() As FebrabanX 'RELACAO DE AGENCIAS
Dim aFebrabanZ() As FebrabanZ 'RODAPE DO ARQUIVO
Dim nNumDoc As Long

Private Sub cmdBaixa_Click()

Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If Not IsNull(!DATABAIXA) Then
        MsgBox "Ja foi efetuado Baixa neste arquivo.", vbCritical, "Atenção"
       .Close
       Exit Sub
       End If
    End If
   .Close
End With

If MsgBox("Deseje efetuar a Baixa nas Parcelas ?", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE BAIXA") = vbNo Then Exit Sub

lblPB.Caption = "Cadastro de Optantes"
lblPB.Refresh
'BaixaRegB
Sql = "DELETE FROM BAIXATMP WHERE COMPUTADOR='" & Trim$(NomeDoComputador) & "'"
cn.Execute Sql, rdExecDirect

BaixaRegF

End Sub

Private Sub GravaBaixaTmp()
Dim qd As New rdoQuery
Dim x As Long
Dim sStatus As String
Dim nSomaTotal As Double
Dim nContaReg As Integer
Dim bDif As Boolean

Set qd.ActiveConnection = cn

nSomaTotal = 0

'nSomaTotal = CDbl(lblRegTot.Caption)
nContaReg = Val(lblNumReg.Caption)

With grdParc
    For x = 1 To .Rows - 1
        bDif = IIf(.TextMatrix(x, 15) = 0, False, True)
        bDup = IIf(.TextMatrix(x, 11) = "Não", False, True)
        nLinha = .TextMatrix(x, 16)
        
        If bDup Then
            sStatus = "DUPLICADO"
        Else
            If bDif Then
                sStatus = "C/DIFERENÇA"
            Else
                sStatus = "NORMAL"
            End If
        End If
        On Error Resume Next
        RdoAux.Close
        On Error GoTo 0
        
'Sql = "INSERT BAIXATMP(COMPUTADOR,ARQUIVO,FUNCIONARIO,BANCO,NUMREG,VALORTOTAL,CODREMESSA,CODCONVENIO,"
'Sql = Sql & "DATAGERACAO,NUMSEQUENCIAL,LAYOUT,TR,CONTAPREF,DATAPAGTO,DATACREDITO,NUMDOC,CODREDUZ,"
'Sql = Sql & "ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,DATAVENCTO,VALORLANCADO,"
'Sql = Sql & "VALORJUROS,VALORMULTA,VALORCORRECAO,VALORCALCULADO,VALORDIF,VALORPAGO,VALORTARIFA,"
'Sql = Sql & "VALORPAGOREAL,SITUACAO,SEQUENCIA,VALORBANCO,REGBANCO) values('"
'Sql = Sql & Left$(Trim$(NomeDoComputador), 30) & "','" & Left$(grdArq.TextMatrix(grdArq.Row, 1), 30) & "','" & Left$(Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8), 30) & "','" & Left$(Trim$(lblBanco.Caption), 50) & "'," & lblNumReg.Caption & "," & Virg2Ponto(Mid(RemovePonto(lblRegTot.Caption), 4, Len(lblRegTot.Caption) - 2)) & ",'" & Left$(lblCR.Caption, 5) & "','" & Left$(lblCC.Caption, 20) & "','"
'Sql = Sql & Format(lblDG.Caption, "mm/dd/yyyy") & "'," & lblNS.Caption & "," & lblVL.Caption & ",'" & grdF.TextMatrix(1, 0) & "','" & Left$(Trim$(grdF.TextMatrix(1, 1)), 20) & "','" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "','" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "','" & nNumDoc & "-" & RetornaDVNumDoc(CStr(nNumDoc)) & "','"
'Sql = Sql & Format(.TextMatrix(x, 1), "0000000") & "'," & .TextMatrix(x, 0) & "," & Val(Left$(.TextMatrix(x, 2), 3)) & "," & .TextMatrix(x, 3) & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ",'"
'Sql = Sql & Format(.TextMatrix(x, 12), "mm/dd/yyyy") & "'," & Virg2Ponto(.TextMatrix(x, 6)) & "," & Virg2Ponto(.TextMatrix(x, 8)) & "," & Virg2Ponto(.TextMatrix(x, 7)) & "," & Virg2Ponto(.TextMatrix(x, 9)) & "," & 0 & ","
'Sql = Sql & Virg2Ponto(.TextMatrix(x, 15)) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) & "," & Virg2Ponto(.TextMatrix(x, 14)) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) & ",'" & Left$(sStatus, 30) & "'," & 0 & "," & Virg2Ponto(sTr(nSomaTotal)) & ","
'Sql = Sql & nContaReg & ")"
'cn.Execute Sql, rdExecDirect
        
        qd.Sql = "{ Call spGRAVABAIXATMP(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = Trim$(NomeDoComputador) 'COMPUTADOR
        qd(1) = grdArq.TextMatrix(grdArq.Row, 1) 'ARQUIVO
        'qd(2) = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8) 'USUARIO
        qd(2) = NomeDeLogin 'USUARIO
        qd(3) = Trim$(lblBanco.Caption) 'BANCO
        qd(4) = lblNumReg.Caption 'NUM REGS
        qd(5) = Virg2Ponto(Mid(RemovePonto(lblRegTot.Caption), 4, Len(lblRegTot.Caption) - 2)) 'TOTAL REG
        qd(6) = lblCR.Caption 'REMESSA
        qd(7) = lblCC.Caption 'CONVENIO
        qd(8) = Format(lblDG.Caption, "mm/dd/yyyy") 'DATA GERACAO
        qd(9) = lblNS.Caption 'NUM SEQ
        qd(10) = lblVL.Caption 'VERSAO LAYOUT
        qd(11) = grdF.TextMatrix(1, 0) 'TR
        qd(12) = Trim$(grdF.TextMatrix(1, 1)) 'CONTA PREFEITURA
        qd(13) = Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") 'DATA PAGTO
        qd(14) = Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") 'DATA CREDITO
        qd(15) = nNumDoc & "-" & RetornaDVNumDoc(CStr(nNumDoc)) 'NUM DOC
        qd(16) = Format(.TextMatrix(x, 1), "0000000")
        qd(17) = .TextMatrix(x, 0) 'EXERCICIO
        qd(18) = Val(Left$(.TextMatrix(x, 2), 3)) 'LANCAMENTO
        qd(19) = .TextMatrix(x, 3) 'SEQLANCAMENTO
        qd(20) = .TextMatrix(x, 4) 'PARCELA
        qd(21) = .TextMatrix(x, 5) 'COMPLEMENTO
        qd(22) = Format(.TextMatrix(x, 12), "mm/dd/yyyy") 'DATA VENCTO
        qd(23) = Virg2Ponto(.TextMatrix(x, 6)) 'VALOR LANCADO
        qd(24) = Virg2Ponto(.TextMatrix(x, 8)) 'VALOR JUROS
        qd(25) = Virg2Ponto(.TextMatrix(x, 7)) 'VALOR MULTA
        qd(26) = Virg2Ponto(.TextMatrix(x, 9)) 'VALOR CORRECAO
        qd(27) = 0 'VALOR CALCULADO
        qd(28) = Virg2Ponto(.TextMatrix(x, 15)) 'VALOR DIF
        qd(29) = Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) 'VALOR PAGO
        qd(30) = Virg2Ponto(.TextMatrix(x, 14)) 'VALOR TARIFA
        qd(31) = Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) 'VALOR PAGO REAL
        qd(32) = sStatus 'SITUACAO
        qd(33) = 0
        qd(34) = Virg2Ponto(sTr(nSomaTotal)) 'SOMA TOTAL DO BANCO
        qd(35) = nContaReg
        qd(36) = Null
        qd(37) = Null
        Set RdoAux = qd.OpenResultset(rdOpenForwardOnly)
    Next
End With

End Sub

Private Sub BaixaRegF()

Ocupado
Pb.Value = 0
lblPB.Caption = "Efetuando Baixa"
lblPB.Refresh
MontaResumo
Pb.Value = 0
Liberado
Pb.Value = 0
lblPB.Caption = "Pronto"
lblPB.Refresh
MostraRpt

End Sub

Private Sub MostraRpt()

'EXIBE RELATORIO
frmReport.ShowReport "BAIXATMPDA", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub EfetuaBaixa()
Dim x As Long
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqLanc As Integer
Dim nNumParc As Integer
Dim nCodTributo As Integer
Dim nCompl As Integer
Dim nStatus As Integer
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim dDataPag As Date
Dim dDataVencto As Date
Dim nSeqAdd As Integer
Dim nSomaDoc As Double
Dim sCodAgencia As String

'EFETUA AS BAIXAS
With grdPag
    For x = 1 To .Rows - 1
        Pb.Value = Abs(x * 100 / .Rows - 1)
        If .TextMatrix(x, 1) <> "N/A" And Val(Left$(Trim$(.TextMatrix(x, 24)), 2)) = 0 Then
             nCodReduz = .TextMatrix(x, 1)
             nAnoExercicio = .TextMatrix(x, 3)
             nCodLanc = .TextMatrix(x, 2)
             nSeqLanc = .TextMatrix(x, 4)
             nNumParc = .TextMatrix(x, 5)
             nCompl = .TextMatrix(x, 6)
             dDataPag = CDate(.TextMatrix(x, 18))
             dDataVencto = CDate(.TextMatrix(x, 21))
             If nNumParc = 13 Then
                If Val(.TextMatrix(x, 15)) = 0 Then
                    nStatus = 1 'UNICA SEM DIF
                Else
                    nStatus = 9 'UNICA COM DIF
                End If
             Else
                If Val(.TextMatrix(x, 15)) = 0 Then
                    nStatus = 2 'PAGO SEM DIF
                Else
                    nStatus = 7 'PAGO COM DIF
                End If
             End If
            'EFETUA BAIXA NA TABELA NUMDOCUMENTO
            nSomaDoc = 0
             For y = 1 To grdF.Rows - 1
                 If Left$(grdF.TextMatrix(y, 8), Len(grdF.TextMatrix(y, 8)) - 1) = .TextMatrix(x, 0) Then
                    nSomaDoc = grdF.TextMatrix(y, 10)
                    Exit For
                 End If
             Next
             'BUSCA AGENCIA
             Sql = "SELECT AGCONTAPREF FROM BANCO WHERE CODBANCO=" & Val(Left$(lblBanco.Caption, 3))
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
             With RdoAux2
                    sCodAgencia = !AGCONTAPREF
                   .Close
             End With
             Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " ,CODAGENCIA ='" & sCodAgencia & "' , VALORPAGO=" & Virg2Ponto(sTr(nSomaDoc))
             Sql = Sql & " WHERE NUMDOCUMENTO=" & Val(.TextMatrix(x, 0))
             cn.Execute Sql, rdExecDirect
            'EFETUA BAIXA NA TABELA DEBITOPARCELA
             Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " ,PAGTODUPLICADO=" & IIf(.TextMatrix(x, 17) = "S", 1, 0) & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
             Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
             Sql = Sql & "CODCOMPLEMENTO=" & nCompl
             cn.Execute Sql, rdExecDirect
            'SE FOR PARCELA UNICA EFETUA BAIXA EM TODAS AS PARCELAS AUTOMATICAMENTO
             If nNumParc = 0 Then
                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=1 ,PAGTODUPLICADO=" & IIf(.TextMatrix(x, 17) = "S", 1, 0) & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                Sql = Sql & " AND NUMPARCELA<>0"
                cn.Execute Sql, rdExecDirect
             End If
            'SE FOR DUPLICADO ATUALIZA O Nº DE VEZES
             If .TextMatrix(x, 17) = "S" Then
                Sql = "SELECT CODREDUZIDO,QTDEDUPLICADO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                Sql = Sql & " AND NUMPARCELA=" & nNumParc
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                       If IsNull(!QTDEDUPLICADO) Then
                          Sql = "UPDATE DEBITOPARCELA SET QTDEDUPLICADO=1 "
                       Else
                          Sql = "UPDATE DEBITOPARCELA SET QTDEDUPLICADO=QTDEDUPLICADO+1 "
                       End If
                       Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND "
                       Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                       Sql = Sql & " AND NUMPARCELA=" & nNumParc
                       cn.Execute Sql, rdExecDirect
                    End If
                End With
               'GRAVA DEBITO ADICIONAL
                Sql = "SELECT MAX(SEQADICIONAL) AS MAXIMO FROM DEBITOADICIONAL WHERE CODREDUZIDO=" & nCodReduz & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                Sql = Sql & " AND NUMPARCELA=" & nNumParc
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                     If IsNull(!MAXIMO) Then
                        nSeqAdd = 1
                     Else
                        If .RowCount = 0 Then
                           nSeqAdd = 1
                        Else
                           nSeqAdd = !MAXIMO + 1
                        End If
                     End If
                    .Close
                End With
                Sql = "INSERT DEBITOADICIONAL (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
                Sql = Sql & "CODCOMPLEMENTO,SEQADICIONAL,VALORADICIONAL,DATAPAGAMENTO,DATARECEBIMENTO) VALUES("
                Sql = Sql & nCodReduz & "," & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & ","
                Sql = Sql & nSeqAdd & "," & Virg2Ponto(CDbl(grdPag.TextMatrix(x, 14))) & ",'" & Format(dDataPag, "mm/dd/yyyy") & "','"
                Sql = Sql & Format(lblDG.Caption, "mm/dd/yyyy") & "')"
                cn.Execute Sql, rdExecDirect
             Else
                'EFETUA BAIXA NA TABELA DEBITOTRIBUTO
                 Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO "
                 Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                 Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                        Do Until .EOF
                            nValorLanc = !valortributo
                            nCodTributo = !CodTributo
                            nValorJuros = 0
                            nValorMulta = 0
                            nValorCorrecao = 0
                                           
                            Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=" & Virg2Ponto(sTr(nValorCorrecao)) & " ,VALORMULTA=" & Virg2Ponto(sTr(nValorMulta)) & " ,VALORJUROS=" & Virg2Ponto(sTr(nValorJuros)) & " ,DATAPAGAMENTO='" & Format(dDataPag, "mm/dd/yyyy") & "' ,DATARECEBIMENTO='" & Format(grdPag.TextMatrix(x, 19), "mm/dd/yyyy") & "',VALORPAGO=" & Virg2Ponto(CDbl(grdPag.TextMatrix(x, 14)))
                            Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                            Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=" & nCodTributo
                            cn.Execute Sql, rdExecDirect
                           .MoveNext
                        Loop
                       .Close
                 End With
             End If
        End If
    Next
End With

'GRAVA NA TABELA ARQUIVOBAIXA
Sql = "INSERT ARQUIVOBAIXA (NOMEARQUIVO,DATACREDITO) VALUES('" & grdArq.TextMatrix(grdArq.Row, 1) & "','" & Format(lblDG.Caption, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect

End Sub


Private Sub BaixaRegB()

Dim RdoAux As rdoResultset, Sql As String
Dim x As Integer
Dim nDist As Integer
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Integer
Dim nFace As Integer
Dim nCodReduz As Integer
Dim nBanco As Integer
Dim nAgencia As Integer
Dim dDataOpcao As Date
Dim nCodMov As Integer
Dim nConta As Long

Pb.Value = 0
'REGISTRO B (CADASTRO DE OPTANTES)
With grdB
    If .Rows > 1 Then
       For x = 1 To .Rows - 1
            CallPb CLng(x), grdB.Rows - 1
            nDist = Val(.TextMatrix(x, 1))
            nSetor = Val(.TextMatrix(x, 2))
            nQuadra = Val(.TextMatrix(x, 3))
            nLote = Val(.TextMatrix(x, 4))
            nFace = Val(.TextMatrix(x, 5))
            nBanco = Val(Left$(lblBanco.Caption, 3))
            nAgencia = Val(.TextMatrix(x, 6))
            nConta = Val(.TextMatrix(x, 7))
            dDataOpcao = CDate(.TextMatrix(x, 8))
            nCodMov = Val(.TextMatrix(x, 9))
           'PROCURA CODIGO REDUZIDO
            Sql = "SELECT CODREDUZIDO FROM CADIMOB WHERE DISTRITO=" & nDist & " AND SETOR=" & nSetor & " AND "
            Sql = Sql & "QUADRA=" & nQuadra & " AND LOTE=" & nLote & " AND SEQ=" & nFace
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    nCodReduz = !CODREDUZIDO
                End If
               .Close
            End With
           'EFETUA CADASTRO
            If nCodReduz > 0 Then
                If nCodMov = 1 Then 'EXCLUSAO
                   Sql = "DELETE FROM DEBITOAUTOMATICO WHERE CODREDUZ= " & nCodReduz & " AND CODBANCO=" & nBanco
                   Sql = Sql & " AND CODAGENCIA=" & nAgencia & " AND NUMEROCONTA=" & nConta
                   cn.Execute Sql, rdExecDirect
                Else 'INCLUSAO
                   Sql = "SELECT CODREDUZ FROM DEBITOAUTOMATICO WHERE CODREDUZ= " & nCodReduz & " AND CODBANCO=" & nBanco
                   Sql = Sql & " AND CODAGENCIA=" & nAgencia & " AND NUMEROCONTA=" & nConta
                   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                   With RdoAux
                        If .RowCount = 0 Then 'VERIFICA DUPLICADOS
                            Sql = "INSERT DEBITOAUTOMATICO(CODREDUZ,CODBANCO,CODAGENCIA,NUMEROCONTA,DATAOPCAO) VALUES("
                            Sql = Sql & nCodReduz & "," & nBanco & "," & nAgencia & "," & nConta & ",'" & Format(dDataOpcao, "mm/dd/yyyy") & "')"
                            cn.Execute Sql, rdExecDirect
                           .Close
                        End If
                   End With
                End If
            End If
       Next
    End If
End With
Pb.Value = 0
End Sub

Private Sub cmdCancel_Click()
Dim RdoAux As rdoResultset, Sql As String

If grdF.Rows = 1 Then
    MsgBox "Não existem registros a reativar.", vbExclamation, "Atenção"
    Exit Sub
End If
Sql = "SELECT NOMEARQ,DATACREDITO,DATABAIXA FROM ARQUIVOBANCO WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If IsNull(!DATABAIXA) Then
           MsgBox "Não foi efetuado Baixa neste arquivo.", vbCritical, "Atenção"
          .Close
           Exit Sub
        End If
    End If
   .Close
End With

If MsgBox("Deseja REATIVAR os pagamentos deste arquivo ?", vbQuestion + vbYesNo, "CONFIRMAÇÃO DE REATIVAÇÃO") = vbYes Then
    Ocupado
    Pb.Value = 0
    lblPB.Caption = "Reativando Parcelas"
    lblPB.Refresh
    Reativa
    Pb.Value = 0
    lblPB.Caption = "Parcelas foram Canceladas"
    lblPB.Refresh
    Liberado
End If

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdShow_Click()
If grdArq.Rows > 1 Then
   x = Shell("NOTEPAD" & " " & grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1), vbNormalFocus)
End If

End Sub

Private Sub Form_Activate()

If grdArq.Rows > 1 And grdF.Rows = 1 Then
    grdArq.Row = grdArq.Rows - 1
    LeArquivo
End If

End Sub

Private Sub Form_Load()

frmMdi.AddWindow Me.Name, Me.Caption
'ImgLogo.Picture = frmDebAutomatico.cmdBanco(frmDebAutomatico.lblAux.Caption).PictureNormal
grdArq.ColWidth(0) = 0
grdF.Rows = 1
End Sub

Private Sub LeArquivoOld()

On Error GoTo Erro

Dim sFullPath As String
Dim Header As FebrabanA
Dim Registro As ControleFebraban
Dim RegistroB As FebrabanB
Dim RegistroF As FebrabanF
Dim RegistroH As FebrabanH
Dim RegistroJ As FebrabanJ
Dim RegistroT As FebrabanT
Dim RegistroX As FebrabanX
Dim Footer As FebrabanZ
Dim sRetorno As String
Dim Posicao  As Long
Dim nCountB As Integer
Dim nCountF As Integer
Dim nCountH As Integer
Dim nCountJ As Integer
Dim nCountT As Integer
Dim nCountX As Integer


ReDim aFebrabanA(0): ReDim aFebrabanB(0): ReDim aFebrabanF(0): ReDim aFebrabanH(0): ReDim aFebrabanJ(0): ReDim aFebrabanZ(0)
ReDim aFebrabanT(0): ReDim aFebrabanX(0): ReDim aFebrabanG(0)

grdF.Rows = 1: grdH.Rows = 1: grdJ.Rows = 1: grdT.Rows = 1: grdX.Rows = 1
nCountB = 0: nCountF = 0: nCountH = 0: nCountJ = 0: nCountT = 0: nCountX = 0
 
sFullPath = grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1)

Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header
    aFebrabanA(0).CodigoRegistro = Header.CodigoRegistro
    aFebrabanA(0).CodigoRemessa = Header.CodigoRemessa
    aFebrabanA(0).CodigoConvenio = Header.CodigoConvenio
    aFebrabanA(0).NomeEmpresa = Header.NomeEmpresa
    aFebrabanA(0).CodigoBanco = Header.CodigoBanco
    aFebrabanA(0).NomeBanco = Header.NomeBanco
    aFebrabanA(0).DataGeracao = Header.DataGeracao
    aFebrabanA(0).NumeroSeq = Header.NumeroSeq
    aFebrabanA(0).VersaoLayout = Header.VersaoLayout
    aFebrabanA(0).Filler = Header.Filler
    lblBanco.Caption = aFebrabanA(0).CodigoBanco & " - " & aFebrabanA(0).NomeBanco
    lblCR.Caption = aFebrabanA(0).CodigoRegistro
    lblCC.Caption = Val(aFebrabanA(0).CodigoConvenio)
    lblDG.Caption = ConvDataSerial(aFebrabanA(0).DataGeracao)
    lblNS.Caption = aFebrabanA(0).NumeroSeq
    lblVL.Caption = aFebrabanA(0).VersaoLayout
    Posicao = Len(Header) + 3
    Do While Not EOF(1)
         Get #1, Posicao, Registro
         If Left$(Registro.Registro, 1) = "B" Then
              Get #1, Posicao, RegistroB
              aFebrabanB(nCountB).CodigoRegistro = RegistroB.CodigoRegistro
              aFebrabanB(nCountB).Distrito = RegistroB.Distrito
              aFebrabanB(nCountB).Setor = RegistroB.Setor
              aFebrabanB(nCountB).Quadra = RegistroB.Quadra
              aFebrabanB(nCountB).Lote = RegistroB.Lote
              aFebrabanB(nCountB).Seq = RegistroB.Seq
              aFebrabanB(nCountB).FillerID = RegistroB.FillerID
              aFebrabanB(nCountB).CodAgencia = RegistroB.CodAgencia
              aFebrabanB(nCountB).ContaCliente = RegistroB.ContaCliente
              aFebrabanB(nCountB).DataOpcao = RegistroB.DataOpcao
              aFebrabanB(nCountB).Filler = RegistroB.Filler
              aFebrabanB(nCountB).CodMovimento = RegistroB.CodMovimento
              With aFebrabanB(nCountB)
                    grdB.AddItem .CodigoRegistro & Chr(9) & .Distrito & Chr(9) & .Setor & Chr(9) & .Quadra & Chr(9) & .Lote & Chr(9) & .Seq & Chr(9) & .CodAgencia & _
                    Chr(9) & .ContaCliente & Chr(9) & ConvDataSerial(.DataOpcao) & Chr(9) & IIf(.CodMovimento = 1, "1 - Exclusão", "2 - Inclusão")
              End With
              nCountB = nCountB + 1
              ReDim Preserve aFebrabanB(nCountB)
         ElseIf Left$(Registro.Registro, 1) = "F" Then
              Get #1, Posicao, RegistroF
              aFebrabanF(nCountF).CodigoRegistro = RegistroF.CodigoRegistro
              aFebrabanF(nCountF).Distrito = RegistroF.Distrito
              aFebrabanF(nCountF).Setor = RegistroF.Setor
              aFebrabanF(nCountF).Quadra = RegistroF.Quadra
              aFebrabanF(nCountF).Lote = RegistroF.Lote
              aFebrabanF(nCountF).Seq = RegistroF.Seq
              aFebrabanF(nCountF).CodAgencia = RegistroF.CodAgencia
              aFebrabanF(nCountF).ContaCliente = RegistroF.ContaCliente
              aFebrabanF(nCountF).DataVencto = RegistroF.DataVencto
              aFebrabanF(nCountF).ValorDebito = RegistroF.ValorDebito
              aFebrabanF(nCountF).CodRetorno = RegistroF.CodRetorno
              aFebrabanF(nCountF).NumDoc = RegistroF.NumDoc
              aFebrabanF(nCountF).CodMovimento = RegistroF.CodMovimento
              With aFebrabanF(nCountF)
                    Select Case .CodRetorno
                            Case "00"
                                    sRetorno = "Débito Efetuado"
                            Case "01"
                                    sRetorno = "Insuficiência de Fundos"
                            Case "02"
                                    sRetorno = "Conta Corrente não Cadastrada"
                            Case "04"
                                    sRetorno = "Outras Restrições"
                            Case "10"
                                    sRetorno = "Agência em Regime de Encerramento"
                            Case "12"
                                    sRetorno = "Valor Inválido"
                            Case "13"
                                    sRetorno = "Data de Lançamento inválida"
                            Case "14"
                                    sRetorno = "Agência Inválida"
                            Case "15"
                                    sRetorno = "DAC da conta corrente inválido"
                            Case "18"
                                    sRetorno = "Data do Débito anterior ao do processamento"
                            Case "30"
                                    sRetorno = "Sem contrato de débito automático"
                            Case "96"
                                    sRetorno = "Manutenção do Cadastro"
                            Case "97"
                                    sRetorno = "Cancelamento - Não Encontrado"
                            Case "98"
                                    sRetorno = "Cancelamento - não efetuado, fora de tempo habil"
                            Case "99"
                                    sRetorno = "Cancelamento - cancelado conforme solicitado"
                            Case Else
                                   sRetorno = "Erro Indefinido"
                    End Select
                    grdF.AddItem .CodigoRegistro & Chr(9) & .Distrito & Chr(9) & .Setor & Chr(9) & .Quadra & Chr(9) & .Lote & Chr(9) & .Seq & Chr(9) & .CodAgencia & _
                    Chr(9) & .ContaCliente & Chr(9) & Format(.NumDoc, "000000000") & Chr(9) & ConvDataSerial(.DataVencto) & Chr(9) & Format(.ValorDebito / 100, "#0.00") & Chr(9) & .CodRetorno & " - " & sRetorno & Chr(9) & .CodMovimento
              End With
              nCountF = nCountF + 1
              ReDim Preserve aFebrabanF(nCountF)
         ElseIf Left$(Registro.Registro, 1) = "H" Then
              Get #1, Posicao, RegistroH
              aFebrabanH(nCountH).CodigoRegistro = RegistroH.CodigoRegistro
              aFebrabanH(nCountH).DistritoOld = RegistroH.DistritoOld
              aFebrabanH(nCountH).SetorOld = RegistroH.SetorOld
              aFebrabanH(nCountH).QuadraOld = RegistroH.QuadraOld
              aFebrabanH(nCountH).LoteOld = RegistroH.LoteOld
              aFebrabanH(nCountH).SeqOld = RegistroH.SeqOld
              aFebrabanH(nCountH).CodAgencia = RegistroH.CodAgencia
              aFebrabanH(nCountH).ContaCliente = RegistroH.ContaCliente
              aFebrabanH(nCountH).Distrito = RegistroH.Distrito
              aFebrabanH(nCountH).Setor = RegistroH.Setor
              aFebrabanH(nCountH).Quadra = RegistroH.Quadra
              aFebrabanH(nCountH).Lote = RegistroH.Lote
              aFebrabanH(nCountH).Seq = RegistroH.Seq
              aFebrabanH(nCountH).Ocorrencia = RegistroH.Ocorrencia
              aFebrabanH(nCountH).CodMovimento = RegistroH.CodMovimento
              With aFebrabanH(nCountH)
                    grdH.AddItem .CodigoRegistro & Chr(9) & .DistritoOld & Chr(9) & .SetorOld & Chr(9) & .QuadraOld & Chr(9) & .LoteOld & Chr(9) & .SeqOld & Chr(9) & .CodAgencia & _
                    Chr(9) & .ContaCliente & Chr(9) & .Distrito & Chr(9) & .Setor & Chr(9) & .Quadra & Chr(9) & .Lote & Chr(9) & .Seq & Chr(9) & .Ocorrencia & Chr(9) & .CodMovimento
              End With
              nCountH = nCountH + 1
              ReDim Preserve aFebrabanH(nCountH)
         ElseIf Left$(Registro.Registro, 1) = "J" Then
              Get #1, Posicao, RegistroJ
              aFebrabanJ(nCountJ).CodigoRegistro = RegistroJ.CodigoRegistro
              aFebrabanJ(nCountJ).CodigoNSA = RegistroJ.CodigoNSA
              aFebrabanJ(nCountJ).DataGeracao = RegistroJ.DataGeracao
              aFebrabanJ(nCountJ).TotalRegistros = RegistroJ.TotalRegistros
              aFebrabanJ(nCountJ).ValorTotal = RegistroJ.ValorTotal
              aFebrabanJ(nCountJ).DataProcessamento = RegistroJ.DataProcessamento
              With aFebrabanJ(nCountJ)
                    grdJ.AddItem .CodigoRegistro & Chr(9) & .CodigoNSA & Chr(9) & ConvDataSerial(.DataGeracao) & Chr(9) & .TotalRegistros & Chr(9) & Format(.ValorTotal / 100, "#0.00") & Chr(9) & ConvDataSerial(.DataProcessamento)
              End With
              nCountJ = nCountJ + 1
              ReDim Preserve aFebrabanJ(nCountJ)
         ElseIf Left$(Registro.Registro, 1) = "T" Then
              Get #1, Posicao, RegistroT
              aFebrabanT(nCountT).CodigoRegistro = RegistroT.CodigoRegistro
              aFebrabanT(nCountT).TotalRegistros = RegistroT.TotalRegistros
              aFebrabanT(nCountT).ValorTotal = RegistroT.ValorTotal
              With aFebrabanT(nCountJ)
                    grdT.AddItem .CodigoRegistro & Chr(9) & .TotalRegistros & Chr(9) & Format(.ValorTotal / 100, "#0.00")
              End With
              nCountT = nCountT + 1
              ReDim Preserve aFebrabanT(nCountT)
         ElseIf Left$(Registro.Registro, 1) = "X" Then
              Get #1, Posicao, RegistroX
              aFebrabanX(nCountX).CodigoRegistro = RegistroX.CodigoRegistro
              aFebrabanX(nCountX).CodigoAgencia = RegistroX.CodigoAgencia
              aFebrabanX(nCountX).NomeAgencia = RegistroX.NomeAgencia
              aFebrabanX(nCountX).EnderecoAgencia = RegistroX.EnderecoAgencia
              aFebrabanX(nCountX).NumAgencia = RegistroX.NumAgencia
              aFebrabanX(nCountX).CepAgencia = RegistroX.CepAgencia
              aFebrabanX(nCountX).SufixoCep = RegistroX.SufixoCep
              aFebrabanX(nCountX).Cidade = RegistroX.Cidade
              aFebrabanX(nCountX).SiglaUF = RegistroX.SiglaUF
              aFebrabanX(nCountX).SituacaoAgencia = RegistroX.SituacaoAgencia
              With aFebrabanX(nCountX)
                    grdX.AddItem .CodigoRegistro & Chr(9) & .CodigoAgencia & Chr(9) & .NomeAgencia & Chr(9) & .EnderecoAgencia & Chr(9) & .NumAgencia & _
                    Chr(9) & .CepAgencia & Chr(9) & .SufixoCep & Chr(9) & .Cidade & Chr(9) & .SiglaUF & Chr(9) & .SituacaoAgencia
              End With
              nCountX = nCountX + 1
              ReDim Preserve aFebrabanX(nCountX)
         Else
              Get #1, Posicao, Footer
              aFebrabanZ(0).CodigoRegistro = Footer.CodigoRegistro
              aFebrabanZ(0).TotalRegistro = Footer.TotalRegistro
              aFebrabanZ(0).ValorTotal = Footer.ValorTotal
              aFebrabanZ(0).Filler = Footer.Filler
              lblNumReg.Caption = Val(aFebrabanZ(0).TotalRegistro) - 2
              lblRegTot.Caption = "R$ " & Format(aFebrabanZ(0).ValorTotal / 100, "#0.00")
              Exit Do
         End If
         Posicao = Posicao + Len(Registro) + 2
    Loop
Close #1


 Exit Sub
Erro:
 MsgBox Err.Description
 Resume Next
End Sub

Private Sub LeArquivo()

On Error GoTo Erro

Dim sFullPath As String
Dim Header As FebrabanA
Dim Registro As ControleFebraban
Dim RegistroF As FebrabanF
Dim Footer As FebrabanZ
Dim sRetorno As String
Dim Posicao  As Long
Dim nCountF As Integer
Dim nValorEfetuado As Double
ReDim aFebrabanA(0): ReDim aFebrabanF(0): ReDim aFebrabanZ(0)

grdF.Rows = 1
nCountF = 0
 
sFullPath = grdArq.TextMatrix(grdArq.Row, 0) & grdArq.TextMatrix(grdArq.Row, 1)

Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header
    aFebrabanA(0).CodigoRegistro = Header.CodigoRegistro
    aFebrabanA(0).CodigoRemessa = Header.CodigoRemessa
    aFebrabanA(0).CodigoConvenio = Header.CodigoConvenio
    aFebrabanA(0).NomeEmpresa = Header.NomeEmpresa
    aFebrabanA(0).CodigoBanco = Header.CodigoBanco
    aFebrabanA(0).NomeBanco = Header.NomeBanco
    aFebrabanA(0).DataGeracao = Header.DataGeracao
    aFebrabanA(0).NumeroSeq = Header.NumeroSeq
    aFebrabanA(0).VersaoLayout = Header.VersaoLayout
    aFebrabanA(0).Filler = Header.Filler
    lblBanco.Caption = aFebrabanA(0).CodigoBanco & " - " & aFebrabanA(0).NomeBanco
    lblCR.Caption = aFebrabanA(0).CodigoRegistro
    lblCC.Caption = Val(aFebrabanA(0).CodigoConvenio)
    lblDG.Caption = ConvDataSerial(aFebrabanA(0).DataGeracao)
    lblNS.Caption = aFebrabanA(0).NumeroSeq
    lblVL.Caption = aFebrabanA(0).VersaoLayout
    Posicao = Len(Header) + 3
    Do While Not EOF(1)
         Get #1, Posicao, Registro
         If Left$(Registro.Registro, 1) = "F" Then
              Get #1, Posicao, RegistroF
              aFebrabanF(nCountF).CodigoRegistro = RegistroF.CodigoRegistro
              aFebrabanF(nCountF).Distrito = RegistroF.Distrito
              aFebrabanF(nCountF).Setor = RegistroF.Setor
              aFebrabanF(nCountF).Quadra = RegistroF.Quadra
              aFebrabanF(nCountF).Lote = RegistroF.Lote
              aFebrabanF(nCountF).Seq = RegistroF.Seq
              aFebrabanF(nCountF).CodAgencia = RegistroF.CodAgencia
              aFebrabanF(nCountF).ContaCliente = RegistroF.ContaCliente
              aFebrabanF(nCountF).DataVencto = RegistroF.DataVencto
              aFebrabanF(nCountF).ValorDebito = RegistroF.ValorDebito
              aFebrabanF(nCountF).CodRetorno = RegistroF.CodRetorno
              aFebrabanF(nCountF).NumDoc = RegistroF.NumDoc
              aFebrabanF(nCountF).CodMovimento = RegistroF.CodMovimento
              With aFebrabanF(nCountF)
                    Select Case .CodRetorno
                            Case "00"
                                    sRetorno = "Débito Efetuado"
                            Case "01"
                                    sRetorno = "Insuficiência de Fundos"
                            Case "02"
                                    sRetorno = "Conta Corrente não Cadastrada"
                            Case "04"
                                    sRetorno = "Outras Restrições"
                            Case "10"
                                    sRetorno = "Agência em Regime de Encerramento"
                            Case "12"
                                    sRetorno = "Valor Inválido"
                            Case "13"
                                    sRetorno = "Data de Lançamento inválida"
                            Case "14"
                                    sRetorno = "Agência Inválida"
                            Case "15"
                                    sRetorno = "DAC da conta corrente inválido"
                            Case "18"
                                    sRetorno = "Data do Débito anterior ao do processamento"
                            Case "30"
                                    sRetorno = "Sem contrato de débito automático"
                            Case "96"
                                    sRetorno = "Manutenção do Cadastro"
                            Case "97"
                                    sRetorno = "Cancelamento - Não Encontrado"
                            Case "98"
                                    sRetorno = "Cancelamento - não efetuado, fora de tempo habil"
                            Case "99"
                                    sRetorno = "Cancelamento - cancelado conforme solicitado"
                            Case Else
                                   sRetorno = "Erro Indefinido"
                    End Select
                    grdF.AddItem .CodigoRegistro & Chr(9) & .Distrito & Chr(9) & .Setor & Chr(9) & .Quadra & Chr(9) & .Lote & Chr(9) & .Seq & Chr(9) & .CodAgencia & _
                    Chr(9) & .ContaCliente & Chr(9) & ConvDataSerial(.DataVencto) & Chr(9) & Format(.ValorDebito / 100, "#0.00") & Chr(9) & .CodRetorno & " - " & sRetorno & Chr(9) & Format(.NumDoc, "000000")
              End With
              nCountF = nCountF + 1
              ReDim Preserve aFebrabanF(nCountF)
'              If nCountF = 13 Then MsgBox "aqui"
         ElseIf Left$(Registro.Registro, 1) = "Z" Then
              Get #1, Posicao, Footer
              aFebrabanZ(0).CodigoRegistro = Footer.CodigoRegistro
              aFebrabanZ(0).TotalRegistro = Footer.TotalRegistro
              aFebrabanZ(0).ValorTotal = Footer.ValorTotal
              aFebrabanZ(0).Filler = Footer.Filler
              lblNumReg.Caption = Val(aFebrabanZ(0).TotalRegistro) - 2
              lblRegTot.Caption = "R$ " & FormatNumber(aFebrabanZ(0).ValorTotal / 100, 2)
              Exit Do
         ElseIf Left$(Registro.Registro, 1) = "X" Then
         ElseIf Left$(Registro.Registro, 1) = "B" Then
         Else
                MsgBox "Erro de Leitura no Arquivo linha: " & nCountF, vbCritical
                Exit Do
         End If
         Posicao = Posicao + Len(Registro) + 2
    Loop
 Close #1

nValorEfetuado = 0
For x = 1 To grdF.Rows - 1
    If Left$(grdF.TextMatrix(x, 10), 2) = "00" Then
       nValorEfetuado = nValorEfetuado + CDbl(grdF.TextMatrix(x, 9))
    End If
Next
lblEfetuado.Caption = "R$ " & FormatNumber(nValorEfetuado, 2)

 Exit Sub
Erro:
 MsgBox Err.Description
 Resume Next
End Sub

Private Function ConvDataSerial(sData As String) As String
ConvDataSerial = Right$(sData, 2) & "/" & Mid(sData, 5, 2) & "/" & Left$(sData, 4)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
End Sub

Private Sub grdArq_Click()

If grdArq.Row > 0 Then
     LeArquivo
     On Error Resume Next
     grdF.SetFocus
End If
End Sub


Private Sub MontaResumo()
Dim x As Long, z As Long, y As Integer

Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqLanc As Integer
Dim nNumParc As Integer
Dim nCodTributo As Integer
Dim nCompl As Integer
Dim nStatus As Integer
Dim nValorLanc As Double
Dim nValorJuros As Double
Dim nValorMulta As Double, nValorDif As Double
Dim nValorCorrecao As Double, nValortotal As Double
Dim nValorPago As Double, nValorPagoSTaxa As Double
Dim nValorTaxa As Double
Dim nSomaJuros As Double, nSomaMulta As Double, nSomaCorrecao As Double
Dim bDupl As Boolean
Dim dDataPag As Date, dDataCred As Date
Dim dDataVencto As Date
Dim nValorPrincipal As Double
Dim nSomaTotal As Double, nSomaTotal2 As Double
Dim nSomaPrincipal As Double
Dim bDupS As Boolean, bDupN As Boolean, nSomaClass As Double, nSomaClass2 As Double
Dim bTemIssVar As Boolean, nValorSIssVar As Double
Dim nCodBanco As Integer, sCodAgencia As String
Dim nValorPagoReal As Double, nResto As Double, nContaResto As Integer, bDebClassificar As Boolean

grdParc.Rows = 1
nSomaClass = 0: nSomaClass2 = 0
bDebClassificar = False
For x = 1 To grdF.Rows - 1
    CallPb x, grdF.Rows - 1
    grdParc.Rows = 1
            
    If Val(Left$(grdF.TextMatrix(x, 10), 2)) > 0 Then GoTo proximo
    
    nCodReduz = Val(grdF.TextMatrix(x, 11))
    'If nCodReduz = 6002 Then MsgBox "aqui"
'    If Val(grdF.TextMatrix(x, 11)) = 9595 Then MsgBox "A"
    dDataPag = Format(grdF.TextMatrix(x, 8), "dd/mm/yyyy")
    dDataCred = Format(grdF.TextMatrix(x, 8), "dd/mm/yyyy")
 '   nSomaTotal2 = CDbl(lblRegTot.Caption)
   
    nValorPago = CDbl(grdF.TextMatrix(x, 9))
    nResto = nValorPago
    nCodBanco = Val(Left$(lblBanco.Caption, 3))
    If nCodBanco = 0 Then nCodBanco = Val(Left$(lblBanco.Caption, 3))
    sCodAgencia = grdF.TextMatrix(x, 6)
    nSomaPrincipal = 0
    
    'CARREGA OS LANÇAMENTOS DO DOCUMENTO
    Sql = "SELECT LANCAMENTO.DESCREDUZ, DEBITOPARCELA.STATUSLANC, SITUACAOLANCAMENTO.DESCSITUACAO, DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE, DEBITOPARCELA.CODREDUZIDO, DEBITOPARCELA.ANOEXERCICIO, DEBITOPARCELA.CODLANCAMENTO,"
    Sql = Sql & "DEBITOPARCELA.SeqLancamento , DEBITOPARCELA.NumParcela, DEBITOPARCELA.CODCOMPLEMENTO FROM LANCAMENTO INNER JOIN DEBITOPARCELA ON LANCAMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO INNER JOIN "
    Sql = Sql & "SITUACAOLANCAMENTO ON DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO "
    Sql = Sql & "WHERE (DEBITOPARCELA.SEQLANCAMENTO < 100) AND (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.CODLANCAMENTO = 1) AND "
    Sql = Sql & "(DEBITOPARCELA.NUMPARCELA > 0) AND (DEBITOPARCELA.DATAVENCIMENTO = '" & Format(grdF.TextMatrix(x, 8), "mm/dd/yyyy") & "')"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "Não localizado lancamentos para o código " & nCodReduz
            'DOCUMENTO NÃO ENCONTRADO (VERIFICAR ...........)
             nSomaClass = nSomaClass + nValorPago
             nResto = nResto - nValorPago
             Sql = "SELECT * FROM RECEITACLASSIFICAR WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATARECEITA='"
             Sql = Sql & Format(dDataCred, "mm/dd/yyyy") & "' AND NUMDOCUMENTO=" & nNumDoc
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                If .RowCount = 0 Then
                    Sql = "INSERT RECEITACLASSIFICAR (NOMEARQ,DATARECEITA,CODBANCO,NUMDOCUMENTO,VALORTOTAL) VALUES('"
                    Sql = Sql & grdArq.TextMatrix(grdArq.Row, 1) & "','" & Format(dDataCred, "mm/dd/yyyy") & "'," & nCodBanco & ","
                    Sql = Sql & nNumDoc & "," & Virg2Ponto(CStr(nValorPago)) & ")"
                Else
                    Sql = "UPDATE RECEITACLASSIFICAR SET VALORTOTAL=VALORTOTAL + " & Virg2Ponto(CStr(nValorPago)) & "  WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATARECEITA='"
                    Sql = Sql & Format(dDataCred, "mm/dd/yyyy") & "' AND NUMDOCUMENTO=" & nNumDoc
                End If
                cn.Execute Sql, rdExecDirect
               .Close
             End With
             GoTo proximo
        Else
             Sql = "SELECT NUMDOCUMENTO.NUMDOCUMENTO,VALORTAXADOC FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO "
             Sql = Sql & "Where CODREDUZIDO = " & !CODREDUZIDO & " And AnoExercicio = " & !AnoExercicio & " AND CodLancamento = " & !CodLancamento & " AND "
             Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                If .RowCount > 0 Then
                    nValorTaxa = FormatNumber(!VALORTAXADOC, 2)
                    nNumDoc = !NumDocumento
                    Sql = "DELETE FROM DEBITOCLASSIFICAR WHERE DATARECEITA='" & Format(dDataCred, "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND NUMDOCUMENTO=" & nNumDoc
                    cn.Execute Sql, rdExecDirect
                Else
                    nValorTaxa = 0
                End If
               .Close
             End With
        End If
        
        nValorPagoSTaxa = nValorPago - nValorTaxa
        nContaResto = 1
        Do Until .EOF '(RDOAUX)
             If IsNull(!statuslanc) Then
                Sql = "INSERT DEBITOCLASSIFICAR (DATARECEITA,CODBANCO,NOMEARQ,NUMDOCUMENTO,CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO) VALUES('"
                Sql = Sql & Format(dDataCred, "mm/dd/yyyy") & "'," & nCodBanco & ",'" & grdArq.TextMatrix(grdArq.Row, 1) & "'," & nNumDoc & "," & nCodReduz & "," & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & ")"
                cn.Execute Sql, rdExecDirect
                bDebClassificar = True
                GoTo proximo
             End If
             nStatus = !statuslanc
             dDataVencto = !DataVencimento
             If nStatus = 1 Or nStatus = 2 Or nStatus = 7 Or nStatus = 9 Then
                bDupS = True
                bDupl = True
             Else
                bDupN = True
                bDupl = False
             End If
            'ADICIONA NO GRID PARCELA
             grdParc.AddItem !AnoExercicio & Chr(9) & Format(!CODREDUZIDO, "000000") & Chr(9) & Format(!CodLancamento, "000") & " - " & !descreduz & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & Format(!NumParcela, "00") & Chr(9) & _
                 !CODCOMPLEMENTO & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & _
                 "-" & Chr(9) & "-" & Chr(9) & "-" & Chr(9) & IIf(bDupl, "Sim", "Não") & Chr(9) & Format(dDataVencto, "dd/mm/yyyy") & Chr(9) & nValorPago & Chr(9) & nValorTaxa & Chr(9) & "-" & Chr(9) & x
            
            'PARA CADA LANCAMENTO CARREGAMOS OS TRIBUTOS
             Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO "
             Sql = Sql & "WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
             Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3 "
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                nSomaPrincipal = 0
                nSomaJuros = 0
                nSomaMulta = 0
                nSomaCorrecao = 0
                nSomaTotal = 0
                grdTrib.Rows = 1
                bTemIssVar = False
                Do Until .EOF
                   If !CodTributo = 13 Then
                     'CALCULA ISS VARIAVEL
                      Sql = "SELECT SUM(VALORTRIBUTO) AS TOTAL FROM DEBITOTRIBUTO "
                      Sql = Sql & "WHERE CODREDUZIDO = " & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO = " & RdoAux!AnoExercicio & " AND CODLANCAMENTO = " & RdoAux!CodLancamento & " AND "
                      Sql = Sql & "SEQLANCAMENTO = " & RdoAux!SeqLancamento & " AND NUMPARCELA = " & RdoAux!NumParcela & " AND CODCOMPLEMENTO = " & RdoAux!CODCOMPLEMENTO & " AND CODTRIBUTO<>3 "
                      Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                      With RdoAux2
                          If Not IsNull(!Total) Then
                             nValorSIssVar = !Total
                          Else
                            nValorSIssVar = 0
                         End If
                        .Close
                      End With
                     nValorLanc = FormatNumber(nValorPago - nValorSIssVar, 2)
                   Else
                      nValorLanc = FormatNumber(!valortributo, 2)
                   End If
                   
                  'ADICIONA NO GRID TRIBUTO
                  If (dDataPag > dDataVencto) Then 'PAGO APOS O VENCIMENTO
                      nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
                      nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                      nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                  Else
                      nValorCorrecao = 0
                      nValorJuros = 0
                      nValorMulta = 0
                  End If
                  nValortotal = nValorLanc + nValorCorrecao + nValorJuros + nValorMulta
                  nSomaJuros = nSomaJuros + nValorJuros
                  nSomaMulta = nSomaMulta + nValorMulta
                  nSomaCorrecao = nSomaCorrecao + nValorCorrecao
                  nSomaTotal = nSomaTotal + nValortotal
                  grdTrib.AddItem !CodTributo & Chr(9) & nValorLanc & Chr(9) & nValorMulta & Chr(9) & nValorJuros & Chr(9) & nValorCorrecao & Chr(9) & nValortotal & Chr(9) & x
                  nSomaPrincipal = nSomaPrincipal + nValorLanc
                 .MoveNext
               Loop
            End With
            
           'ATUALIZA GRID PARCELA
            nValorDif = Round(nValorPago - (nValorTaxa + nSomaTotal), 2)
            grdParc.TextMatrix(grdParc.Rows - 1, 6) = nSomaPrincipal
            grdParc.TextMatrix(grdParc.Rows - 1, 7) = nSomaMulta
            grdParc.TextMatrix(grdParc.Rows - 1, 8) = nSomaJuros
            grdParc.TextMatrix(grdParc.Rows - 1, 9) = nSomaCorrecao
            grdParc.TextMatrix(grdParc.Rows - 1, 10) = nSomaTotal
            grdParc.TextMatrix(grdParc.Rows - 1, 15) = nValorDif
            
            'CORRIGE VALOR PAGO QUANDO + DE 1 LANCAMENTO
            If grdParc.Rows > 2 Then
                For z = 1 To grdParc.Rows - 1
                    If grdParc.TextMatrix(z, 6) <> "N/A" Then
                        nValorPrincipal = CDbl(grdParc.TextMatrix(z, 6)) + CDbl(grdParc.TextMatrix(z, 14))
                        If nValorPago >= nValorPrincipal Then
                           grdParc.TextMatrix(z, 13) = FormatNumber(nValorPrincipal, 2)
                           grdParc.TextMatrix(z, 15) = FormatNumber(CDbl(grdParc.TextMatrix(z, 13)) - (CDbl(grdParc.TextMatrix(z, 6)) + CDbl(grdParc.TextMatrix(z, 14))), 2)
                        End If
                    End If
                Next
            End If
            
           'CARREGA DADOS PARA BAIXA DE PARCELA
            With grdParc
                nValorLanc = .TextMatrix(.Rows - 1, 13)
                nCodReduz = .TextMatrix(.Rows - 1, 1)
                nAnoExercicio = .TextMatrix(.Rows - 1, 0)
                nCodLanc = Val(Left$(.TextMatrix(.Rows - 1, 2), 3))
                nSeqLanc = .TextMatrix(.Rows - 1, 3)
                nNumParc = .TextMatrix(.Rows - 1, 4)
                nCompl = .TextMatrix(.Rows - 1, 5)
                If nNumParc = 0 Then
                   If CDbl(.TextMatrix(.Rows - 1, 15)) <= 0 Then
                       nStatus = 1 'UNICA SEM DIF
                   Else
                       nStatus = 9 'UNICA COM DIF
                   End If
                Else
                   If CDbl(.TextMatrix(.Rows - 1, 15)) <= 0 Then
                       nStatus = 2 'PAGO SEM DIF
                   Else
                       nStatus = 7 'PAGO COM DIF
                   End If
                End If

                If UCase$(.TextMatrix(.Rows - 1, 11)) <> "SIM" Then 'não é duplicado
                     'EFETUA BAIXA NA TABELA DEBITOPARCELA
                      Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
                      Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                      Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                      cn.Execute Sql, rdExecDirect
                     'SE FOR PARCELA UNICA EFETUA BAIXA EM TODAS AS PARCELAS AUTOMATICAMENTO
                     'SERA?
                      If nNumParc = 0 Then
                         Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=1  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                         Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                         Sql = Sql & " AND NUMPARCELA<>0"
                         cn.Execute Sql, rdExecDirect
                      Else
                         Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                         Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                         Sql = Sql & " AND NUMPARCELA=0"
                         cn.Execute Sql, rdExecDirect
                      End If
                     'EFETUA BAIXA NA TABELA DEBITOTRIBUTO
                      With grdTrib
                          For y = 1 To .Rows - 1
                              nCodTributo = .TextMatrix(y, 0)
                              nValorCorrecao = .TextMatrix(y, 4)
                              nValorJuros = .TextMatrix(y, 3)
                              nValorMulta = .TextMatrix(y, 2)
                              Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=" & Virg2Ponto(sTr(nValorCorrecao)) & " ,VALORMULTA=" & Virg2Ponto(sTr(nValorMulta)) & " ,VALORJUROS=" & Virg2Ponto(sTr(nValorJuros))
                              Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                              Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=" & nCodTributo
                              cn.Execute Sql, rdExecDirect
                          Next
                      End With
                 End If
                'EFETUA BAIXA NA TABELA DEBITOPAGO
                 Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND "
                 Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                 Sql = Sql & " AND NUMPARCELA=" & nNumParc
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                      If IsNull(!MAXIMO) Then
                         nSeqAdd = 0
                      Else
                         If .RowCount = 0 Then
                            nSeqAdd = 0
                        Else
                           nSeqAdd = !MAXIMO + 1
                        End If
                     End If
                    .Close
                End With
                If nContaResto = RdoAux.RowCount Then
                    nValorPagoReal = nResto
                    nResto = 0
                Else
                    If nResto >= nSomaTotal Then
                        nValorPagoReal = nSomaTotal
                        nResto = nResto - nSomaTotal
                    Else
                        nValorPagoReal = nResto
                        nResto = 0
                    End If
                End If
                nContaResto = nContaResto + 1
                nSomaClass2 = nSomaClass2 + nValorPagoReal
                grdF.TextMatrix(x, 12) = Virg2Ponto(sTr(nValorPagoReal))
                Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL) VALUES(" & nCodReduz & ","
                Sql = Sql & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nSeqAdd & ",'"
                Sql = Sql & Format(dDataPag, "mm/dd/yyyy") & "','" & Format(dDataCred, "mm/dd/yyyy") & "'," & Virg2Ponto(sTr(nSomaTotal)) & ","
                Sql = Sql & nCodBanco & ",'" & sCodAgencia & "'," & nNumDoc & "," & Virg2Ponto(sTr(nValorPagoReal)) & ")"
                cn.Execute Sql, rdExecDirect
            End With
           'PROXIMO LANCAMENTO (RDOAUX)
           .MoveNext
        Loop
        
       'EFETUA BAIXA NO DOCUMENTO
        Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=" & nCodBanco & " ,CODAGENCIA ='" & sCodAgencia & "' , VALORPAGO=" & Virg2Ponto(sTr(nValorPago))
        Sql = Sql & " WHERE NUMDOCUMENTO=" & nNumDoc
        cn.Execute Sql, rdExecDirect
        
        GravaBaixaTmp
       'GoTo FIM
    End With
    
proximo:
Next

'SE TIVER DEBITO A CLASSIFICAR GRAVA NELE"
nResto = nSomaTotal2 - (nSomaClass + nSomaClass2)
If nResto > 0 Then
    Sql = "UPDATE DEBITOCLASSIFICAR SET VALORCLASS=" & Virg2Ponto(CStr(nResto)) & " WHERE DATARECEITA='" & Format(dDataCred, "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "'"
    cn.Execute Sql, rdExecDirect
End If

'EFETUA BAIXA NO ARQUIVO
Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA='" & Format(Now, "mm/dd/yyyy") & " ' WHERE "
Sql = Sql & "NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "'"
cn.Execute Sql, rdExecDirect

Fim:
Liberado

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)

If ((nVal * 100) / nTot) <= 100 Then
   Pb.Value = (nVal * 100) / nTot
Else
   Pb.Value = 100
End If

Me.Refresh
DoEvents

End Sub

Private Function CalculaJuros2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

If Year(dDataPagto) > Year(Now) Then
    CalculaJuros2 = 0
    Exit Function
End If

If dDataVencto >= dDataPagto Then
    CalculaJuros2 = 0
    Exit Function
End If

'****

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= dDataPagto Then
    CalculaJuros2 = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(dDataPagto) And Year(dDataVencto) = Year(dDataPagto) Then
    CalculaJuros2 = 0
    Exit Function
End If

'If Not dcJuros.Exists(Year(dDataPagto)) Then
'   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
'   CalculaJuros = 0
'   Exit Function
'End If

'MONTA O NOVO VENCIMENTO A PARTIR DO DIA 1 DO MES SUBSEQUENTE
nDia = Day(dDataVencto)
nMes = Month(dDataVencto)
nAno = Year(dDataVencto)
nDia = 1
If nMes = 12 Then
    nMes = 1
    nAno = nAno + 1
Else
    nMes = nMes + 1
End If

sDataVencto = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
dDataVencto = Format(sDataVencto, "dd/mm/yyyy")
nNumMes = Int(DateDiff("d", dDataVencto, dDataPagto) / 30) + 1

'****

'nNumMes = Int((DateDiff("d", dDataVencto, dDataPagto)) / 30)
Sql = "SELECT PERCJUROS FROM JUROS WHERE ANOJUROS=" & Year(dDataPagto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
        CalculaJuros2 = 0
        Exit Function
    Else
        nValorPerc = !PERCJUROS
    End If
   .Close
End With
nValorPerc = nValorPerc / 100

CalculaJuros2 = nValorDebito * nValorPerc * nNumMes
If CalculaJuros2 > 0 Then
   CalculaJuros2 = FormatNumber(CalculaJuros2, 3)
End If

End Function

Private Function CalculaJuros2oLD(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String

If Year(dDataPagto) > Year(Now) Then
    CalculaJuros2oLD = 0
    Exit Function
End If

If dDataVencto >= dDataPagto Then
    CalculaJuros2oLD = 0
    Exit Function
End If
nNumMes = Int((DateDiff("d", dDataVencto, dDataPagto)) / 30)
Sql = "SELECT PERCJUROS FROM JUROS WHERE ANOJUROS=" & Year(dDataPagto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
        CalculaJuros2oLD = 0
        Exit Function
    Else
        nValorPerc = !PERCJUROS
    End If
   .Close
End With
nValorPerc = nValorPerc / 100

CalculaJuros2oLD = nValorDebito * nValorPerc * nNumMes
If CalculaJuros2oLD > 0 Then
   CalculaJuros2oLD = FormatNumber(CalculaJuros2oLD, 3)
End If

End Function

Private Function CalculaMulta2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String


If dDataVencto >= dDataPagto Then
    CalculaMulta2 = 0
    Exit Function
End If

nNumDia = Abs(DateDiff("d", dDataPagto, dDataVencto))

If nNumDia = 0 Then
   CalculaMulta2 = 0
   Exit Function
End If

Sql = "SELECT MINDIA,MAXDIA,PERCDIA FROM MULTA WHERE ANOMULTA=" & Year(dDataVencto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         If nNumDia >= !MINDIA And nNumDia <= !MAXDIA Then
             nValorPerc = !PERCDIA
             Exit Do
         ElseIf nNumDia >= !MINDIA And !MAXDIA = 0 Then
             nValorPerc = !PERCDIA
             Exit Do
         End If
        .MoveNext
    Loop
End With

nValorPerc = nValorPerc / 100
CalculaMulta2 = nValorDebito * nValorPerc
If CalculaMulta2 > 0 Then
   CalculaMulta2 = FormatNumber(CalculaMulta2, 3)
End If

End Function

Private Function CalculaCorrecao2(nValorDebito As Double, dDataBase As Date, dDataVencto As Date) As Double

Dim UfirAtual As Double
Dim UfirBase As Double

If Year(dDataVencto) > Year(Now) Then
   CalculaCorrecao2 = 0
   Exit Function
End If
UfirAtual = RetornaUFIR(Year(dDataVencto))
UfirBase = RetornaUFIR(Year(dDataBase))

CalculaCorrecao2 = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecao2 > 0 Then
   CalculaCorrecao2 = FormatNumber(CalculaCorrecao2, 2)
End If
End Function

Private Sub Reativa()
Dim nNumDoc As Long
Dim x As Integer
Dim nCodReduz As Long
Dim nAnoExercicio As Integer
Dim nCodLanc As Integer
Dim nSeqLanc As Integer
Dim nNumParc As Integer
Dim nCompl As Integer
Dim RdoAux As rdoResultset, Sql As String

With grdF
    For x = 1 To .Rows - 1
        CallPb CLng(x), .Rows - 1
        nCodReduz = Val(grdF.TextMatrix(x, 11))
        If Val(Left$(grdF.TextMatrix(x, 10), 2)) > 0 Then
            GoTo proximo
        End If
        Sql = "SELECT LANCAMENTO.DESCREDUZ, DEBITOPARCELA.STATUSLANC, SITUACAOLANCAMENTO.DESCSITUACAO, DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE, DEBITOPARCELA.CODREDUZIDO, DEBITOPARCELA.ANOEXERCICIO, DEBITOPARCELA.CODLANCAMENTO,"
        Sql = Sql & "DEBITOPARCELA.SeqLancamento , DEBITOPARCELA.NumParcela, DEBITOPARCELA.CODCOMPLEMENTO FROM LANCAMENTO INNER JOIN DEBITOPARCELA ON LANCAMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO INNER JOIN "
        Sql = Sql & "SITUACAOLANCAMENTO ON DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO "
        Sql = Sql & "WHERE (DEBITOPARCELA.SEQLANCAMENTO < 100) AND (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.CODLANCAMENTO = 1) AND "
        Sql = Sql & "(DEBITOPARCELA.NUMPARCELA > 0) AND (DEBITOPARCELA.DATAVENCIMENTO = '" & Format(grdF.TextMatrix(x, 8), "mm/dd/yyyy") & "')"
            
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                GoTo proximo
            End If
            Sql = "SELECT NUMDOCUMENTO.NUMDOCUMENTO,VALORTAXADOC FROM PARCELADOCUMENTO INNER JOIN NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO "
            Sql = Sql & "Where CODREDUZIDO = " & !CODREDUZIDO & " And AnoExercicio = " & !AnoExercicio & " AND CodLancamento = " & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount = 0 Then
                GoTo proximo
            End If
            nNumDoc = RdoAux2!NumDocumento
            RdoAux2.Close
            Do Until .EOF
                nAnoExercicio = !AnoExercicio
                nCodLanc = !CodLancamento
                nSeqLanc = !SeqLancamento
                nNumParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                'ATUALIZA A TABELA DEBITOPAGO
                 Sql = "UPDATE DEBITOPAGO SET RESTITUIDO='" & Format(Now, "mm/dd/yyyy") & "' "
                 Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                 Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS NULL"
                 cn.Execute Sql, rdExecDirect
                'ATUALIZA A TABELA NUMDOCUMENTO
                 Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=0,CODAGENCIA=0,VALORPAGO=0 "
                 Sql = Sql & "WHERE NUMDOCUMENTO = " & nNumDoc
                 cn.Execute Sql, rdExecDirect
                 'SE TODOS OS REGISTROS EM DEBITOPAGO FOREM RESTITUIDOS ENTÃO ATUALIZA DÉBITOPARCELA
                 Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPAGO "
                 Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                 Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl & " AND RESTITUIDO IS  NULL"
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                     If !CONTADOR = 0 Then
                        'SE FOR ZERO SINAL QUE A PARCELA FOI TOTALMENTE RESTITUIDA
                        'ENTÃO PODEMOS ATUALIZAR O SEU STATUS PARA NÃO PAGO
                         Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=3 "
                         Sql = Sql & "WHERE CODREDUZIDO = " & nCodReduz & " AND ANOEXERCICIO = " & nAnoExercicio & " AND CODLANCAMENTO = " & nCodLanc & " AND "
                         Sql = Sql & "SEQLANCAMENTO = " & nSeqLanc & " AND NUMPARCELA = " & nNumParc & " AND CODCOMPLEMENTO = " & nCompl
                         cn.Execute Sql, rdExecDirect
                     End If
                 End With
              .MoveNext
            Loop
           .Close
        End With
        
proximo:
    Next
End With

Sql = "SELECT * FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    Sql = "DELETE FROM DEBITOPAGO WHERE DATARECEBIMENTO='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(lblBanco.Caption, 3)) & " AND RESTITUIDO IS NULL"
    cn.Execute Sql, rdExecDirect
End If

Sql = "DELETE FROM DEBITOCLASSIFICAR WHERE DATARECEITA='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "'"
cn.Execute Sql, rdExecDirect

Sql = "DELETE FROM RECEITACLASSIFICAR WHERE DATARECEITA='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "' AND NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "'"
cn.Execute Sql, rdExecDirect

Sql = "UPDATE ARQUIVOBANCO SET DATABAIXA=NULL WHERE NOMEARQ='" & grdArq.TextMatrix(grdArq.Row, 1) & "' AND DATACREDITO='" & Format(grdF.TextMatrix(1, 8), "mm/dd/yyyy") & "'"
cn.Execute Sql, rdExecDirect

MsgBox "Todos os lançamentos descriminados e seus documentos foram reativados.", vbInformation, "INFORMAÇÃO"
grdParc.Rows = 1

End Sub

