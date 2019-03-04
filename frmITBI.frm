VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmITBI 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imposto sobre Transmissão de Bens Intervivos"
   ClientHeight    =   5460
   ClientLeft      =   4635
   ClientTop       =   3075
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5460
   ScaleWidth      =   7590
   Begin VB.CheckBox chkRural 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Rural"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6720
      TabIndex        =   60
      Top             =   420
      Width           =   705
   End
   Begin VB.TextBox txtImovel 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   6510
      MaxLength       =   6
      TabIndex        =   2
      Top             =   90
      Width           =   945
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   2640
      TabIndex        =   58
      Top             =   5610
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tipo de Endereço"
      Height          =   600
      Left            =   2475
      TabIndex        =   54
      Top             =   4815
      Width           =   2805
      Begin VB.OptionButton optEnd 
         Caption         =   "Residencial"
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   56
         Top             =   270
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.OptionButton optEnd 
         Caption         =   "Comercial"
         Height          =   240
         Index           =   1
         Left            =   1440
         TabIndex        =   55
         Top             =   270
         Width           =   1230
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Endereço Comercial"
      Height          =   1095
      Left            =   45
      TabIndex        =   40
      Top             =   1620
      Width           =   7485
      Begin VB.Label lblUF 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4725
         TabIndex        =   57
         Top             =   810
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço.......:"
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   53
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblRuaEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1275
         TabIndex        =   52
         Top             =   270
         Width           =   4860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Index           =   3
         Left            =   6330
         TabIndex        =   51
         Top             =   270
         Width           =   405
      End
      Begin VB.Label lblNumEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6750
         TabIndex        =   50
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.:"
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   49
         Top             =   525
         Width           =   1155
      End
      Begin VB.Label lblComplentrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1275
         TabIndex        =   48
         Top             =   525
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...:"
         Height          =   225
         Index           =   5
         Left            =   4290
         TabIndex        =   47
         Top             =   540
         Width           =   690
      End
      Begin VB.Label lblBairroEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4995
         TabIndex        =   46
         Top             =   525
         Width           =   2460
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade...........:"
         Height          =   225
         Index           =   7
         Left            =   45
         TabIndex        =   45
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label lblCidadeEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1275
         TabIndex        =   44
         Top             =   780
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep......:"
         Height          =   225
         Index           =   8
         Left            =   5325
         TabIndex        =   43
         Top             =   795
         Width           =   585
      End
      Begin VB.Label lblCepEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6015
         TabIndex        =   42
         Top             =   780
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...:"
         Height          =   225
         Index           =   12
         Left            =   4275
         TabIndex        =   41
         Top             =   795
         Width           =   390
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Endereço Residencial"
      Height          =   870
      Left            =   45
      TabIndex        =   29
      Top             =   720
      Width           =   7500
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...................:"
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   39
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label lblRua 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1710
         TabIndex        =   38
         Top             =   270
         Width           =   3690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Index           =   1
         Left            =   6060
         TabIndex        =   37
         Top             =   270
         Width           =   405
      End
      Begin VB.Label lblNumImovel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6480
         TabIndex        =   36
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label lblCep 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6480
         TabIndex        =   35
         Top             =   525
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep.:"
         Height          =   225
         Index           =   9
         Left            =   6060
         TabIndex        =   34
         Top             =   525
         Width           =   420
      End
      Begin VB.Label lblCompl 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1710
         TabIndex        =   33
         Top             =   525
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.............:"
         Height          =   225
         Index           =   10
         Left            =   45
         TabIndex        =   32
         Top             =   525
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro..:"
         Height          =   225
         Index           =   11
         Left            =   3450
         TabIndex        =   31
         Top             =   525
         Width           =   570
      End
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4050
         TabIndex        =   30
         Top             =   525
         Width           =   1845
      End
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "DAM"
      Height          =   225
      Index           =   0
      Left            =   1215
      TabIndex        =   9
      Top             =   4875
      Value           =   -1  'True
      Width           =   780
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Certidão"
      Height          =   225
      Index           =   1
      Left            =   1215
      TabIndex        =   25
      Top             =   5130
      Width           =   1050
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   885
      Left            =   1170
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3165
      Width           =   6345
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1170
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4080
      Width           =   1155
   End
   Begin VB.TextBox txtArtigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3210
      MaxLength       =   15
      TabIndex        =   6
      Top             =   4080
      Width           =   1200
   End
   Begin VB.TextBox txtTipo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5190
      MaxLength       =   20
      TabIndex        =   7
      Top             =   4080
      Width           =   2310
   End
   Begin VB.TextBox txtFunc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   8
      Top             =   4395
      Width           =   6330
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1695
      MaxLength       =   6
      TabIndex        =   1
      Top             =   90
      Width           =   945
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   330
      Left            =   6480
      TabIndex        =   13
      ToolTipText     =   "Sair da Tela"
      Top             =   4950
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
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
      MICON           =   "frmITBI.frx":0000
      PICN            =   "frmITBI.frx":001C
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
      Height          =   330
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Emissão do/a DAM/Certidão"
      Top             =   4950
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   582
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
      MICON           =   "frmITBI.frx":008A
      PICN            =   "frmITBI.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1260
      Left            =   0
      TabIndex        =   0
      Top             =   6045
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   2223
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Código     |^Ano     |^Lanc. |^Seq  |^Parc. |^Compl. |^Vencimento      |>Vl.Lançado  |<Num.Documento      "
   End
   Begin esMaskEdit.esMaskedEdit mskVencto 
      Height          =   285
      Left            =   1170
      TabIndex        =   3
      Top             =   2850
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmITBI.frx":0200
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
   Begin prjChameleon.chameleonButton cmdLoadHist 
      Height          =   300
      Left            =   2325
      TabIndex        =   17
      ToolTipText     =   "Emissão do/a DAM/Certidão"
      Top             =   2820
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "Carregar Histórico"
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
      MICON           =   "frmITBI.frx":021C
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
      Caption         =   "Imóvel"
      Height          =   225
      Index           =   20
      Left            =   5910
      TabIndex        =   59
      Top             =   120
      Width           =   585
   End
   Begin VB.Label lblTipoEnd 
      Caption         =   "Label3"
      Height          =   345
      Left            =   675
      TabIndex        =   28
      Top             =   5535
      Width           =   465
   End
   Begin VB.Label lblUF2 
      BackColor       =   &H00000000&
      Caption         =   "Label3"
      Height          =   345
      Left            =   135
      TabIndex        =   27
      Top             =   5625
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Emitir por.......:"
      Height          =   210
      Left            =   60
      TabIndex        =   26
      Top             =   4875
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Vencto..:"
      Height          =   225
      Index           =   13
      Left            =   60
      TabIndex        =   24
      Top             =   2895
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição......:"
      Height          =   225
      Index           =   14
      Left            =   60
      TabIndex        =   23
      Top             =   3210
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "(Max 2000)"
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   15
      Left            =   135
      TabIndex        =   22
      Top             =   3450
      Width           =   885
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total.....:"
      Height          =   225
      Index           =   16
      Left            =   75
      TabIndex        =   21
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Artigo.....:"
      Height          =   225
      Index           =   17
      Left            =   2445
      TabIndex        =   20
      Top             =   4140
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo.....:"
      Height          =   225
      Index           =   18
      Left            =   4530
      TabIndex        =   19
      Top             =   4125
      Width           =   630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Funcionário....:"
      Height          =   225
      Index           =   19
      Left            =   75
      TabIndex        =   18
      Top             =   4455
      Width           =   1095
   End
   Begin VB.Label lblNumInsc 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3750
      TabIndex        =   16
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido/I.M.:"
      Height          =   225
      Index           =   0
      Left            =   45
      TabIndex        =   15
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblProp 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1695
      TabIndex        =   14
      Top             =   435
      Width           =   4680
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário.................:"
      Height          =   225
      Left            =   45
      TabIndex        =   12
      Top             =   405
      Width           =   1695
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ..:"
      Height          =   225
      Left            =   2790
      TabIndex        =   10
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmITBI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xImovel As clsImovel
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset

Private Sub cmdLoadHist_Click()
If Not IsDate(mskVencto.Text) Then
    MsgBox "Digite a Data de Vencimento.", vbCritical, "Atenção"
    Exit Sub
Else
    txtDesc.Text = ""
    Sql = "SELECT OBSPARCELA.CODREDUZIDO, OBSPARCELA.OBS, DEBITOPARCELA.DATAVENCIMENTO "
    Sql = Sql & "FROM OBSPARCELA INNER JOIN DEBITOPARCELA ON OBSPARCELA.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO "
    Sql = Sql & "WHERE OBSPARCELA.CODREDUZIDO=" & Val(txtCod.Text) & " AND DEBITOPARCELA.DATAVENCIMENTO = '" & Format(mskVencto.Text, "mm/dd/yyyy") & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            Do Until .EOF
                txtDesc.Text = txtDesc.Text & SubNull(!obs) & " "
               .MoveNext
            Loop
        Else
            txtDesc.Text = ""
            MsgBox "Não localizado observação/débito para este Vencimento.", vbExclamation, "Atenção"
        End If
       .Close
    End With
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataDam As String, sDataVencto As String, RdoAuxCli As rdoResultset
Dim nCodReduz As Long, sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, sQuadras As String, sLotes As String
Dim sUsuario As String, nNumDoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String, bGerado As Boolean
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date
Dim sBarra As String, sDigitavel2 As String, nValorDam As Double, nValorPrincDam As Double, nNumGuia As Long, sTipoEnd As String
Dim sValor As String, dDataVencto As Date, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String

If MsgBox("Confirma Emissão da DAM ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect

nValorTaxa = 0
sDoc = ""
nPos = 0
nValorDoc = 0
nValorDam = 0
nValorPrincDam = 0
sUsuario = NomeDeLogin
sDataDam = "21/08/2017"
bMulta = False

Sql = "SELECT proprietario.codreduzido FROM proprietario INNER JOIN debitoparcela ON proprietario.codreduzido = debitoparcela.codreduzido WHERE(proprietario.codcidadao = 578651) AND "
Sql = Sql & "(proprietario.principal = 1) AND (debitoparcela.anoexercicio = 2017) AND (debitoparcela.numparcela = 1) AND (debitoparcela.codlancamento = 79) ORDER BY proprietario.codreduzido"
Set RdoAuxCli = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAuxCli
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        DoEvents
        Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !Inscricao
            sNome = !NomeCidadao
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
                If sDoc = "" Then
                    sDoc = SubNull(!rg)
                End If
            End If
            sEnd = SubNull(!Logradouro)
            nNum = Val(SubNull(!Li_Num))
            sCompl = Left(SubNull(!Li_Compl), 30)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!li_uf)
            sQuadras = Left(SubNull(!Li_Quadras), 15)
            sLotes = Left(SubNull(!Li_Lotes), 10)
           .Close
        End With
    
        'grava documento
        Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!maximo) Then
           nNumDoc = 0
        Else
           nNumDoc = RdoAux!maximo + 1
        End If
        RdoAux.Close
        sNumDoc = CStr(nNumDoc) & "-" & RetornaDVNumDoc(nNumDoc)
        sNumDoc2 = CStr(nNumDoc) & RetornaDVNumDoc(nNumDoc)
        sNumDoc3 = CStr(nNumDoc) & Modulo11(nNumDoc)
    
        Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO,TIPODOC,emissor) VALUES("
        Sql = Sql & nNumDoc & ",'" & Format(sDataDam, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & IIf(bMulta, 1, 0) & "," & 0 & ",1,'" & NomeDeLogin & " (DAM)" & "')"
        cn.Execute Sql, rdExecDirect
        
        Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela,debitoparcela.CODCOMPLEMENTO , debitoparcela.DataVencimento, "
        Sql = Sql & "debitotributo.CodTributo, debitotributo.ValorTributo, lancamento.descreduz, TRIBUTO.abrevtributo FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND "
        Sql = Sql & "debitoparcela.anoexercicio = debitotributo.anoexercicio AND debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
        Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN lancamento ON debitoparcela.codlancamento = lancamento.codlancamento INNER JOIN "
        Sql = Sql & "tributo ON debitotributo.codtributo = tributo.codtributo WHERE (debitoparcela.codreduzido = " & nCodReduz & ") AND (debitoparcela.anoexercicio = 2017) AND (debitoparcela.codlancamento = 79)"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                nAno = !AnoExercicio
                nLanc = !CodLancamento
                nSeq = !SeqLancamento
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
                sDA = "N"
                sAj = "N"
                nValorPrincipal = FormatNumber(CDbl(!ValorTributo), 2)
                nValorJuros = FormatNumber(0, 2)
                nValorMulta = FormatNumber(0, 2)
                nValorCorrecao = FormatNumber(0, 2)
                nValorTotal = FormatNumber(nValorPrincipal, 2)
                nValorDoc = nValorTotal
                sFullTrib = !ABREVTRIBUTO
                'GRAVA PARCELADOCUMENTO
                
                Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,VALORJUROS,VALORMULTA,VALORCORRECAO,PLANO) VALUES(" & nCodReduz & ","
                Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & ","
                Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & "," & 0 & ")"
                cn.Execute Sql, rdExecDirect
                
                Sql = "insert boleto(usuario,computer,sid,seq,inscricao,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,quadra,lote,numdoc,nomefunc,datadam,fulllanc,fulltrib,"
                Sql = Sql & "anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datavencto,aj,da,principal,juros,multa,correcao,total,numdoc2,valordam) values('"
                Sql = Sql & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & nPos & ",'" & sInsc & "'," & nCodReduz & ",'" & Left(Mask(sNome), 40) & "','" & sDoc & "','"
                Sql = Sql & Left(Mask(sEnd), 40) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Mask(sCidade) & "','" & sUF & "','" & Mask(sQuadras) & "','"
                Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataDam, sDataFormat) & "','" & sLANCAMENTO & "','" & sFullTrib & "'," & nAno & ","
                Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",'" & Format(sDataVencto, sDataFormat) & "','" & sAj & "','" & sDA & "'," & Virg2Ponto(Format(nValorPrincipal, "#0.00")) & ","
                Sql = Sql & Virg2Ponto(Format(nValorJuros, "#0.00")) & "," & Virg2Ponto(Format(nValorMulta, "#0.00")) & "," & Virg2Ponto(Format(nValorCorrecao, "#0.00")) & "," & Virg2Ponto(Format(nValorTotal, "#0.00")) & ",'" & sNumDoc2
                Sql = Sql & "'," & Virg2Ponto(RemovePonto(CStr(nValorTotal * 3))) & ")"
                cn.Execute Sql, rdExecDirect
                
                Sql = "update numdocumento set valorguia=" & Virg2Ponto(CStr(nValorTotal * 3)) & " where numdocumento=" & nNumDoc
                cn.Execute Sql, rdExecDirect
                
                nPos = nPos + 1
                            
               .MoveNext 'parcela
            Loop
           .Close
        End With 'parcela
       
        sValor = CStr(nValorTotal * 3)
        dDataVencto = CDate(sDataDam)
        nNumGuia = nNumDoc
        NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumDoc, nCodReduz)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
        
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        Sql = "update boleto set numbarra2a='" & NumBarra2a & "',numbarra2b='" & NumBarra2b & "',numbarra2c='" & NumBarra2c & "',numbarra2d='" & NumBarra2d & "',codbarra='" & Mask(StrBarra2) & "' where sid=" & nSid & " and numdoc2=" & Val(sNumDoc2)
        cn.Execute Sql, rdExecDirect
       
       .MoveNext 'cliente
    Loop
   .Close
End With

'Exit Sub

frmReport.ShowReport2 "BOLETODAM_V42", frmMdi.HWND, Me.HWND, nSid, nSid

Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Form_Load()
Set xImovel = New clsImovel
Centraliza Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub txtCod_GotFocus()

txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    txtCod_LostFocus
    Exit Sub
End If

Tweak txtCod, KeyAscii, IntegerPositive

End Sub

Private Sub txtCod_LostFocus()
Dim nCodImovel As Long, sEnd As String

If Val(txtCod.Text) = 0 Then Exit Sub
If Val(txtCod.Text) < 500000 Then
    MsgBox "ITBI só pode ser emitida para Cidadão.", vbExclamation, "Atenção"
    Exit Sub
End If
nCodImovel = Val(txtCod.Text)
Limpa
Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & txtCod.Text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If !Inativo = 1 Then
           MsgBox "Este imóvel encontra-se inativo.", vbExclamation, "Atenção"
           Exit Sub
        End If
        lblRS.Caption = "Proprietário"
        CarregaImovel nCodImovel
    Else
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,CNPJ,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & txtCod.Text
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               If Not IsNull(!dataencerramento) Or !dataencerramento <> CDate("01/01/1900") Then
                  MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
                  Exit Sub
               End If
               lblNumInsc.Caption = SubNull(!Cnpj)
              'suspenção
               Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.Text
               Sql = Sql & " ORDER BY DATAEVENTO DESC"
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                       If !CODTIPOEVENTO = 2 Then
                           MsgBox "Esta empresa esta SUSPENSA", vbExclamation, "Atenção"
                           Exit Sub
                       End If
                   End If
                  .Close
               End With
               
               lblNumInsc.Caption = SubNull(!INSCESTADUAL)
               lblRS.Caption = "Raz.Social"
               lblProp.Caption = !RazaoSocial
               lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
               lblNumImovel.Caption = Val(SubNull(!Numero))
               lblCep.Caption = IIf(IsNull(!Cep), "", Left$(!Cep, 5) & "-" & Right$(!Cep, 3))
               lblCompl.Caption = SubNull(!Complemento)
               Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO=" & !CodBairro
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                        lblBairro.Caption = !DescBairro
                   Else
                        lblBairro.Caption = ""
                   End If
                  .Close
               End With
               Sql = "SELECT NOMELOGRADOURO,NUMIMOVEL,COMPLEMENTO,UF,CIDADE.DESCCIDADE AS DESCCIDADE1,"
               Sql = Sql & "BAIRRO.DESCBAIRRO AS DESCBAIRRO1,CEP,MOBILIARIOENDENTREGA.DESCBAIRRO,"
               Sql = Sql & "MOBILIARIOENDENTREGA.DESCCIDADE FROM CIDADE INNER JOIN BAIRRO ON "
               Sql = Sql & "CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE RIGHT OUTER Join "
               Sql = Sql & "MOBILIARIOENDENTREGA ON BAIRRO.CODCIDADE = MOBILIARIOENDENTREGA.CODCIDADE AND "
               Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.Text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                    If .RowCount > 0 Then
                        lblTipoEnd.Caption = "(Endereço de Entrega Específico)"
                        lblRuaEntrega.Caption = SubNull(!NomeLogradouro)
                        lblNumEntrega.Caption = SubNull(!NUMIMOVEL)
                        lblComplentrega.Caption = SubNull(!Complemento)
                        lblBairroEntrega.Caption = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        lblCidadeEntrega.Caption = IIf(IsNull(!descCidade), SubNull(!DESCCIDADE1), SubNull(!descCidade))
                        lblCepEntrega.Caption = SubNull(!Cep)
                        lblUF.Caption = SubNull(!UF)
                    Else
                        lblTipoEnd.Caption = "(Endereço da Empresa)"
                        lblRuaEntrega.Caption = lblRua.Caption
                        lblNumEntrega.Caption = lblNumImovel.Caption
                        lblComplentrega.Caption = lblCompl.Caption
                        lblBairroEntrega.Caption = lblBairro.Caption
                        lblCidadeEntrega.Caption = "JABOTICABAL"
                        lblCepEntrega.Caption = lblCep.Caption
                        lblUF.Caption = "SP"
                    End If
                   .Close
                End With
            Else

                
                Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO,"
                Sql = Sql & "CODLOGRADOURO2 AS fCODLOGRADOURO2,NUMIMOVEL2 AS fNUMIMOVEL2,"
                Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO2,CODBAIRRO2 AS fCODBAIRRO2,CODCIDADE2 AS fCODCIDADE2,SIGLAUF2 AS fSIGLAUF2,"
                Sql = Sql & "CEP2 AS fCEP2,TELEFONE2 AS fTELEFONE2,EMAIL2 AS fEMAIL2,RG AS fRG2,NOMELOGRADOURO2 AS fNOMELOGRADOURO2,ORGAO AS fORGAO2"
                Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(txtCod.Text)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                     If .RowCount > 0 Then
                          lblProp.Caption = !NomeCidadao
                          
                          
                          If Val(SubNull(!FCodLogradouro)) > 0 Then
                              Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                              Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                              Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                              Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                              Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                              With RdoS
                                  If .RowCount > 0 Then
                                     sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                  Else
                                     sEnd = ""
                                  End If
                                 .Close
                              End With
                          Else
                             sEnd = SubNull(!FNomeLogradouro)
                          End If
                          lblRua.Caption = sEnd
                          lblNumImovel.Caption = SubNull(!fNUMIMOVEL)
                          lblCompl.Caption = SubNull(!fcomplemento)
                          lblCep.Caption = SubNull(!FCEP)
                          If lblCep.Caption = "" Then
                            lblCep.Caption = RetornaCEP(Val(SubNull(!FCodLogradouro)), Val(SubNull(!fNUMIMOVEL)))
                          End If
                          If Not IsNull(!fCodBairro) Then
                              Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & SubNull(!fsiglauf) & "' AND CODCIDADE=" & Val(SubNull(!fCodCidade)) & " AND CODBAIRRO=" & !fCodBairro
                              Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                              lblBairro = SubNull(RdoS!DescBairro)
                          End If
                     
                     
                          If Val(SubNull(!fcodlogradouro2)) > 0 Then
                              Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                              Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                              Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                              Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !fcodlogradouro2
                              Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                              With RdoS
                                  If .RowCount > 0 Then
                                     sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                                  Else
                                     sEnd = ""
                                  End If
                                 .Close
                              End With
                          Else
                             sEnd = SubNull(!FNomeLogradouro2)
                          End If
                          lblRuaEntrega.Caption = sEnd
                          lblNumEntrega.Caption = SubNull(!fnumimovel2)
                          lblComplentrega.Caption = SubNull(!fcomplemento2)
                          lblCepEntrega.Caption = SubNull(!fcep2)
                          If lblCepEntrega.Caption = "" Then
                            lblCepEntrega.Caption = RetornaCEP(Val(SubNull(!fcodlogradouro2)), Val(SubNull(!fnumimovel2)))
                          End If
                           
                          Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & SubNull(!fsiglauf2) & "' AND CODCIDADE=" & Val(SubNull(!fCodCidade2))
                          Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                          lblCidadeEntrega = SubNull(RdoS!descCidade)
                          
                          If Not IsNull(!fCodBairro2) Then
                              Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf2 & "' AND CODCIDADE=" & !fCodCidade2 & " AND CODBAIRRO=" & !fCodBairro2
                              Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                              lblBairroEntrega = SubNull(RdoS!DescBairro)
                          End If
                            lblUF.Caption = SubNull(!fsiglauf2)
                       If SubNull(!CPF) <> "" Then
                           lblNumInsc.Caption = !CPF
                       Else
                            If SubNull(!Cnpj) <> "" Then
                                lblNumInsc.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
                            Else
                                lblNumInsc.Caption = SubNull(!frg)
                            End If
                       End If
                     
                     End If
                    .Close
                    
                End With


            End If
           .Close
        End With
    End If
End With
End Sub

Private Sub CarregaImovel(nCodigoImovel As Long)
Dim Sql As String, RdoAux As rdoResultset

Ocupado
With xImovel
    .CarregaImovel nCodigoImovel
    If .CodigoImovel > 0 Then
          lblProp.Caption = .NomePropPrincipal
          lblRua.Caption = Trim$(.AbrevTipoLog) & " " & Trim$(.AbrevTitLog) & " " & .NomeLogradouro
          lblNumImovel.Caption = .Li_Num
          lblCep.Caption = RetornaCEP(.CodLogr, .Li_Num)
          lblCompl.Caption = .Li_Compl
          lblBairro.Caption = .DescBairro
          Select Case .Ee_TipoEnd
                Case 0
                    lblTipoEnd.Caption = "(Endereço do Imóvel)"
                    lblRuaEntrega.Caption = lblRua.Caption
                    lblNumEntrega.Caption = lblNumImovel.Caption
                    lblComplentrega.Caption = lblCompl.Caption
                    lblBairroEntrega.Caption = lblBairro.Caption
                    lblCidadeEntrega.Caption = "JABOTICABAL"
                    lblCepEntrega.Caption = lblCep.Caption
                    lblUF.Caption = lblUF.Caption
                Case 1
                    lblTipoEnd.Caption = "(Endereço do Proprietário)"
                    CarregaEndCidadao .CodPropPrincipal
                Case 2
                    lblTipoEnd.Caption = "(Endereço de Entrega Específico)"
                    lblRuaEntrega.Caption = .Ee_NomeLog
                    lblNumEntrega.Caption = .Ee_NumImovel
                    lblComplentrega.Caption = .Ee_Complemento
                    Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & .Ee_Uf & "' AND CODCIDADE=" & .Ee_Cidade & " AND CODBAIRRO=" & .Ee_Bairro
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            lblBairroEntrega.Caption = !DescBairro
                        End If
                       .Close
                    End With
                    lblCidadeEntrega.Caption = .Ee_Cidade
                    lblCepEntrega.Caption = .Ee_Cep
                    lblUF.Caption = .Ee_Uf
          End Select
    End If
End With

FIM:
Liberado

End Sub

Private Sub CarregaEndCidadao(nCodigo As Long)

Sql = "SELECT CIDADAO.CODCIDADAO,vwLOGRADOUROCEP.ABREVTIPOLOG,vwLOGRADOUROCEP.ABREVTITLOG,"
Sql = Sql & "vwLOGRADOUROCEP.NOMELOGRADOURO,vwLOGRADOUROCEP.CEP, CIDADAO.NUMIMOVEL,"
Sql = Sql & "CIDADAO.COMPLEMENTO, CIDADAO.CODBAIRRO,CIDADAO.CODCIDADE, CIDADAO.SIGLAUF,"
Sql = Sql & "Cidade.DESCCIDADE , BAIRRO.DescBairro FROM CIDADAO INNER JOIN vwLOGRADOUROCEP ON "
Sql = Sql & "CIDADAO.CODLOGRADOURO = vwLOGRADOUROCEP.CODLOGRADOURO Inner Join BAIRRO ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND "
Sql = Sql & "CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO INNER JOIN "
Sql = Sql & "CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = Cidade.CODCIDADE "
Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & nCodigo
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    lblRuaEntrega.Caption = SubNull(!NomeLogradouro)
    lblNumEntrega.Caption = SubNull(!NUMIMOVEL)
    lblComplentrega.Caption = SubNull(!Complemento)
    lblBairroEntrega.Caption = SubNull(!DescBairro)
    lblCidadeEntrega.Caption = SubNull(!descCidade)
    lblCepEntrega.Caption = SubNull(!Cep)
    lblUF.Caption = SubNull(!Cep)
End With

End Sub

Private Sub Limpa()
lblProp.Caption = ""
lblRua.Caption = ""
lblNumImovel.Caption = ""
lblCompl.Caption = ""
lblBairro.Caption = ""
lblCep.Caption = ""
lblRuaEntrega.Caption = ""
lblNumEntrega.Caption = ""
lblComplentrega.Caption = ""
lblBairroEntrega.Caption = ""
lblCidadeEntrega.Caption = ""
lblCepEntrega.Caption = ""
lblUF.Caption = ""
lblNumInsc.Caption = ""
lblTipoEnd.Caption = ""
LimpaMascara mskVencto
txtDesc.Text = ""
txtValor.Text = ""
txtArtigo.Text = ""
txtTipo.Text = ""
txtFunc.Text = ""
Opt(0).value = True


End Sub
Private Sub cmdBaixa_Click()

If lblProp.Caption = "" Then
    MsgBox "Selecione o Contribuinte.", vbExclamation, "Atenção"
    Exit Sub
End If


If Val(txtImovel.Text) = 0 And chkRural.value = vbUnchecked Then
    MsgBox "Digite o código do imóvel.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(lblCep.Caption) = 0 Then
    MsgBox "CEP obrigatório.", vbExclamation, "Atenção"
    Exit Sub
End If


If lblNumInsc.Caption = "" Then
    MsgBox "Contribuinte deve ter um CPF/CNPJ válido.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskVencto.Text) And Opt(0).value Then
    MsgBox "Data de vencimento inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If txtDesc = "" Then
    MsgBox "Digite a Descrição.", vbExclamation, "Atenção"
    Exit Sub
End If

If txtValor.Text = "" Then txtValor.Text = 0
If CDbl(txtValor.Text) = 0 And Opt(0).value = True Then
    MsgBox "Digite o Valor.", vbExclamation, "Atenção"
    Exit Sub
End If


If txtArtigo.Text = "" And Opt(0).value = False Then
    MsgBox "Digite o Artigo.", vbExclamation, "Atenção"
    Exit Sub
End If

If txtTipo.Text = "" And Opt(0).value = False Then
    MsgBox "Digite o Tipo de Transação.", vbExclamation, "Atenção"
    Exit Sub
End If

If Opt(1).value = True And txtFunc.Text = "" Then
    MsgBox "Digite o Nome do Funcionário que irá assinar a certidão.", vbExclamation, "Atenção"
    Exit Sub
End If

'GravaCarneTmp
'EmiteBoleto
EmiteBoletoRegistrado
If bGerado Then Limpa
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub

Private Sub GravaCarneTmp()
On Error Resume Next

Dim x As Integer
Dim RdoAux2 As rdoResultset, qd As New rdoQuery
Dim sNumInsc As String
Dim nCodReduz As Long
Dim sNomeResp As String
Dim sTipoImposto As String
Dim sEndImovel As String
Dim nNumImovel As Integer
Dim sComplImovel As String
Dim sBairroImovel As String
Dim nCodLogr As Long
Dim sEndEntrega As String
Dim nNumEntrega As Integer
Dim sBairroEntrega As String
Dim sComplEntrega As String
Dim sCepEntrega As String
Dim sCidadeEntrega As String
Dim sUFEntrega As String
Dim sDescImposto As String
Dim nAno As Integer
Dim sNumProc As String
Dim dDataProc As Date
Dim dDataVencto As Date
Dim nNumDoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim sVencimento As String
Dim nCodLanc As Integer
Dim nSeq As Integer
Dim nComplemento As Integer
Dim nValorTotal As Double
Dim sValorParc As String
Dim nValorParc As Double
Dim nValorParcUnica As Double
Dim NumBarra1 As String, sCPF As String
Dim StrBarra1 As String
Dim NumBarra2 As String
Dim NumBarra2a As String
Dim NumBarra2b As String
Dim NumBarra2c As String
Dim NumBarra2d As String
Dim StrBarra2 As String
Dim nLastCod As Long
Dim sFullTrib As String
Dim nSid As Long

nSid = Int(Rnd(10) * 1000000)
nCodReduz = Val(txtCod.Text)
If Not IsDate(mskVencto.Text) Then mskVencto.Text = Format(Now, "dd/mm/yyyy")

If Opt(0).value Then
    If MsgBox("Confirma criação da Guia ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
       bGerado = False
       Exit Sub
    End If
Else
    If MsgBox("Confirma criação da Certidão ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
       bGerado = False
       Exit Sub
    End If
End If


nValorParc = 0
nValorParcUnica = 0
'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
   nLastCod = 0
Else
   nLastCod = RdoAux!maximo
End If
RdoAux.Close

sCPF = lblNumInsc.Caption
If nCodReduz < 500000 Then
   Sql = "SELECT CODREDUZIDO,CPF,CNPJ,RG,ORGAO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodReduz
Else
   Sql = "SELECT CODCIDADAO,CPF,CNPJ,RG,ORGAO FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
End If
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        If Not IsNull(!CPF) And Trim(!CPF) <> "" Then
           sCPF = !CPF
        ElseIf Not IsNull(!Cnpj) And Trim(!Cnpj) <> "" Then
           sCPF = !Cnpj
        ElseIf Not IsNull(!rg) Then
           sCPF = !rg
        End If
    End If
End With

'CARREGA GRID TEMPORARIO
grdTemp.Rows = 1
nCodLanc = 36 'ITBI
nValorTxExpParc = 1.6
'Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 3"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'     nValorTxExpParc = FormatNumber(!VALORDAM, 2)
'    .Close
'End With

'CALCULA O VALOR PARCELADO

nValorTotal = CDbl(txtValor.Text)
nValorParc = nValorTotal

'MONTA TRIBUTOS
sVencimento = mskVencto.Text
nAno = Year(CDate(sVencimento))
'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
Sql = Sql & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nCodLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & 1
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
    Else
        nSeq = 0
    End If
   .Close
End With
nLastCod = nLastCod + 1
grdTemp.AddItem nCodReduz & Chr(9) & nAno & Chr(9) & nCodLanc & Chr(9) & nSeq & Chr(9) & 1 & Chr(9) & 0 & Chr(9) & sVencimento & Chr(9) & nValorParc & Chr(9) & nLastCod

'DADOS CABEÇALHO
sNumProc = Format(txtCod.Text, "000000") & "/" & sTr(Year(Now))
dDataProc = Format(Now, "dd/mm/yyyy")
sDescImposto = "ITBI"
NumBarra1 = Format(ExtraiNumero(sNumProc), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)

'GERAÇÃO DOS DÉBITOS
If Opt(0).value = True Then
    With grdTemp
         x = 1
         'GRAVA DEBITOPARCELA    // (STATUS 3 - NAO PAGO)
'         Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'         Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
'         Sql = Sql & "NUMPROCESSO,USUARIO) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & ","
'         Sql = Sql & .TextMatrix(x, 2) & "," & nSeq & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
'         Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
'         Sql = Sql & 1 & ",'" & sNumProc & "','" & Left$(NomeDeLogin, 25) & "')"
         Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
         Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,CODMOEDA,"
         Sql = Sql & "NUMPROCESSO,USERID) VALUES(" & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & ","
         Sql = Sql & .TextMatrix(x, 2) & "," & nSeq & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ","
         Sql = Sql & 3 & ",'" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
         Sql = Sql & 1 & ",'" & sNumProc & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
         cn.Execute Sql, rdExecDirect
         sFullTrib = ""
         'GRAVA HISTORICOPARCELA    // (STATUS 3 - NAO PAGO)
'         Sql = "SELECT CODREDUZIDO FROM LANCAMENTOOBS WHERE CODREDUZIDO=" & Val(.TextMatrix(x, 0)) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
'         Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & .TextMatrix(x, 4)
'         Sql = Sql & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
'         Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'         'With RdoAux
'             If RdoAux.RowCount = 0 Then
'                 Sql = "INSERT LANCAMENTOOBS (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,OBS) VALUES("
'                 Sql = Sql & Val(.TextMatrix(x, 0)) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & nSeq & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & ",'" & Mask(txtDesc.text) & "')"
'             Else
'                 Sql = "UPDATE LANCAMENTOOBS SET OBS='" & Mask(txtDesc.text) & "' WHERE CODREDUZIDO=" & Val(.TextMatrix(x, 0)) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
'                 Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & .TextMatrix(x, 4)
'                 Sql = Sql & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
 '             End If
'             cn.Execute Sql, rdExecDirect
'            RdoAux.Close
'         'End With
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & Val(.TextMatrix(x, 0)) & " AND ANOEXERCICIO=" & .TextMatrix(x, 1)
            Sql = Sql & " AND CODLANCAMENTO=" & .TextMatrix(x, 2) & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & .TextMatrix(x, 4)
            Sql = Sql & " AND CODCOMPLEMENTO=" & .TextMatrix(x, 5)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeqLanc = 1
                Else
                    nSeqLanc = !maximo + 1
                End If
               .Close
            End With
            sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
'            Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & Val(.TextMatrix(x, 0)) & "," & .TextMatrix(x, 1) & ","
'            Sql = Sql & .TextMatrix(x, 2) & "," & nSeq & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeqLanc & ",'" & Mask(txtDesc.Text) & "','" & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
            Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & Val(.TextMatrix(x, 0)) & "," & .TextMatrix(x, 1) & ","
            Sql = Sql & .TextMatrix(x, 2) & "," & nSeq & "," & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & nSeqLanc & ",'" & Mask(txtDesc.Text) & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sData, "mm/dd/yyyy") & "')"
            cn.Execute Sql, rdExecDirect
            sFullTrib = Mask(txtDesc.Text)
        'GRAVA DEBITOTRIBUTO
         Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
         Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
         Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & nSeq & ","
         Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 84 & "," & Virg2Ponto(CStr(nValorParc)) & ")"
         cn.Execute Sql, rdExecDirect
'         Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'         Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
'         Sql = Sql & .TextMatrix(x, 0) & "," & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & nSeq & ","
'         Sql = Sql & .TextMatrix(x, 4) & "," & .TextMatrix(x, 5) & "," & 3 & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ")"
'         cn.Execute Sql, rdExecDirect
'         sFullTrib = "Histórico: " & Mask(txtDesc.text)
        'GRAVA NUMDOCUMENTO
         Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC) VALUES("
         Sql = Sql & .TextMatrix(x, 8) & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ")"
         cn.Execute Sql, rdExecDirect
        'GRAVA PARCELADOCUMENTO
         Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
         Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(txtCod.Text) & ","
         Sql = Sql & .TextMatrix(x, 1) & "," & .TextMatrix(x, 2) & "," & nSeq & "," & .TextMatrix(x, 4) & ","
         Sql = Sql & .TextMatrix(x, 5) & "," & .TextMatrix(x, 8) & ")"
         cn.Execute Sql, rdExecDirect
    End With

    'DELETA TEMPORARIO
    'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDeLogin & "'"
    Sql = "DELETE FROM DAM WHERE SID=" & nSid
    cn.Execute Sql, rdExecDirect
    
    sCodReduz = Format(Val(txtCod.Text), "000000")
    sNomeResp = lblProp.Caption
    sTipoImposto = "TESTE3"
    sEndImovel = lblRua.Caption
    nNumImovel = lblNumImovel.Caption
    sComplImovel = lblCompl.Caption
    sBairroImovel = lblBairro.Caption
    nCodLogr = 0
    
    sEndEntrega = lblRuaEntrega.Caption
    nNumEntrega = lblNumEntrega.Caption
    sBairroEntrega = lblBairroEntrega.Caption
    sComplEntrega = lblCompl.Caption
    sCepEntrega = lblCepEntrega.Caption
    sCidadeEntrega = lblCidadeEntrega.Caption
    sUFEntrega = lblUF.Caption
    
    Set qd.ActiveConnection = cn
    
    nAno = grdTemp.TextMatrix(1, 1)
    nCodLanc = grdTemp.TextMatrix(1, 2)
    nSeq = grdTemp.TextMatrix(1, 3)
    nNumParc = grdTemp.TextMatrix(1, 4)
    nComplemento = grdTemp.TextMatrix(1, 5)
    dDataVencto = CDate(sVencimento)
    sValorParc = nValorTotal + nValorTxExpParc
    nValorTotal = nValorTotal + nValorTxExpParc
    nNumDoc = grdTemp.TextMatrix(1, 8)
    'NumBarra2 = Gera2of5Cod(sValorParc, dDataVencto, nNumDoc & RetornaDVNumDoc(nNumDoc), nNumParc, nCodLanc, nSeq, nComplemento)
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)
    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))

    Sql = "INSERT DAM(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,CPF,ENDERECO,NUMERO,COMPLEMENTO,"
    Sql = Sql & "BAIRRO,CIDADE,UF,QUADRA,LOTE,FULLLANC,FULLTRIB,NUMDAM,ANOEXERC,LANC,NUMSEQ,NUMPARCELA,COMP,DATAVENCTO,"
    Sql = Sql & "SIT,AJ,DA,PRINCIPAL,CORRECAO,MULTA,JUROS,TOTAL,STRBARRA2,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
    Sql = Sql & "VALORDAM,VALORPRINCDAM,CODTRIBUTO,USUARIO,SID) VALUES('" & NomeDeLogin & "'," & x & ",'" & sNumInsc & "','"
    Sql = Sql & Format(nCodReduz, "000000") & "','" & "DAM" & "','" & Left$(Mask(sNomeResp), 40) & "','" & sCPF & "','" & Left$(Mask(sEndImovel), 40) & "',"
    Sql = Sql & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left$(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','"
    Sql = Sql & Left$(sQuadra, 10) & "','" & Left$(sLote, 10) & "','" & "ITBI" & "','" & Left$(sFullTrib, 2000) & "','"
    Sql = Sql & CStr(nNumDoc) & CStr(RetornaDVNumDoc(nNumDoc)) & "','" & nAno & "','" & nCodLanc & "','"
    Sql = Sql & nSeq & "','" & nNumParc & "','" & nComplemento & "','" & Format(dDataVencto, "mm/dd/yyyy") & "','"
    Sql = Sql & 3 & "','" & "N" & "','" & "N" & "'," & Virg2Ponto(sTr(nValorParc)) & ","
    Sql = Sql & Virg2Ponto(0) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & "," & Virg2Ponto(CStr(nValorParc)) & ",'"
    Sql = Sql & Mask(StrBarra2) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "'," & Virg2Ponto(CStr(nValorTotal)) & ","
    Sql = Sql & Virg2Ponto(CStr(nValorTotal)) & "," & 0 & ",'" & NomeDeLogin & "'," & nSid & ")"
    cn.Execute Sql, rdExecDirect
    
    modLg "Emissão de ITBI nº " & nNumDoc & " Código: " & nCodReduz & " - " & sNomeResp
    
    If nValorTxExpParc > 0 Then
        Sql = "INSERT DAM(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,CPF,ENDERECO,NUMERO,COMPLEMENTO,"
        Sql = Sql & "BAIRRO,CIDADE,UF,QUADRA,LOTE,FULLLANC,FULLTRIB,NUMDAM,ANOEXERC,LANC,NUMSEQ,NUMPARCELA,COMP,DATAVENCTO,"
        Sql = Sql & "SIT,AJ,DA,PRINCIPAL,CORRECAO,MULTA,JUROS,TOTAL,STRBARRA2,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
        Sql = Sql & "VALORDAM,VALORPRINCDAM,CODTRIBUTO,USUARIO,SID) VALUES('" & NomeDeLogin & "'," & x & ",'" & sNumInsc & "','"
        Sql = Sql & Format(nCodReduz, "000000") & "','" & "DAM" & "','" & Left$(Mask(sNomeResp), 40) & "','" & sCPF & "','" & Left$(Mask(sEndImovel), 40) & "',"
        Sql = Sql & nNumImovel & ",'" & Left(Mask(sComplImovel), 30) & "','" & Left$(Mask(sBairroImovel), 25) & "','" & Mask(sCidadeEntrega) & "','" & sUFEntrega & "','"
        Sql = Sql & Left$(sQuadra, 10) & "','" & Left$(sLote, 10) & "','" & "TAXA EXP.DOC." & "','" & "TAXA EXP.DOC." & "','"
        Sql = Sql & CStr(nNumDoc) & CStr(RetornaDVNumDoc(nNumDoc)) & "','" & nAno & "','" & 4 & "','"
        Sql = Sql & nSeq & "','" & nNumParc & "','" & nComplemento & "','" & Format(dDataVencto, "mm/dd/yyyy") & "','"
        Sql = Sql & 3 & "','" & "N" & "','" & "N" & "'," & Virg2Ponto(sTr(nValorTxExpParc)) & ","
        Sql = Sql & Virg2Ponto(0) & "," & Virg2Ponto(0) & "," & Virg2Ponto(0) & "," & Virg2Ponto(CStr(nValorTxExpParc)) & ",'"
        Sql = Sql & Mask(StrBarra2) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "'," & Virg2Ponto(CStr(nValorTotal)) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorTotal)) & "," & 3 & ",'" & NomeDeLogin & "'," & nSid & ")"
        cn.Execute Sql, rdExecDirect
    End If
    
    frmReport.ShowReport "DAM", frmMdi.HWND, Me.HWND, nSid
Else
    Sql = "INSERT ITBI (CODREDUZIDO,NOME,CPF,DATAVENCTO,DESCRICAO,VALORTOTAL,ARTIGO,TIPO,FUNCIONARIO,TIPODOC) VALUES("
    Sql = Sql & nCodReduz & ",'" & lblProp.Caption & "','" & sCPF & "','" & Format(mskVencto.Text, "mm/dd/yyyy") & "','"
    Sql = Sql & Mask(txtDesc.Text) & "'," & Virg2Ponto(txtValor.Text) & ",'" & txtArtigo.Text & "','" & txtTipo.Text & "','"
    Sql = Sql & txtFunc.Text & "'," & IIf(Opt(0).value, 0, 1) & ")"
    cn.Execute Sql, rdExecDirect
    
    frmReport.ShowReport "ITBI", frmMdi.HWND, Me.HWND, nLastCod + 1
End If

On Error Resume Next
'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSid
cn.Execute Sql, rdExecDirect

bGerado = True
Limpa
Exit Sub

Erro:
For z = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(z).Description
Next
Resume Next
End Sub

Private Sub EmiteBoleto()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataDam As String, sDataVencto As String
Dim nCodReduz As Long, sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, sQuadras As String, sLotes As String, nSeqLanc As Integer
Dim sUsuario As String, nNumDoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String, bGerado As Boolean
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSid As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date
Dim sBarra As String, sDigitavel2 As String, nValorDam As Double, nValorPrincDam As Double, nNumGuia As Long, nValorParc As Double
Dim sTipoEnd As String, bBoleto As Boolean


bBoleto = False
If MsgBox("Confirma Emissão do ITBI ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nSid = Int(Rnd(100) * 1000000)

Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect

'RETORNA VALOR EXPEDIENTE
'Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'nValorTaxa = RdoAux!VALORDAM
'RdoAux.Close
nValorTaxa = 0
bMulta = False
sDoc = ""
nPos = 0
nValorDoc = 0
nValorDam = 0
nValorParc = CDbl(txtValor.Text)
nValorPrincDam = 0
nCodReduz = Val(txtCod.Text)
sUsuario = NomeDeLogin

Select Case nCodReduz
    
    Case 1 To 99999
        Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !Inscricao
            sNome = !NomeCidadao
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
                If sDoc = "" Then
                    sDoc = SubNull(!rg)
                End If
            End If
            sEnd = SubNull(!Logradouro)
            nNum = Val(SubNull(!Li_Num))
            sCompl = Left(SubNull(!Li_Compl), 30)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!descCidade)
            sUF = SubNull(!li_uf)
            sQuadras = Left(SubNull(!Li_Quadras), 15)
            sLotes = Left(SubNull(!Li_Lotes), 10)
           .Close
        End With
     Case 500000 To 800000

        Sql = "SELECT CODCIDADAO,CODBAIRRO,CODBAIRRO2 FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If Val(SubNull(RdoAux!CodBairro)) > 0 Then
           sTipoEnd = "R"
        Else
           If Val(SubNull(RdoAux!CodBairro2)) > 0 Then
              sTipoEnd = "C"
           Else
              sTipoEnd = "R"
           End If
        End If
        RdoAux.Close
               
        If sTipoEnd = "R" Then
             Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
             Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
             Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
             Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        Else
             Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
             Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
             Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
             Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        End If
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNome = !NomeCidadao
                If SubNull(!Cnpj) <> "" Then
'                    If Val(!Cnpj) > 0 Then
                        sDoc = Format(!Cnpj, "0#\.###\.###/####-##")
 '                   End If
                Else
                    If SubNull(!CPF) <> "" Then
'                        If Val(!CPF) > 0 Then
                            sDoc = Format(RetornaNumero(!CPF), "00#\.###\.###-##")
 '                       End If
                    End If
                End If
                
                If Val(SubNull(!FCodLogradouro)) > 0 Then
                    Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                    Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                    Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                    Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                    Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoS
                        If .RowCount > 0 Then
                           sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                        Else
                           sEnd = ""
                        End If
                       .Close
                    End With
                Else
                   sEnd = SubNull(!FNomeLogradouro)
                End If
                nNum = SubNull(RdoAux!fNUMIMOVEL)
                sCompl = SubNull(RdoAux!fcomplemento)
                
                If Trim(!fsiglauf) <> "" Then
                    Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                    Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                    If RdoS.RowCount > 0 Then
                        sCidade = RdoS!descCidade
                    Else
                        sCidade = ""
                    End If
                End If
                If Not IsNull(RdoAux!fCodBairro) Then
                    Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                    Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                    If .RowCount > 0 Then
                        sBairro = RdoS!DescBairro
                    Else
                        sBairro = ""
                    End If
                Else
                    sBairro = ""
                End If
                sUF = SubNull(!fsiglauf)
                sCep = SubNull(!FCEP)
                   
            sQuadras = ""
            sLotes = ""
            End If
        End With
End Select

'MONTA TRIBUTOS
sDataVencto = mskVencto.Text
nAno = Year(CDate(sDataVencto))
'VERIFICA PRÓXIMA SEQUENCIA DE LANÇAMENTO
Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
Sql = Sql & " AND CODLANCAMENTO=36 AND ANOEXERCICIO=" & nAno
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=36 AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & 1
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!SEQMAXIMA) Then
           nSeq = 0
        Else
           nSeq = RdoAux2!SEQMAXIMA + 1
        End If
    Else
        nSeq = 0
    End If
   .Close
End With

'GRAVA DEBITOPARCELA
'Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
'Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USUARIO) VALUES(" & nCodReduz & "," & nAno & ",36," & nSeq & ",1,0,3,'"
'Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "','" & NomeDeLogin & "')"
Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES(" & nCodReduz & "," & nAno & ",36," & nSeq & ",1,0,3,'"
Sql = Sql & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect
sFullTrib = ""
         
'GRAVA HISTORICOPARCELA
Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
Sql = Sql & " AND CODLANCAMENTO=36 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=1 AND CODCOMPLEMENTO=0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If IsNull(!maximo) Then
        nSeqLanc = 1
    Else
        nSeqLanc = !maximo + 1
    End If
   .Close
End With
Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES("
Sql = Sql & nCodReduz & "," & nAno & ",36," & nSeq & ",1,0," & nSeqLanc & ",'" & Mask(txtDesc.Text) & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect
sFullTrib = Mask(txtDesc.Text)
        
If chkRural.value = vbUnchecked Then
    'GRAVA historico IMÓVEL
    Sql = "SELECT max(SEQ) as MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(txtImovel.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqLanc = 1
        Else
            nSeqLanc = !maximo + 1
        End If
       .Close
    End With
    sHist = "Emissão de ITBI no código cidadão " & CStr(nCodReduz)
    Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST,DATAHIST2,USERID) VALUES("
    Sql = Sql & Val(txtImovel.Text) & "," & nSeqLanc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sHist & "','" & Format(Now, "mm/dd/yyyy") & "'," & 236 & ")"
    cn.Execute Sql, rdExecDirect
End If
        
        
'GRAVA DEBITOTRIBUTO
Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
Sql = Sql & nCodReduz & "," & nAno & ",36," & nSeq & ",1,0,84," & Virg2Ponto(CStr(nValorParc)) & ")"
cn.Execute Sql, rdExecDirect

'grava documento
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
   nNumDoc = 0
Else
   nNumDoc = RdoAux!maximo + 1
End If
RdoAux.Close

Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,EMISSOR,VALORGUIA) VALUES("
Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & ",'" & NomeDeLogin & "'," & Virg2Ponto(CStr(Round(nValorParc, 2))) & ")"
cn.Execute Sql, rdExecDirect

'GRAVA PARCELADOCUMENTO
Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
Sql = Sql & nAno & ",36," & nSeq & ",1,0," & nNumDoc & ")"
cn.Execute Sql, rdExecDirect

sNumDoc = CStr(nNumDoc) & "-" & RetornaDVNumDoc(nNumDoc)
sNumDoc2 = CStr(nNumDoc) & RetornaDVNumDoc(nNumDoc)
sNumDoc3 = CStr(nNumDoc) & Modulo11(nNumDoc)
    
Sql = "insert boleto(usuario,computer,sid,seq,inscricao,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,quadra,lote,numdoc,nomefunc,datadam,fulllanc,fulltrib,"
Sql = Sql & "anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datavencto,aj,da,principal,juros,multa,correcao,total,numdoc2,valordam) values('"
Sql = Sql & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & 0 & ",'" & sInsc & "'," & nCodReduz & ",'" & Left(Mask(sNome), 40) & "','" & sDoc & "','"
Sql = Sql & Mask(sEnd) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Left(Mask(sCidade), 25) & "','" & sUF & "','" & Mask(sQuadras) & "','"
Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & "ITBI" & "','" & sFullTrib & "'," & nAno & ","
Sql = Sql & 36 & "," & nSeq & "," & 1 & "," & 0 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sAj & "','" & sDA & "'," & Virg2Ponto(Format(nValorParc, "#0.00")) & ","
Sql = Sql & Virg2Ponto(Format(0, "#0.00")) & "," & Virg2Ponto(Format(0, "#0.00")) & "," & Virg2Ponto(Format(0, "#0.00")) & "," & Virg2Ponto(Format(nValorParc, "#0.00")) & ",'" & sNumDoc2
Sql = Sql & "'," & Virg2Ponto(Format(nValorParc + nValorTaxa, "#0.00")) & ")"
cn.Execute Sql, rdExecDirect
    
Sql = "insert boleto(usuario,computer,sid,seq,inscricao,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,quadra,lote,numdoc,nomefunc,datadam,fulllanc,fulltrib,"
Sql = Sql & "anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datavencto,aj,da,principal,juros,multa,correcao,total,numdoc2,valordam) values('"
Sql = Sql & NomeDeLogin & "','" & NomeDoComputador & "'," & nSid & "," & 1 & ",'" & sInsc & "'," & nCodReduz & ",'" & Left(Mask(sNome), 40) & "','" & sDoc & "','"
Sql = Sql & Mask(sEnd) & "'," & nNum & ",'" & Left(Mask(sCompl), 30) & "','" & Left(Mask(sBairro), 25) & "','" & Left(Mask(sCidade), 25) & "','" & sUF & "','" & Mask(sQuadras) & "','"
Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataVencto, "mm/dd/yyyy") & "','" & "ITBI" & "','003-Taxa de Expediente'," & nAno & ","
Sql = Sql & 4 & "," & nSeq & "," & 1 & "," & 0 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sAj & "','" & sDA & "'," & Virg2Ponto(Format(nValorTaxa, "#0.00")) & ","
Sql = Sql & Virg2Ponto(Format(0, "#0.00")) & "," & Virg2Ponto(Format(0, "#0.00")) & "," & Virg2Ponto(Format(0, "#0.00")) & "," & Virg2Ponto(Format(nValorTaxa, "#0.00")) & ",'" & sNumDoc2
Sql = Sql & "'," & Virg2Ponto(Format(nValorParc + nValorTaxa, "#0.00")) & ")"
'cn.Execute Sql, rdExecDirect
    
nValorDoc = nValorParc + nValorTaxa
sDataDam = mskVencto.Text
bBoleto = False
If bBoleto Then
    '**** GERADOR DE CÓDIGO DE BARRAS ********
    sNossoNumero = "2678478"
    sDigitavel = "001900000"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv & "0" & sNossoNumero & "01"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv & Right(sNumDoc3, 8) & "18"
    sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
    sDigitavel = sDigitavel & sDv
    
    dDataBase = "07/10/1997"
    nFatorVencto = CDate(sDataDam) - dDataBase
    sQuintoGrupo = Format(nFatorVencto, "0000")
    sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")
    sBarra = "0019" & Format(nFatorVencto, "0000") & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000") & "00000026784780"
    sBarra = sBarra & sNumDoc3 & "18"
    sDv = Trim(Calculo_DV11(sBarra))
    sBarra = Left(sBarra, 4) & sDv & Mid(sBarra, 5, Len(sBarra) - 4)
    
    sDigitavel = sDigitavel & sDv & sQuintoGrupo
    
    sDigitavel2 = Left(sDigitavel, 5) & "." & Mid(sDigitavel, 6, 5) & " " & Mid(sDigitavel, 11, 5) & "." & Mid(sDigitavel, 16, 6) & " "
    sDigitavel2 = sDigitavel2 & Mid(sDigitavel, 22, 5) & "." & Mid(sDigitavel, 27, 6) & " " & Mid(sDigitavel, 33, 1) & " " & Right(sDigitavel, 14)
    sBarra = Gera2of5Str(sBarra)
    Sql = "update boleto set digitavel='" & sDigitavel2 & "',codbarra='" & Mask(sBarra) & "',valorprincdam=" & Virg2Ponto(Format(nValorPrincDam, "#0.00")) & " where sid=" & nSid
    cn.Execute Sql, rdExecDirect

Else
    Dim sValor As String, dDataVencto As Date, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, NumBarra2c As String, NumBarra2d As String, StrBarra2 As String
    sValor = nValorDoc
    dDataVencto = CDate(sDataDam)
    'nNumDoc = Val(sNumDoc2)
 
    NumBarra2 = Gera2of5Cod(sValor, dDataVencto, nNumDoc, nCodReduz)
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)

    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
    Sql = "update boleto set numbarra2a='" & NumBarra2a & "',numbarra2b='" & NumBarra2b & "',numbarra2c='" & NumBarra2c & "',numbarra2d='" & NumBarra2d & "',codbarra='" & Mask(StrBarra2) & "' where sid=" & nSid
    cn.Execute Sql, rdExecDirect
    
    nNumGuia = nNumDoc
 '   If frmMdi.frTeste.Visible = False Then
 '       frmReport.ShowReport2 "BOLETODAM_V4", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
 '   Else
   '     frmReport.ShowReport2 "BOLETODAM_v4_TESTE", frmMdi.hwnd, Me.hwnd, nSid, nNumGuia
 '   End If
End If
    
    '*******************************************

nNumGuia = nNumDoc
If bBoleto Then
    frmReport.ShowReport2 "BOLETODAM", frmMdi.HWND, Me.HWND, nSid, nNumGuia
Else
    frmReport.ShowReport2 "BOLETODAM_V4", frmMdi.HWND, Me.HWND, nSid, nNumGuia
End If

Sql = "delete from boleto where sid=" & nSid
cn.Execute Sql, rdExecDirect
Limpa

End Sub

Private Sub EmiteBoletoRegistrado()
Dim v1 As String, v2 As String, v3 As String, v4 As String, v5 As String, v6 As String, v7 As String, v8 As String, v9 As String, V10 As String
Dim RdoAux As rdoResultset, Sql As String, nNumDoc As Long, nSeq As Integer, nCodReduz As Long, nLanc As Integer, nParc As Integer, nAno As Integer
Dim sDataBase As String, sDataVencto As String, nSeqObs As Integer, bAbateu As Boolean
Dim sNome As String, sEnd As String, nNum As Integer, nValorDoc As Double, sCompl As String, sBairro As String, sCidade As String, sUF As String, sDoc As String

nCodReduz = Val(txtCod.Text)
sDataBase = Right$(frmMdi.Sbar.Panels(6).Text, 10)
sDataVencto = mskVencto.Text
nLanc = 36
nParc = 1
nAno = Year(Now)

Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nNumDoc = RdoAux!maximo + 1
RdoAux.Close

Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,ISENTOMJ,PERCISENCAO,TIPODOC,emissor,registrado) VALUES("
Sql = Sql & nNumDoc & ",'" & Format(Now, sDataFormat) & "'," & 0 & "," & 0 & "," & 0 & "," & 0 & "," & 0 & ",3,'" & NomeDeLogin & "',1)"
cn.Execute Sql, rdExecDirect

Sql = "SELECT MAX(SEQLANCAMENTO) AS SEQMAXIMA FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND CODLANCAMENTO=" & nLanc & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux2!SEQMAXIMA) Then
   nSeq = 0
Else
   nSeq = RdoAux2!SEQMAXIMA + 1
End If

Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & ","
Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & nNumDoc & ")"
cn.Execute Sql, rdExecDirect

Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','"
Sql = Sql & Format(sDataBase, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect

Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
Sql = Sql & nCodReduz & "," & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & ","
Sql = Sql & 84 & "," & Virg2Ponto(txtValor.Text) & ")"
cn.Execute Sql, rdExecDirect

If Trim(txtDesc.Text) <> "" Then
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=0"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqObs = 1
        Else
            nSeqObs = !maximo + 1
        End If
       .Close
    End With
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & ","
    Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & 0 & "," & nSeqObs & ",'" & Mask(Trim(txtDesc.Text)) & "',"
    Sql = Sql & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(sDataBase, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
End If

If chkRural.value = vbUnchecked Then
    'GRAVA historico IMÓVEL
    Sql = "SELECT max(SEQ) as MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & Val(txtImovel.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeqLanc = 1
        Else
            nSeqLanc = !maximo + 1
        End If
       .Close
    End With
    sHist = "Emissão de ITBI no código cidadão " & CStr(nCodReduz)
    Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST,DATAHIST2,USERID) VALUES("
    Sql = Sql & Val(txtImovel.Text) & "," & nSeqLanc & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sHist & "','" & Format(Now, "mm/dd/yyyy") & "'," & 236 & ")"
    cn.Execute Sql, rdExecDirect
End If


v1 = lblProp.Caption
v2 = Left(lblRua.Caption & ", " & lblNum.Caption & IIf(lblCompl.Caption <> "", " " & lblCompl.Caption, "") & " - " & lblBairro.Caption, 60)
v3 = mskVencto.Text
v4 = RetornaNumero(lblNumInsc.Caption)
v5 = "287353200" & Format(nNumDoc, "00000000")
v6 = RetornaNumero(txtValor.Text)
v7 = Left("JABOTICABAL", 18)
v8 = "SP"
v9 = lblCep.Caption
V10 = NomeDeLogin
If Trim(lblCep.Caption) = "" Or Trim(lblCep.Caption) = "-" Then
    v9 = "14870-000"
End If
ShellExecute HWND, "open", "http://sistemas.jaboticabal.sp.gov.br/gti/Pages/boletoBB.aspx?f1=" & v1 & "&f2=" & v2 & "&f3=" & v3 & "&f4=" & v4 & "&f5=" & v5 & "&f6=" & v6 & "&f7=" & v7 & "&f8=" & v8 & "&f9=" & v9 & "&f10=" & V10, vbNullString, vbNullString, conSwNormal

Limpa

End Sub


