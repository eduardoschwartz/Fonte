VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCertidao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de certidões ao contribuinte"
   ClientHeight    =   3705
   ClientLeft      =   5085
   ClientTop       =   2835
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   8985
   Visible         =   0   'False
   Begin VB.CheckBox chkAss 
      Caption         =   "Ocultar assinatura"
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   3375
      Width           =   2220
   End
   Begin VB.TextBox txtArea 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2670
      TabIndex        =   21
      Top             =   2895
      Width           =   1695
   End
   Begin VB.TextBox txtProcDem 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      TabIndex        =   19
      Top             =   2895
      Width           =   1695
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      Top             =   570
      Width           =   1695
   End
   Begin prjChameleon.chameleonButton cmdLoadDeb 
      Height          =   300
      Left            =   7740
      TabIndex        =   14
      ToolTipText     =   "Carregar os Dados"
      Top             =   900
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "&Carregar"
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
      MICON           =   "frmCertdao.frx":0000
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
      Left            =   7530
      TabIndex        =   15
      ToolTipText     =   "Sair da Tela"
      Top             =   3285
      Width           =   1335
      _ExtentX        =   2355
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
      MICON           =   "frmCertdao.frx":001C
      PICN            =   "frmCertdao.frx":0038
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
      Height          =   345
      Left            =   6120
      TabIndex        =   16
      ToolTipText     =   "Gera as guias informadas"
      Top             =   3285
      Width           =   1350
      _ExtentX        =   2381
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmCertdao.frx":00A6
      PICN            =   "frmCertdao.frx":00C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   8400
      TabIndex        =   18
      ToolTipText     =   "Consulta Cidadão"
      Top             =   1290
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmCertdao.frx":021C
      PICN            =   "frmCertdao.frx":0238
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7200
      TabIndex        =   1
      Top             =   570
      Width           =   1695
   End
   Begin VB.Label lblPercIsencao 
      Height          =   255
      Left            =   3375
      TabIndex        =   39
      Top             =   4815
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblDataProcIsencao 
      Height          =   255
      Left            =   1035
      TabIndex        =   38
      Top             =   5085
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblProcIsencao 
      Height          =   255
      Left            =   315
      TabIndex        =   37
      Top             =   5085
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblVVI 
      Height          =   255
      Left            =   3195
      TabIndex        =   36
      Top             =   4545
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label lblVVC 
      Height          =   255
      Left            =   2160
      TabIndex        =   35
      Top             =   4545
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblVVT 
      Height          =   255
      Left            =   945
      TabIndex        =   34
      Top             =   4545
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblInscricao 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6525
      TabIndex        =   33
      Top             =   4095
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblLote 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   5715
      TabIndex        =   32
      Top             =   4635
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblQuadra 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   5670
      TabIndex        =   31
      Top             =   4950
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label lblComplemento 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2565
      TabIndex        =   30
      Top             =   4095
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblNum 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7965
      TabIndex        =   29
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label lblDataProc 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   270
      TabIndex        =   28
      Top             =   4095
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Nº do CPF/CNPJ........:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   27
      Top             =   2580
      Width           =   2505
   End
   Begin VB.Label lblCPF 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   26
      Top             =   2580
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Bairro........:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   25
      Top             =   2268
      Width           =   2505
   End
   Begin VB.Label lblBairro 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   24
      Top             =   2268
      Width           =   6135
   End
   Begin VB.Label lblArea 
      Caption         =   "Área Demolida.........:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   2925
      Width           =   2505
   End
   Begin VB.Label lblProc 
      Caption         =   "Processo Demolição....:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4650
      TabIndex        =   22
      Top             =   2925
      Width           =   2505
   End
   Begin VB.Label lblCodCert 
      Height          =   255
      Left            =   270
      TabIndex        =   17
      Top             =   4545
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblRequerente 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   13
      Top             =   1332
      Width           =   5655
   End
   Begin VB.Label lblProp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   12
      Top             =   1644
      Width           =   6165
   End
   Begin VB.Label lblEnd 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   11
      Top             =   1950
      Width           =   5130
   End
   Begin VB.Label lblCertidao 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   10
      Top             =   1020
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label lblTipo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   9
      Top             =   240
      Width           =   6165
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo de Certidão......:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço do Imóvel....:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   7
      Top             =   1956
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Código/Inscrição......:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   4650
      TabIndex        =   6
      Top             =   600
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Proprietário/Empresa..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1644
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Requerente (Cidadão)..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1332
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Processo nº c/digito..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Nº da Certidão........:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   2505
   End
End
Attribute VB_Name = "frmCertidao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'PARAMETROS
Dim bCanPrint As Boolean
Dim nNumeroImovel As Integer
Dim bTemPredial As Boolean
Dim bFracaoIdeal As Boolean
Dim nAreaTerreno As Double
Dim nAreaPrincipal As Double
Dim nCodAgrupamento As Integer
Dim nValorAgrupamento As Double
Dim nNumTestadas As Integer
Dim nTestadaPrincipal As Double
Dim nCodGleba As Integer
Dim nFatorGleba As Double
Dim nCodProfundidade As Integer
Dim nValorProfundidade As Double
Dim nFatorProfundidade As Double
Dim nCodSituacao As Integer
Dim nFatorSituacao As Double
Dim nCodPedologia As Integer
Dim nFatorPedologia As Double
Dim nCodTopografia As Integer
Dim nFatorTopografia As Double
Dim nFatorDistrito As Double
Dim nValorFatores As Double
Dim nFatorCategoria As Double
Dim nValorVenalTerritorial As Double
Dim nValorVenalPredial As Double
Dim nValorVenalImovel As Double
Dim nTaxaLimpeza As Double, nTaxaConservacao As Double
'GERAL
Dim nCodReduz As Long
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim Sql As String
Dim nAnoCalculo As Integer
'TIPOS
Private Type PROFUNDIDADE
    Distrito As Integer
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORPROFUN
    Distrito As Integer
    Codigo As Integer
    Fator As Double
End Type
Private Type GLEBA
    Codigo As Integer
    Min As Double
    Max As Double
End Type
Private Type FATORCATEG
    Uso As Integer
    Tipo As Integer
    Categoria As Integer
    Fator As Double
End Type

'MATRIZES
Dim aFatorD() As Double
Dim aFatorP() As Double
Dim aFatorT() As Double
Dim aFatorS() As Double
Dim aFatorG() As Double
Dim aFatorR() As Double
Dim aProf() As PROFUNDIDADE
Dim aFatorF() As FATORPROFUN
Dim aFatorC() As FATORCATEG
Dim aGleba() As GLEBA

Dim xImovel As clsImovel
Dim sTipo As String, bIsento65 As Boolean, nAreaIsento As Double, sTextoIsento As String, bImune As Boolean

Private Sub cmdCnsImovel_Click()
Set frm = frmCnsCidadao
frm.sForm = "frmCertidao"
frm.show
frm.ZOrder 0

End Sub

Private Sub cmdLoadDeb_Click()
Dim bResidencial As Boolean, bResideImovel As Boolean
Limpa
If Trim$(txtNumProc.Text) = "" Then
    MsgBox "Digite o nº do Processo.", vbCritical, "Atenção"
    txtNumProc.SetFocus
    Exit Sub
End If

If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) = 0 Then
    MsgBox "Nº do processo inválido." & vbCrLf & "Formato deve ser: Nº do Processo/Ano", vbCritical, "Atenção"
    txtNumProc.SetFocus
    Exit Sub
End If

If Not IsNumeric(Right$(txtNumProc.Text, 4)) Then
    MsgBox "Nº do processo inválido." & vbCrLf & "O ano deve ter 4 digitos", vbCritical, "Atenção"
    txtNumProc.SetFocus
    Exit Sub
End If

nCodReduz = Val(txtCod.Text)
nNumeroImovel = 0
If nCodReduz < 100000 Then
   Sql = "SELECT CODREDUZIDO,SETOR,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & nCodReduz
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
        If .RowCount > 0 Then
            With xImovel
                If RdoAux!Inativo Then
                    MsgBox "Este imóvel encontra-se inativo.", vbCritical, "Atenção"
                    lblPercIsencao.Caption = 0
                    Limpa
                    Exit Sub
                End If
                
                .CarregaImovel nCodReduz
                nNumeroImovel = .Li_Num
                lblProp.Caption = .NomePropPrincipal
                lblQuadra.Caption = .Li_Quadras
                bResideImovel = .ResideImovel
                lblLote.Caption = .Li_Lotes
                lblInscricao.Caption = .Inscricao
                'lblEnd.Caption = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro & ", " & .Li_Num & " " & .Li_Compl
                lblEnd.Caption = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro
                lblNum.Caption = .Li_Num
                lblComplemento.Caption = .Li_Compl
                lblBairro.Caption = .DescBairro
                Sql = "SELECT CODCIDADAO,CPF,CNPJ FROM CIDADAO WHERE CODCIDADAO=" & .CodPropPrincipal
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                With RdoAux2
                    If .RowCount > 0 Then
                        If Not IsNull(!CPF) Then
                            lblCPF.Caption = Format(Trim(!CPF), "000\.000\.000-00")
                        Else
                            If Not IsNull(!Cnpj) Then
                                lblCPF.Caption = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
                            End If
                        End If
                    End If
                   .Close
                End With
            End With
        Else
            MsgBox "Imóvel não Cadastrado.", vbExclamation, "Atenção"
            Exit Sub
        End If
       .Close
   End With
ElseIf nCodReduz > 100000 And nCodReduz < 500000 Then
   Sql = "SELECT CODIGOMOB,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,NOMELOGR FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
        If .RowCount > 0 Then
            lblProp.Caption = !RazaoSocial
            If SubNull(!NomeLogradouro) = "" Then
                lblEnd.Caption = Trim$(SubNull(!NomeLogr)) & ", " & !Numero
            Else
                lblEnd.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & ", " & !Numero
            End If
        Else
            MsgBox "Empresa não Cadastrada.", vbExclamation, "Atenção"
            Exit Sub
        End If
       .Close
   End With
Else
    Sql = "SELECT NOMECIDADAO, ABREVTIPOLOG, ABREVTITLOG, NOMELOGRADOURO, NOMELOGRADOURO2, NUMIMOVEL, COMPLEMENTO "
    Sql = Sql & "From vwCIDADAO Where CodCidadao =" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         If .RowCount > 0 Then
             lblProp.Caption = !NomeCidadao
             If Not IsNull(!NomeLogradouro) Then
                If !NomeLogradouro <> "" Then
                    lblEnd.Caption = Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL))
                Else
                    lblEnd.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL
                End If
             Else
                lblEnd.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL
             End If
         Else
             MsgBox "Cidadão não Cadastrado.", vbExclamation, "Atenção"
             Exit Sub
         End If
        .Close
    End With
End If

If lblCodCert.Caption = 1 Then
    VerificaDebito
ElseIf lblCodCert.Caption = 6 Then
    bImune = False
    If Not bResideImovel Then
        MsgBox "Proprietario não reside no imóvel", vbInformation, "Atenção"
        GoTo FimIsento
    End If
    Sql = "select codreduzido,imune from cadimob where codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux!Imune = True Then
        bImune = True
        sTipo = "IPTU"
        lblPercIsencao.Caption = "100"
        RdoAux.Close
        GoTo FimIsento
    End If

    bIsento65 = False
    
    bResidencial = True
    Sql = "SELECT * FROM VWISENCAOprocesso WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
          Sql = "SELECT * FROM vwISENCAOPROCESSO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODISENCAO=1 AND NUMPROCESSO IS NOT NULL"
          Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          If RdoAux2.RowCount = 0 Then
            'BUSCA ÁREA PRINCIPAL
             'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
             Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
             Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux2
                 Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
                 Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux3
                     If Not IsNull(!soma) Then
                         If !soma <= 65 And RdoAux2!USOCONSTR = 1 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                            Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                            If RdoAux4.RowCount > 0 Then
                                cmdPrint.Enabled = False
                                MsgBox "Proprietario possue mais de um imóvel", vbInformation, "Atenção"
                                RdoAux4.Close
                                Exit Sub
                            End If
                            RdoAux4.Close
                            sTipo = "IPTU"
                            bIsento65 = True
                            nAreaIsento = FormatNumber(!soma, 2)
                            lblVVC.Caption = nAreaIsento
                         Else
                            MsgBox "Este Imóvel/Empresa não possue isenção para este ano.", vbExclamation, "Atenção"
                            cmdPrint.Enabled = False
                            Exit Sub
                         End If
                     End If
                    .Close
                 End With
                 Do Until .EOF
                    If !USOCONSTR > 1 Then
                        bResidencial = False
                        Exit Do
                    End If
                   .MoveNext
                 Loop
             End With
             If Not bResidencial Then
                MsgBox "Este Imóvel possui área não residencial.", vbExclamation, "Atenção"
                cmdPrint.Enabled = False
                Exit Sub
             End If
             
           Else
                If Not IsNull(RdoAux2!numprocesso) Then
                   lblProcIsencao.Caption = RdoAux2!numprocesso
                   lblDataProcIsencao.Caption = Format(RdoAux2!DATAPROCESSO, "dd/mm/yyyy")
                   'lblPercIsencao.Caption = RdoAux2!percisencao
                   sTextoIsento = "Isento de acordo com o Processo nº " & RdoAux2!numprocesso & " de " & Format(RdoAux2!DATAPROCESSO, "dd/mm/yyyy") & "."
                Else
                   sTextoIsento = ""
                End If
           End If
        Else
            If Not IsNull(!numprocesso) Then
               lblProcIsencao.Caption = !numprocesso
               lblDataProcIsencao.Caption = Format(!DATAPROCESSO, "dd/mm/yyyy")
               lblPercIsencao.Caption = !percisencao
               sTextoIsento = "Isento de acordo com o Processo nº " & !numprocesso & " de " & Format(!DATAPROCESSO, "dd/mm/yyyy") & "."
            Else
               sTextoIsento = ""
            End If
            If nCodReduz < 100000 Then
                sTipo = "IPTU"
            Else
                sTipo = "ISS"
            End If
        End If
    End With
FimIsento:
ElseIf lblCodCert.Caption = 7 And nNumeroImovel = 0 Then
    MsgBox "Este Imóvel não possue Número. Verifique junto ao Cadastro Técnico.", vbExclamation, "Atenção"
    cmdPrint.Enabled = False
    Exit Sub
End If
cmdPrint.Enabled = True

End Sub

Private Sub VerificaDebito()
Dim x As Integer, bAchou As Boolean, sDescReduz As String
Dim aTipo() As String, bTemValor As Boolean
ReDim aTipo(0)

sTipo = ""
Sql = "SELECT DEBITOPARCELA.CODREDUZIDO, DEBITOPARCELA.ANOEXERCICIO, DEBITOPARCELA.CODLANCAMENTO, DEBITOPARCELA.SEQLANCAMENTO, "
Sql = Sql & "DEBITOPARCELA.NumParcela , DEBITOPARCELA.CODCOMPLEMENTO, LANCAMENTO.DESCREDUZ, Sum(DEBITOTRIBUTO.VALORTRIBUTO) AS VALORTOTAL, "
Sql = Sql & "DEBITOTRIBUTO.CODTRIBUTO FROM DEBITOPARCELA INNER JOIN LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO INNER JOIN "
Sql = Sql & "DEBITOTRIBUTO ON DEBITOPARCELA.CODREDUZIDO = DEBITOTRIBUTO.CODREDUZIDO AND DEBITOPARCELA.ANOEXERCICIO = DEBITOTRIBUTO.ANOEXERCICIO AND "
Sql = Sql & "DEBITOPARCELA.CODLANCAMENTO = DEBITOTRIBUTO.CODLANCAMENTO AND DEBITOPARCELA.SEQLANCAMENTO = DEBITOTRIBUTO.SEQLANCAMENTO AND "
Sql = Sql & "DEBITOPARCELA.NUMPARCELA = DEBITOTRIBUTO.NUMPARCELA AND DEBITOPARCELA.CODCOMPLEMENTO = DEBITOTRIBUTO.CODCOMPLEMENTO "
Sql = Sql & "Where (DEBITOPARCELA.STATUSLANC = 3) And (DateDiff(Day, DEBITOPARCELA.DATAVENCIMENTO, GETDATE()) > 0) "
Sql = Sql & "GROUP BY LANCAMENTO.DESCREDUZ, DEBITOPARCELA.CODREDUZIDO, DEBITOPARCELA.ANOEXERCICIO, DEBITOPARCELA.CODLANCAMENTO, "
Sql = Sql & "DEBITOPARCELA.SEQLANCAMENTO, DEBITOPARCELA.NUMPARCELA, DEBITOPARCELA.CODCOMPLEMENTO, LANCAMENTO.DESCREDUZ,"
Sql = Sql & "DEBITOTRIBUTO.CodTributo HAVING (DEBITOPARCELA.CODREDUZIDO = " & nCodReduz & ") AND (DEBITOPARCELA.NUMPARCELA > 0) AND (DEBITOPARCELA.CODLANCAMENTO <> 11) AND CODTRIBUTO <> 3"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then GoTo REPARC
    bTemValor = False
    Do Until .EOF
        If !CodLancamento = 1 And !AnoExercicio = 2003 Then
            bTemValor = False
        Else
            If !ValorTotal > 0 And !CodLancamento <> 20 Then
                bTemValor = True
                Exit Do
            End If
        End If
       .MoveNext
    Loop
    If Not bTemValor Then GoTo REPARC

    
    .MoveFirst
    Do Until .EOF
        If !CodLancamento = 5 Then
        
            If Not IsNull(RdoAux!ValorTotal) Then
                If RdoAux!ValorTotal = 0 Then
                    GoTo proximo
                End If
            End If
        End If
    
        bAchou = False
        sDescReduz = !descreduz
        If sDescReduz = "RECALCULO IPTU" Then sDescReduz = "ITU / IPTU"
        For x = 1 To UBound(aTipo)
            If aTipo(x) = sDescReduz Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aTipo(UBound(aTipo) + 1)
            aTipo(UBound(aTipo)) = sDescReduz
        End If
proximo:
       .MoveNext
    Loop
End With

If UBound(aTipo) > 0 Then
    For x = 1 To UBound(aTipo)
        sTipo = sTipo & aTipo(x) & "/ "
    Next
    sTipo = Left(sTipo, Len(sTipo) - 2)
    lblTipo.Caption = "CERTIDÃO DE DÉBITO POSITIVA"
Else
    lblTipo.Caption = "CERTIDÃO DE DÉBITO NEGATIVA"
    GoTo JULGAMENTO
End If

Exit Sub

REPARC:
Sql = "SELECT DEBITOPARCELA.*, LANCAMENTO.DESCREDUZ FROM DEBITOPARCELA INNER JOIN "
Sql = Sql & "LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND DEBITOPARCELA.CODLANCAMENTO = 20 AND STATUSLANC=3  AND DATAVENCIMENTO < '" & Format(Now, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        lblTipo.Caption = "CERTIDÃO DE DÉBITO POSITIVA"
    Else
        Sql = "SELECT DEBITOPARCELA.*, LANCAMENTO.DESCREDUZ FROM DEBITOPARCELA INNER JOIN "
        Sql = Sql & "LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO "
        Sql = Sql & "WHERE CODREDUZIDO=" & nCodReduz & " AND DEBITOPARCELA.CODLANCAMENTO = 20 AND STATUSLANC=3  AND DATAVENCIMENTO >= '" & Format(Now, "mm/dd/yyyy") & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                lblTipo.Caption = "CERTIDÃO DE DÉBITO POSITIVA COM EFEITO DE NEGATIVA"
            Else
                lblTipo.Caption = "CERTIDÃO DE DÉBITO NEGATIVA"
                GoTo JULGAMENTO
            End If
        End With
    End If
   .Close
End With
If Val(txtCod.Text) < 100000 Then
    sTipo = "ITU / IPTU"
Else
    sTipo = "ISS/TX.LIC/VIG.SANITARIA"
End If
Exit Sub

JULGAMENTO:
Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND (STATUSLANC=19 or STATUSLANC=20)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
    lblTipo.Caption = "CERTIDÃO DE DÉBITO POSITIVA COM EFEITO DE NEGATIVA"
End If

End Sub

Private Sub cmdPrint_Click()
Dim sTexto As String, sNomeParam As String, t As Integer
Dim sNomeReport As String, RdoAux As rdoResultset, sChave As String
Dim sAut1 As String, sAut2 As String, sAut3 As String, sAut4 As String, sAut5 As String, sAut6 As String

If Trim(lblRequerente.Caption) = "" Then
    MsgBox "Selecione o requerente do Processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(lblCodCert.Caption) = 2 Then
    If Val(txtCod.Text) > 100000 Then
        MsgBox "Certidão de Endereço Atualizado apenas para imóveis.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

If Val(lblCodCert.Caption) = 5 Then
    If Val(txtArea.Text) = 0 Then
        MsgBox "Digite a área demolida.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If txtProcDem.Text = "" Then
        MsgBox "Digite o nº de processo de demolição.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If
modLg "Emissão de " & lblTipo.Caption & " - Código: " & txtCod.Text & " - Processo nº: " & txtNumProc.Text

sChave = Chr(75) & Chr(79) & Chr(66) & Chr(85) & Chr(68) & Chr(69) & Chr(82) & Chr(65)
Sql = "DELETE FROM CERTIDAO WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect
Select Case lblTipo.Caption
    Case "CERTIDÃO DE DÉBITO NEGATIVA"
        sNomeParam = "CETDNE"
        sNomeReport = "CERTIDAODEBITO"
        sTexto = "A PREFEITURA MUNICIPAL DE JABOTICABAL certifica que o imóvel/inscrição acima referido consta do Cadastro Fiscal desta municipalidade "
        sTexto = sTexto & "não apresentando débitos de Tributo(s) Municipal(is) até a presente data, sendo a certidão válida por 30 (trinta) dias de acordo "
        sTexto = sTexto & "com o Artigo 1º do Decreto nº 5407 de 18 de fevereiro de 2010, ficando à Fazenda Municipal de Jaboticabal reservado o direito de verificação e cobranças futuras. NADA MAIS, todo referido é verdade."
    Case "CERTIDÃO DE DÉBITO POSITIVA"
        sNomeParam = "CETDPO"
        sNomeReport = "CERTIDAODEBITO"
        sTexto = "A PREFEITURA MUNICIPAL DE JABOTICABAL certifica que o imóvel/inscrição acima referido consta do Cadastro Fiscal desta municipalidade "
        sTexto = sTexto & "apresentando débitos de Tributo(s) Municipal(is) até a presente data, sendo a certidão válida por 30 (trinta) dias de acordo "
        sTexto = sTexto & "com o Artigo 1º do Decreto nº 5407 de 18 de fevereiro de 2010, ficando à Fazenda Municipal de Jaboticabal reservado o direito de verificação e cobranças futuras. NADA MAIS, todo referido é verdade."
    Case "CERTIDÃO DE DÉBITO POSITIVA COM EFEITO DE NEGATIVA"
        sNomeParam = "CETDPN"
        sNomeReport = "CERTIDAODEBITO"
        sTexto = "A PREFEITURA MUNICIPAL DE JABOTICABAL certifica que o imóvel/inscrição acima referido consta do Cadastro Fiscal desta municipalidade "
        sTexto = sTexto & "tendo quitado até a presente data parte do reparcelamento de Tributo(s) Municipal(is) até a presente data, sendo a certidão válida por 30 (trinta) dias de acordo "
        sTexto = sTexto & "com o Artigo 1º do Decreto nº 5407 de 18 de fevereiro de 2010, ficando à Fazenda Municipal de Jaboticabal reservado o direito de verificação e cobranças futuras. NADA MAIS, todo referido é verdade."
    Case "CERTIDÃO DE ENDEREÇO ATUALIZADO"
        sNomeParam = "CETEND"
        sNomeReport = "CENDERECO"
'        sTexto = "A PREFEITURA MUNICIPAL DE JABOTICABAL certifica que consta do Cadastro Técnico Fiscal desta municipalidade o imóvel de propriedade "
'        sTexto = sTexto & " do(a) Sr(a). " & xImovel.NomePropPrincipal & ", sob inscrição " & xImovel.Inscricao & ", Quadra " & xImovel.Li_Quadras & " e Lote " & xImovel.Li_Lotes & " do Bairro " & xImovel.DescBairro & ", situado à "
'        If xImovel.Li_Compl = "" Then
'            sTexto = sTexto & Trim$(xImovel.AbrevTipoLog) & " " & Trim$(xImovel.AbrevTitLog) & " " & xImovel.NomeLogradouro & " sob número " & xImovel.Li_Num & " " & xImovel.Li_Compl & ". NADA MAIS, todo referido é verdade."
'        Else
'            sTexto = sTexto & Trim$(xImovel.AbrevTipoLog) & " " & Trim$(xImovel.AbrevTitLog) & " " & xImovel.NomeLogradouro & " sob número " & xImovel.Li_Num & ", complemento: " & xImovel.Li_Compl & ". NADA MAIS, todo referido é verdade."
'        End If
    Case "CERTIDÃO DE VALOR VENAL"
        sNomeParam = "CETVVN"
        sNomeReport = "CVVIMOVEL"


'        Sql = "select vvt,vvc,vvi,areaterreno from laseriptu where codreduzido=" & Val(txtCod.Text) & " and ano=" & Year(Now)
'        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux
 '           If .RowCount > 0 Then
 '               nValorVenalTerritorial = !VVT
 '               nValorVenalPredial = !VVC
 '               nValorVenalImovel = !VVI
 '               nAreaTerreno = !AreaTerreno
 '           Else
'                CalculoIndividual Val(txtCod.Text)
 '           End If
 '          .Close
 '       End With
 Dim qd As New rdoQuery
    Set qd.ActiveConnection = cn
    qd.QueryTimeout = 0
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spCALCULO(?) }"
    qd(0) = Val(txtCod.Text)
    'qd(1) = 0
'   qd(2) = 0
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    With RdoAux
        lblVVT.Caption = FormatNumber(!vvt, 2) & " (" & FormatNumber(!vvt / !AreaTerreno, 2) & " R$/m²)"
        lblVVC.Caption = FormatNumber(!VVP, 2)
        lblVVI.Caption = FormatNumber(!VVI, 2)
        .Close
    End With
'        lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2) & " (" & FormatNumber(nValorVenalTerritorial / nAreaTerreno, 2) & " R$/m²)"
'        lblVVC.Caption = FormatNumber(nValorVenalPredial, 2)
'        lblVVI.Caption = FormatNumber(nValorVenalImovel, 2)

'        CalculoIndividual Val(txtCod.Text)
'        sTexto = "A Prefeitura Municipal de Jaboticabal, CERTIFICA, a requerimento de pessoa interessada, conforme processo nº " & txtNumProc.Text & " de 22 de janeiro de 2010,"
'        sTexto = sTexto & " que consta no Cadastro Imobiliário Fiscal do Município, lançamento do imóvem da Rua Tal, nº 999, complemento se tiver, quadra A1, lote B4, cadastro nº 12345,"
'        sTexto = sTexto & " inscrição 1.03.0416.00019.002.00.000, em nome de NOME SOBRENOME DA PESSOA, sendo o Valor Venal Territorial de R$" & FormatNumber(nValorVenalTerritorial, 2)
'        sTexto = sTexto & " e o Valor Venal Predial de R$" & FormatNumber(nValorVenalPredial, 2) & ", totalizando o Valor do Imóvel em R$" & FormatNumber(nValorVenalPredial + nValorVenalTerritorial, 2) & "."
'        sTexto = sTexto & " A presente certidão é válida pelo período de 30(trinta) dias, conforme disposto no artigo 1º do Decreto nº5407, de 18 de fevereiro de 2010, nada mais. Todo o "
'        sTexto = sTexto & "referido é verdade e dou fé."
'        txtTextoCertidao.Text = sTexto
'        sTexto = "A PREFEITURA MUNICIPAL DE JABOTICABAL certifica que o imóvel acima referido consta do Cadastro Técnico Fiscal "
'        sTexto = sTexto & " desta municipalidade sob inscrição " & xImovel.Inscricao & " ,Quadra: " & xImovel.Li_Quadras & " - Lote: " & xImovel.Li_Lotes & "  sendo o Valor Venal Territorial de R$ " & FormatNumber(nValorVenalTerritorial, 2)
 '       sTexto = sTexto & " (" & FormatNumber(nValorVenalTerritorial / xImovel.Dt_AreaTerreno, 2) & " R$/m²) e  o Valor Venal Predial de R$ "
'        sTexto = sTexto & FormatNumber(nValorVenalPredial, 2) & ", totalizando no Valor Venal do Imóvel em R$ " & FormatNumber(nValorVenalPredial + nValorVenalTerritorial, 2) & "."
    Case "CERTIDÃO DE DEMOLIÇÃO"
        sNomeParam = "CETDEM"
        sNomeReport = "CERTIDAODEMOLICAO"
        sTexto = "Certifico, a requerimento de pessoa interessada, que consta do Cadastro Fiscal desta municipalidade a demolição da área construída equivalente a " & txtArea.Text & " m² do "
        sTexto = sTexto & "imóvel acima descrito, conforme processo de demolição nº " & txtProcDem.Text & ". NADA MAIS, todo referido é verdade."
    Case "CERTIDÃO DE ISENÇÃO"
        sNomeParam = "CETISE"
        If bIsento65 Then
            sNomeReport = "CISENCAOAREA"
        Else
            Sql = "SELECT * FROM vwISENCAOPROCESSO WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODISENCAO=1 AND NUMPROCESSO IS NOT NULL"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                lblProcIsencao.Caption = SubNull(RdoAux!numprocesso)
                lblDataProcIsencao.Caption = SubNull(RdoAux!DATAPROCESSO)
                sNomeReport = "CISENCAO"
            Else
                If Not bImune Then
                    Sql = "SELECT * FROM VWISENCAOprocesso WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOISENCAO=" & Year(Now)
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    t = RdoAux.RowCount
                    RdoAux.Close
                    If t = 0 Then
                        MsgBox "Este imóvel não esta isento de IPTU em " & Year(Now), vbExclamation, "Atenção"
                        Exit Sub
                    End If
                End If
                sNomeReport = "CISENCAO"
            End If
        End If
'        sTexto = "A PREFEITURA MUNICIPAL DE JABOTICABAL certifica que o imóvel/inscrição acima referido consta do Cadastro Fiscal desta municipalidade "
'        sTexto = sTexto & "estando isento de " & sTipo & " no exercício de " & Year(Now) & "."
'        If bIsento65 = True Then
'            sTexto = sTexto & "Isento por ter área construída menor que 65m² ( " & nAreaIsento & "m² ), atendendo os preceitos do artigo 50,incisos X e XI "
'            sTexto = sTexto & " e § único, da Lei Complementar nº 07 de dezembro de 1992, cominado com o artigo 14 da Lei Complementar nº 64 de 29 de dezembro de 2003."
'        Else
'            sTexto = sTexto & sTextoIsento
 '       End If
'        sTexto = sTexto & " NADA MAIS, todo referido é verdade."
End Select

'nº da certdão


Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='" & sNomeParam & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Dados não Carregados.", vbCritical, "Atenção"
        .Close
        Exit Sub
    End If
    lblCertidao.Caption = Format(!valparam + 1, "000000")
   .Close
End With

If sNomeReport = "CVVIMOVEL" Then
    frmReport.ShowReport2 sNomeReport, frmMdi.HWND, Me.HWND
    GoTo FIM
ElseIf sNomeReport = "CENDERECO" Then
    frmReport.ShowReport2 sNomeReport, frmMdi.HWND, Me.HWND
    GoTo FIM
ElseIf sNomeReport = "CISENCAO" Or sNomeReport = "CISENCAOAREA" Then
    frmReport.ShowReport2 sNomeReport, frmMdi.HWND, Me.HWND
    GoTo FIM
End If

sAut1 = Encrypt128(NomeDeLogin, sChave)
sAut2 = Encrypt128(NomeDoComputador, sChave)
sAut3 = Encrypt128(Format(Now, "dd/mm/yyyy hh:mm"), sChave)
sAut4 = Encrypt128(lblCertidao.Caption & Format(Now, "dd/mm"), sChave)
sAut5 = Encrypt128(sNomeParam, sChave)
sAut6 = Encrypt128(Format(txtCod.Text, "000000"), sChave)

Sql = "INSERT CERTIDAO (COMPUTER,NUMCERTIDAO,REQUERENTE,CIDADAO,CODREDUZIDO,ENDERECO,TEXTO,TITULO,NUMPROCESSO,AUT1,AUT2,AUT3,AUT4,AUT5,AUT6) VALUES('"
Sql = Sql & NomeDoUsuario & "','" & lblCertidao.Caption & "','" & Mask(lblRequerente.Caption) & "','" & Mask(Left$(lblProp.Caption, 50)) & "','" & Format(txtCod.Text, "000000") & "','"
Sql = Sql & lblEnd.Caption & "','" & sTexto & "','" & lblTipo.Caption & "','" & txtNumProc.Text & "','" & sAut1 & "','" & sAut2 & "','" & sAut3 & "','" & sAut4 & "','" & sAut5 & "','" & sAut6 & "')"
cn.Execute Sql, rdExecDirect

Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM + 1 WHERE NOMEPARAM='" & sNomeParam & "'"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport sNomeReport, frmMdi.HWND, Me.HWND

Sql = "DELETE FROM CERTIDAO WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect


FIM:
Sql = "UPDATE PARAMETROS SET VALPARAM=VALPARAM + 1 WHERE NOMEPARAM='" & sNomeParam & "'"
cn.Execute Sql, rdExecDirect

Limpa

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub



Private Sub Form_Activate()
If Val(lblCodCert.Caption) = 5 Then
    lblProc.Enabled = True
    txtProcDem.Enabled = True
    lblArea.Enabled = True
    txtArea.Enabled = True
Else
    lblProc.Enabled = False
    txtProcDem.Enabled = False
    lblArea.Enabled = False
    txtArea.Enabled = False
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Set xImovel = New clsImovel
LoadMatrix
cmdPrint.Enabled = False


End Sub

Private Sub Limpa()
lblEnd.Caption = ""
lblNum.Caption = ""
lblComplemento.Caption = ""
'lblDataProc = ""
lblProp.Caption = ""
lblProcIsencao.Caption = ""
'lblDataProcIsencao.Caption = ""
lblQuadra.Caption = ""
lblLote.Caption = ""
lblInscricao.Caption = ""
lblVVC.Caption = ""
lblVVT.Caption = ""
lblVVI.Caption = ""
lblCertidao.Caption = ""
If lblCodCert.Caption = "1" Then
   lblTipo.Caption = "CERTIDÃO DE DÉBITO"
End If
lblBairro.Caption = ""
lblCPF.Caption = ""
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, DecimalPositive
End Sub

Private Sub txtCod_Change()
Limpa
cmdPrint.Enabled = False
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumProc_Change()
Limpa
cmdPrint.Enabled = False
End Sub
Private Sub CalculoIndividual(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double, RdoAux4 As rdoResultset, RdoAux5 As rdoResultset, RdoAux6 As rdoResultset
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer, bCalcProc As Boolean
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

bCalcProc = False
bIsento = False
'lblPerc.Caption = "0"

Sql = "SELECT * FROM CALCPROC WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        bCalcProc = True
    End If
   .Close
End With

If bCalcProc Then GoTo FASE1

'Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO,PERCISENCAO "
'Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND ANOISENCAO=" & Val(txtAnoCalculo.Text)
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If .RowCount > 0 Then
'        If !percisencao > 0 And !percisencao < 100 Then
'            lblPerc.Caption = !percisencao
'        Else
'            MsgBox "Este imóvel esta classificado como: " & !DESCTIPO, vbExclamation, "Atenção"
'            bIsento = True
'        End If
'    End If
'   .Close
'End With

Sql = "SELECT ISENCAO.CODREDUZIDO,ISENCAO.ANOISENCAO,ISENCAO.CODISENCAO,TIPOISENCAO.DESCTIPO "
Sql = Sql & "FROM ISENCAO INNER JOIN TIPOISENCAO ON ISENCAO.CODISENCAO = TIPOISENCAO.CODTIPO WHERE CODREDUZIDO=" & nCodReduz & " AND CODISENCAO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Este imóvel esta classificado como: " & !DESCTIPO, vbExclamation, "Atenção"
'        Exit Sub
        bIsento = True
    End If
   .Close
End With

FASE1:

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where (CADIMOB.CODREDUZIDO = " & nCodReduz & ") GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = !Li_CodBairro
 '   lblIC.Caption = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00")
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
'    lblFracao.Caption = FormatNumber(!Dt_FracaoIdeal, 2)
'    lblAreaTerreno.Caption = FormatNumber(nAreaTerreno, 2)
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
'    lblAreaPrincipal.Caption = FormatNumber(nAreaPrincipal, 2)
'    lblPredial.Caption = IIf(bTemPredial, "Sim", "Não")
    'TESTADAS
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
 '   lblTestada.Caption = FormatNumber(nTestada1, 2)
 '   lblTestadaMedia.Caption = FormatNumber(nTestadaPrincipal, 2)
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    
    'BUSCA ÁREA PRINCIPAL
    'Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        If .RowCount > 0 Then
            Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                If Not IsNull(!soma) Then
                    If !soma <= 65 And RdoAux2!USOCONSTR = 1 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) And RdoAux2!QTDEPAV < 2 And nAreaTerreno < 600 Then
                        If nAnoCalculo > 2006 Then
                            If bCalcProc = False Then
                                Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
                                Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                                If RdoAux4.RowCount = 0 Then
                                    bIsento = True
                                    MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                Else
                                    If ImovelAreaUnica(RdoAux4!CODPROPRIETARIO) Then
                                        MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                        bIsento = True
                                    End If
                                End If
                                RdoAux4.Close
                            End If
                        Else
                            If bCalcProc = False Then
                                bIsento = True
                                MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                                Limpa
                            End If
                        End If
                    End If
                End If
               .Close
            End With
        Else
            bIsento = False
            
'            Sql = "SELECT CODREDUZIDO,CODPROPRIETARIO FROM vwPROPRIETARIODUPLICADO2 WHERE CODREDUZIDO=" & nCodReduz
'            Set RdoAux4 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
 '           If RdoAux4.RowCount > 0 Then
 '               If ImovelAreaUnica(RdoAux4!CODPROPRIETARIO) Then
 '                   MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
 '                   bIsento = True
 '               End If
 '           End If
 '           RdoAux4.Close
        End If

        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
        
        'novo VVP ***********************************
        If nAnoCalculo > 2007 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            If bTemPredial Then
                Do Until .EOF
                    nUso = !USOCONSTR
                    nTipo = !TIPOCONSTR
                    nCat = !CATCONSTR
                    nFatorCategoria = 0
                    For x = 1 To UBound(aFatorC)
                        If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                           nFatorCategoria = aFatorC(x).Fator
                           Exit For
                        End If
                    Next
                    nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
                   .MoveNext
                Loop
            End If
        Else
            If bTemPredial Then
                 nUso = !USOCONSTR
                 nTipo = !TIPOCONSTR
                 nCat = !CATCONSTR
            End If
        End If
       .Close
    End With
    
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
 '   lblAgrup.Caption = FormatNumber(nValorAgrupamento, 2)
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    For x = 1 To UBound(aGleba)
        If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
             Exit For
        ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
             Exit For
        End If
    Next
    nCodGleba = aGleba(x).Codigo
    'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
    nFatorGleba = aFatorG(nCodGleba)
 '   lblFatorG.Caption = FormatNumber(nFatorGleba, 2)
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(x).Min, 2) And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
 '       lblFatorF.Caption = FormatNumber(nFatorProfundidade, 2)
     Else
        nFatorProfundidade = 1
 '       lblFatorF.Caption = "1,00"
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
 '   lblFatorS.Caption = FormatNumber(nFatorSituacao, 2)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
 '   lblFatorP.Caption = FormatNumber(nFatorPedologia, 2)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
 '   lblFatorT.Caption = FormatNumber(nFatorTopografia, 2)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
 '   lblMulF.Caption = FormatNumber(nValorFatores, 2)
    'CÁLCULO VALOR VENAL TERRITORIAL
    nFatorDistrito = aFatorD(!Distrito)
    nValorFatores = nValorFatores * nFatorDistrito
    nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * Format(nValorFatores, "#0.00")
 '   lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
'        nFatorDistrito = aFatorD(!Distrito)
'        nValorFatores = nValorFatores * nFatorDistrito
        nValorVenalTerritorial = nAreaTerreno * nValorAgrupamento * Format(nValorFatores, "#0.00")
'        lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2)
'        lblMulF.Caption = FormatNumber(nValorFatores, 2)
'        lblFatorD.Caption = FormatNumber(nFatorDistrito, 2)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        If nAnoCalculo < 2008 Then
            nValorVenalPredial = 0
            nFatorCategoria = 0
            For x = 1 To UBound(aFatorC)
                If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                   nFatorCategoria = aFatorC(x).Fator
                   Exit For
                End If
            Next
            nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
        End If
'        lblFatorC.Caption = FormatNumber(nFatorCategoria, 2)
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
'        lblVVP.Caption = FormatNumber(nValorVenalPredial, 2)
    Else
        nFatorDistrito = 0
        nFatorCategoria = 0
'        lblFatorD.Caption = FormatNumber(nFatorDistrito, 2)
'        lblFatorC.Caption = FormatNumber(nFatorCategoria, 2)
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nCodTributo = 1
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
        nValorIptu = nValorVenalImovel * (nAliquotaPredial / 100)  'reajuste 2004-2005 (TIRADO)
'        lblIPTU.Caption = FormatNumber(nValorVenalImovel * (nAliquotaPredial / 100), 2)
'        lblIPTUCorrigido.Caption = FormatNumber(nValorIPTU, 2)
    Else
        nCodTributo = 2
        nValorVenalImovel = nValorVenalTerritorial
        nValorITU = nValorVenalImovel * (nAliquotaTerritorial / 100)  'reajuste 2004-2005 (TIRADO)
'        lblIPTU.Caption = FormatNumber(nValorVenalImovel * (nAliquotaTerritorial / 100), 2)
'        lblIPTUCorrigido.Caption = FormatNumber(nValorITU, 2)
    End If
'    lblVVI.Caption = FormatNumber(nValorVenalImovel, 2)
    'COMPARAÇÃO ENTRE OS CÁLCULOS
    If bTemPredial Then
       nValorFinal = nValorIptu
    Else
       nValorFinal = nValorITU
    End If
    
    'PERCENTUAL ISENÇÃO
'    If Val(lblPerc.Caption) > 0 Then
'        nValorFinal = nValorFinal - (nValorFinal * Val(lblPerc.Caption) / 100)
'    End If
    
    
'    If bIsento Then
'        lblValorFinal.Caption = FormatNumber(0, 2)
'        lblUnica.Caption = FormatNumber(0, 2)
'        lblParcela.Caption = FormatNumber(0, 2)
'    Else
'        lblValorFinal.Caption = FormatNumber(nValorFinal, 2)
'        lblUnica.Caption = FormatNumber(nValorFinal - (nValorFinal * CDbl(lblPercUnica.Caption) / 100), 2)
'        lblParcela.Caption = FormatNumber(nValorFinal / CDbl(txtNumParc.Text), 2)
 '   End If

    lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2) & " (" & FormatNumber(nValorVenalTerritorial / nAreaTerreno, 2) & " R$/m²)"
    lblVVC.Caption = FormatNumber(nValorVenalPredial, 2)
    lblVVI.Caption = FormatNumber(nValorVenalImovel, 2)


End With

End Sub

Private Sub CalculoIndividualOld2(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = Val(SubNull(!Li_CodBairro))
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                If nNumTestadas = 0 Then
                    MsgBox "O imovel esta sem testada cadastrada", vbCritical, "Atenção"
                    Exit Sub
                End If
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If Not IsNull(!soma) Then
                If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                    bIsento = True
                    MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                End If
            End If
           .Close
        End With
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
             Select Case !USOCONSTR
                  Case 0
                     nTaxaLimpeza = 3.78
                  Case 1, 2, 3, 4, 5
                     nTaxaLimpeza = 10.57
                  Case Else
                     nTaxaLimpeza = 3.01
             End Select
        Else
             nTaxaLimpeza = 3.01
        End If
        nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
       '--CÁLCULO DA TAXA DE CONSERVAÇÃO
        If RdoAux!PAVIMENTO = 1 Then
           nTaxaConservacao = 1.35 * nTestadaPrincipal
        Else
           nTaxaConservacao = 0
        End If
        If nCodBairro = 81 Then
           nTaxaLimpeza = 1
           nTaxaConservacao = 1
        End If
        
        'novo VVP ***********************************
        nValorVenalPredial = 0
        nFatorCategoria = 0
        If bTemPredial Then
            Do Until .EOF
                nUso = !USOCONSTR
                nTipo = !TIPOCONSTR
                nCat = !CATCONSTR
                nFatorCategoria = 0
                For x = 1 To UBound(aFatorC)
                    If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
                       nFatorCategoria = aFatorC(x).Fator
                       Exit For
                    End If
                Next
                nValorVenalPredial = nValorVenalPredial + FormatNumber(!AREACONSTR, 2) * FormatNumber(nFatorCategoria, 2)
               .MoveNext
            Loop
        End If

       .Close
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    'If !Dt_CodUsoTerreno = 6 Then
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For x = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
    'Else
    '    nFatorGleba = 1
    'End If
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(x).Min, 2) And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
     Else
        nFatorProfundidade = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        'FATOR DISTRITO 98
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
'        nValorVenalPredial = 0
'        For x = 1 To UBound(aFatorC)
'            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
'               nFatorCategoria = aFatorC(x).Fator
'               Exit For
'            End If
'        Next
'        nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
        
        
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
    Else
        nValorVenalPredial = 0
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
    Else
        nValorVenalImovel = nValorVenalTerritorial
    End If
End With

lblVVT.Caption = FormatNumber(nValorVenalTerritorial, 2) & " (" & FormatNumber(nValorVenalTerritorial / nAreaTerreno, 2) & " R$/m²)"
lblVVC.Caption = FormatNumber(nValorVenalPredial, 2)
lblVVI.Caption = FormatNumber(nValorVenalImovel, 2)

End Sub

Private Sub CalculoIndividualOld(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

'CÁLCULO
Sql = "SELECT CADIMOB.CODREDUZIDO,LI_CODBAIRRO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE,"
Sql = Sql & "CADIMOB.DT_AREATERRENO,DT_CODUSOTERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.DT_CODTOPOG, FACEQUADRA.CODAGRUPA,FACEQUADRA.PAVIMENTO, "
Sql = Sql & "SUM(Areas.AREACONSTR) As SOMAAREA FROM CADIMOB LEFT OUTER JOIN AREAS ON CADIMOB.CODREDUZIDO = AREAS.CODREDUZIDO LEFT OUTER Join "
Sql = Sql & "FACEQUADRA ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "CADIMOB.Seq = FACEQUADRA.CODFACE Where CADIMOB.CODREDUZIDO = " & nCodReduz & " GROUP BY CADIMOB.CODREDUZIDO, CADIMOB.DISTRITO,CADIMOB.SETOR, CADIMOB.QUADRA, CADIMOB.LOTE,CADIMOB.SEQ, CADIMOB.UNIDADE,"
Sql = Sql & "CADIMOB.SUBUNIDADE,CADIMOB.DT_AREATERRENO, CADIMOB.DT_FRACAOIDEAL,CADIMOB.DT_CODPEDOL, CADIMOB.DT_CODSITUACAO,CADIMOB.Dt_CodTopog , FACEQUADRA.CODAGRUPA,DT_CODUSOTERRENO,LI_CODBAIRRO,PAVIMENTO "

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    'DADOS DO IMOVEL0
    nCodBairro = Val(SubNull(!Li_CodBairro))
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    If Not IsNull(!SOMAAREA) Then
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        bTemPredial = False
        nAreaPrincipal = 0
    End If
    Sql = "SELECT NUMFACE,AREATESTADA FROM TESTADA WHERE CODREDUZIDO = " & nCodReduz
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        nNumTestadas = .RowCount
        If nNumTestadas = 0 Then
            nTestadaPrincipal = 1
            nTestada1 = 1
        Else
            If nNumTestadas = 1 Then
                nTestadaPrincipal = !AREATESTADA
                nTestada1 = !AREATESTADA
            Else
                nSomaTestada = 0
                Do Until .EOF
                   If !NUMFACE = RdoAux!Seq Then
                      nTestada1 = !AREATESTADA
                   End If
                   nSomaTestada = nSomaTestada + !AREATESTADA
                  .MoveNext
                Loop
                If nNumTestadas = 0 Then
                    MsgBox "O imovel esta sem testada cadastrada", vbCritical, "Atenção"
                    Exit Sub
                End If
                nTestadaPrincipal = nSomaTestada / nNumTestadas
            End If
        End If
       .Close
    End With
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P' "
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
        Sql = "SELECT SUM(AREACONSTR) AS SOMA FROM AREAS WHERE CODREDUZIDO = " & nCodReduz
        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux3
            If Not IsNull(!soma) Then
                If !soma <= 65 And RdoAux2!USOCONSTR = 0 And (RdoAux2!CATCONSTR = 4 Or RdoAux2!CATCONSTR = 7) Then
                    bIsento = True
                    MsgBox "Imóvel isento de IPTU por ter área construida menor que 65 m²", vbInformation, "Atenção"
                End If
            End If
           .Close
        End With
        If bFracaoIdeal Then
            nTestadaPrincipal = nAreaTerreno * nTestadaPrincipal / nAreaPrincipal
        End If
       'APROVEITANDO O SELECT CALCULA TAXA DE LIMPEZA
        If bTemPredial Then
             nUso = !USOCONSTR
             nTipo = !TIPOCONSTR
             nCat = !CATCONSTR
             Select Case !USOCONSTR
                  Case 0
                     nTaxaLimpeza = 3.78
                  Case 1, 2, 3, 4, 5
                     nTaxaLimpeza = 10.57
                  Case Else
                     nTaxaLimpeza = 3.01
             End Select
        Else
             nTaxaLimpeza = 3.01
        End If
        nTaxaLimpeza = nTaxaLimpeza * nTestadaPrincipal
       '--CÁLCULO DA TAXA DE CONSERVAÇÃO
        If RdoAux!PAVIMENTO = 1 Then
           nTaxaConservacao = 1.35 * nTestadaPrincipal
        Else
           nTaxaConservacao = 0
        End If
        If nCodBairro = 81 Then
           nTaxaLimpeza = 1
           nTaxaConservacao = 1
        End If
       .Close
    End With
    'VALOR DOS AGRUPAMENTOS
    If !Dt_CodUsoTerreno = 6 Then
       nValorAgrupamento = aFatorR(7)
    Else
       nValorAgrupamento = aFatorR(nCodAgrupamento)
    End If
    
    '**************************
    'CÁLCULO DOS FATORES
    '**************************
    '**************************
    '### FATOR GLEBA ###
    '**************************
    'If !Dt_CodUsoTerreno = 6 Then
        'LOCALIZAMOS PRIMEIRO O CODIGO DA GLEBA A QUE PERTENCE O IMOVEL DE ACORDO COM A SUA AREA DO TERRENO
        For x = 1 To UBound(aGleba)
            If nAreaTerreno >= aGleba(x).Min And nAreaTerreno <= aGleba(x).Max Then
                 Exit For
            ElseIf nAreaTerreno >= aGleba(x).Min And aGleba(x).Max = 0 Then
                 Exit For
            End If
        Next
        nCodGleba = aGleba(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA
        nFatorGleba = aFatorG(nCodGleba)
        'PROCURAMOS AGORA O VALOR DO FATOR GLEBA98
    'Else
    '    nFatorGleba = 1
    'End If
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= Round(aProf(x).Min, 2) And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        nCodProfundidade = aProf(x).Codigo
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE98
     Else
        nFatorProfundidade = 1
     End If
    '**************************
    '### FATOR SITUAÇÃO ###
    '**************************
    nFatorSituacao = aFatorS(nCodSituacao)
    '**************************
    '### FATOR PEDOLOGIA ###
    '**************************
    nFatorPedologia = aFatorP(nCodPedologia)
    '**************************
    '### FATOR TOPOGRAFIA ###
    '**************************
    nFatorTopografia = aFatorT(nCodTopografia)
    '**************************
    'FIM DO CÁLCULO DOS FATORES
    '**************************
    'MULTIPLICA OS FATORES
    nValorFatores = nFatorTopografia * nFatorSituacao * nFatorPedologia * nFatorProfundidade * nFatorGleba
    'CÁLCULO VALOR VENAL TERRITORIAL
    nValorVenalTerritorial = nAreaTerreno * FormatNumber(nValorAgrupamento, 2) * FormatNumber(nValorFatores, 2)
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        'FATOR DISTRITO 98
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        For x = 1 To UBound(aFatorC)
            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
               nFatorCategoria = aFatorC(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (nAreaPrincipal * nFatorCategoria)
        
        
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
    Else
        nValorVenalPredial = 0
    End If
    'VALOR ITU/IPTU
    If bTemPredial Then
        nValorVenalImovel = nValorVenalTerritorial + nValorVenalPredial
    Else
        nValorVenalImovel = nValorVenalTerritorial
    End If
End With

End Sub

Private Sub LoadMatrix()

ReDim aFatorD(3)
ReDim aFatorP(6)
ReDim aFatorT(6)
ReDim aFatorS(6)
ReDim aFatorG(23)
ReDim aFatorR(7)

nAnoCalculo = Year(Now)

Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE CODAGRUPAMENTO<8 AND ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        aFatorP(!CODPEDOLOGIA) = !FATORPEDOLOGIA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorT(!CODTOPOG) = !FATORTOPOG
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorS(!Codsituacao) = !FATORSITUACAO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorG(!CODGLEBA) = !FATORGLEBA
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorD(!CODDISTRITO) = !FATORDISTRITO
       .MoveNext
     Loop
    .MoreResults
     Do Until .EOF
        aFatorR(!codagrupamento) = !valorterreno
       .MoveNext
     Loop
    .Close
End With

ReDim aProf(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,MINPROFUN,MAXPROFUN FROM PROFUNDIDADE ORDER BY CODDISTRITO,CODPROFUN "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aProf(UBound(aProf) + 1)
        aProf(UBound(aProf)).Distrito = !CODDISTRITO
        aProf(UBound(aProf)).Codigo = !CODPROFUN
        aProf(UBound(aProf)).Min = !MINPROFUN
        aProf(UBound(aProf)).Max = !MAXPROFUN
       .MoveNext
     Loop
    .Close
End With


ReDim aFatorF(0)
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorF(UBound(aFatorF) + 1)
        aFatorF(UBound(aFatorF)).Distrito = !CODDISTRITO
        aFatorF(UBound(aFatorF)).Codigo = !CODPROFUN
        aFatorF(UBound(aFatorF)).Fator = !FATORPROFUN
       .MoveNext
     Loop
    .Close
End With

ReDim aGleba(0)
Sql = "SELECT CODGLEBA,MINGLEBA,MAXGLEBA FROM GLEBA ORDER BY CODGLEBA "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aGleba(UBound(aGleba) + 1)
        aGleba(UBound(aGleba)).Codigo = !CODGLEBA
        aGleba(UBound(aGleba)).Min = !MINGLEBA
        aGleba(UBound(aGleba)).Max = !MAXGLEBA
       .MoveNext
     Loop
    .Close
End With

ReDim aFatorC(0)
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !CODUSO
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !CODCATEG
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub

Private Sub txtNumProc_LostFocus()
Dim nCodCidadao As Long, nNumproc As Long, nAnoproc As Integer

If Trim(txtNumProc.Text) = "" Then Exit Sub

'sValidaProc = ValidaProcesso(txtNumProc.text)
If NovoProtocolo = 0 Then
    Sql = "SELECT CODCIDAPRO FROM PROCESSO WHERE ANOPROCESS=" & Val(Right$(txtNumProc.Text, 4)) & " AND NUMEROPROC=" & Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "Cidadão não localizado no protocolo.", vbExclamation, "Atenção"
            Exit Sub
        Else
            nCodCidadao = !CODCIDAPRO
        End If
       .Close
    End With
Else
    On Error Resume Next
    nNumproc = Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 1))
    nAnoproc = Val(Right$(txtNumProc.Text, 4))

    If Right$(nNumproc, 1) <> RetornaDVProcesso(Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))) Then
        MsgBox "Número de Processo inválido", vbExclamation, "Atenção"
        Exit Sub
    Else
        Sql = "SELECT CODCIDADAO,DATAENTRADA FROM PROCESSOGTI WHERE ANO=" & nAnoproc & " AND NUMERO=" & Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Cidadão não localizado no protocolo.", vbExclamation, "Atenção"
                Exit Sub
            Else
                nCodCidadao = !CodCidadao
                lblDataProc.Caption = Format(!DATAENTRADA, "dd/mm/yyyy")
            End If
           .Close
        End With
    End If
End If

If nCodCidadao > 0 Then
    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            lblRequerente.Caption = !NomeCidadao
        End If
       .Close
    End With
Else
    lblRequerente.Caption = ""
End If
End Sub

