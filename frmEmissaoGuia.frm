VERSION 5.00
Begin VB.Form frmEmissaoGuia 
   BackColor       =   &H00E8F7F0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de guias"
   ClientHeight    =   3255
   ClientLeft      =   11115
   ClientTop       =   6300
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   7095
   Begin Tributacao.jcFrames jcFrames3 
      Height          =   825
      Left            =   60
      Top             =   60
      Width           =   6980
      _ExtentX        =   12303
      _ExtentY        =   1455
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5250
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   435
         Width           =   1605
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   435
         Width           =   4185
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   26
         Top             =   110
         Width           =   1065
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "(Digite o código reduzido e tecle ENTER)"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2160
         TabIndex        =   31
         Top             =   150
         Width           =   3105
      End
      Begin VB.Label lblRS 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome........:"
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   465
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código......:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   27
         Top             =   150
         Width           =   855
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   975
      Left            =   60
      Top             =   2220
      Width           =   6980
      _ExtentX        =   12303
      _ExtentY        =   1720
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço de entrega"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtUFEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6420
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox txtCidadeEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtCepEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5940
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   270
         Width           =   885
      End
      Begin VB.TextBox txtBairroEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtEnderecoEnt 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF.:"
         Height          =   225
         Index           =   13
         Left            =   6060
         TabIndex        =   20
         Top             =   630
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade..:"
         Height          =   225
         Index           =   12
         Left            =   3270
         TabIndex        =   19
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço..:"
         Height          =   225
         Index           =   10
         Left            =   60
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep..:"
         Height          =   225
         Index           =   8
         Left            =   5460
         TabIndex        =   17
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro........:"
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   16
         Top             =   630
         Width           =   870
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   1305
      Left            =   60
      Top             =   900
      Width           =   6980
      _ExtentX        =   12303
      _ExtentY        =   2302
      FrameColor      =   8421504
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Endereço de localização"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   1
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtInscricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   930
         Width           =   2175
      End
      Begin VB.TextBox txtLote 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5730
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   930
         Width           =   1080
      End
      Begin VB.TextBox txtUF 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   6420
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Width           =   405
      End
      Begin VB.TextBox txtCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   1875
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3990
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   930
         Width           =   1080
      End
      Begin VB.TextBox txtCep 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         Height          =   285
         Left            =   5940
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Width           =   885
      End
      Begin VB.TextBox txtBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8F7F0&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   960
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   4395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição...:"
         Height          =   225
         Index           =   14
         Left            =   60
         TabIndex        =   14
         Top             =   990
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote..:"
         Height          =   225
         Index           =   4
         Left            =   5190
         TabIndex        =   6
         Top             =   990
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra..:"
         Height          =   225
         Index           =   3
         Left            =   3270
         TabIndex        =   5
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF.:"
         Height          =   225
         Index           =   2
         Left            =   6060
         TabIndex        =   4
         Top             =   645
         Width           =   345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade..:"
         Height          =   225
         Index           =   1
         Left            =   3270
         TabIndex        =   3
         Top             =   645
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço..:"
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep..:"
         Height          =   225
         Index           =   9
         Left            =   5460
         TabIndex        =   1
         Top             =   300
         Width           =   510
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro........:"
         Height          =   225
         Index           =   11
         Left            =   60
         TabIndex        =   0
         Top             =   650
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmEmissaoGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTipo As Integer

Private Sub Form_Load()
Centraliza Me
Me.Top = Me.Top - 1000
Me.Left = Me.Left - 1000
nTipo = 0
End Sub

Private Sub txtCodigo_Change()
Limpa
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If Val(txtCodigo.Text) > 0 Then
        CarregaContribuinte CLng(txtCodigo.Text)
    End If
Else
    Tweak txtCodigo, KeyAscii, IntegerPositive
End If

End Sub

Private Sub Limpa()

txtNome.Text = ""
txtDoc.Text = ""
txtEndereco.Text = ""
txtCep.Text = ""
txtBairro.Text = ""
txtCidade.Text = ""
txtUF.Text = ""
txtInscricao.Text = ""
txtQuadra.Text = ""
txtLote.Text = ""
txtEnderecoEnt.Text = ""
txtCepEnt.Text = ""
txtBairroEnt.Text = ""
txtCidadeEnt.Text = ""
txtUFEnt.Text = ""

End Sub

Private Sub CarregaContribuinte(nCodReduz As Long)
Dim Sql As String, RdoAux As rdoResultset, sNome As String, sDoc As String, sLote As String, sQuadra As String, sInscricao As String, tTipoEnd As SeqEndereco
Dim xImovel As clsImovel, sEndereco As String, nNum As Integer, sComplemento As String, sBairro As String, sCidade As String, sUF As String, sCep As String
Dim sEnderecoEnt As String, nNumEnt As Integer, sComplementoEnt As String, sBairroEnt As String, sCidadeEnt As String, sUFEnt As String, sCepEnt As String, nTipoEnd As Integer
Set xImovel = New clsImovel

If nCodReduz < 100000 Then
    nTipo = 1
    tTipoEnd = Imobiliario
ElseIf nCodReduz >= 100000 And nCodReduz < 300000 Then
    nTipo = 2
    tTipoEnd = mobiliario
ElseIf nCodReduz >= 500000 And nCodReduz < 700000 Then
    nTipo = 3
    tTipoEnd = cidadao
Else
    nTipo = 0
End If
sDoc = ""

Ocupado
If tTipoEnd = Imobiliario Then
    Sql = "select * from vwfullimovel where codreduzido=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            GoTo SemCadastro
        Else
            sNome = !nomecidadao
            sDoc = SubNull(!Cnpj)
            If sDoc = "" Then
                sDoc = Format(Val(SubNull(!CPF)), "00000000000")
            End If
            sQuadra = SubNull(!Li_Quadras)
            sLote = SubNull(!Li_Lotes)
            sInscricao = !Inscricao
            nTipoEnd = !Ee_TipoEnd
        End If
       .Close
    End With
ElseIf tTipoEnd = mobiliario Then
    Sql = "select * from mobiliario where codigomob=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            GoTo SemCadastro
        Else
            sNome = !razaosocial
            sDoc = SubNull(!Cnpj)
            If sDoc = "" Then
                sDoc = Format(Val(SubNull(!CPF)), "00000000000")
            End If
            sQuadra = "": sLote = ""
            sInscricao = SubNull(!inscestadual)
        
        End If
       .Close
    End With
ElseIf tTipoEnd = cidadao Then
    Sql = "select * from cidadao where codcidadao=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            GoTo SemCadastro
        Else
            sNome = !nomecidadao
            sDoc = SubNull(!Cnpj)
            If sDoc = "" Then
                sDoc = Format(Val(SubNull(!CPF)), "00000000000")
            End If
            sQuadra = "": sLote = ""
            sInscricao = ""
            
        End If
       .Close
    End With
Else
    GoTo SemCadastro
End If

xImovel.RetornaEndereco nCodReduz, tTipoEnd, Localizacao
sEndereco = xImovel.Endereco
nNum = Val(xImovel.Numero)
sComplemento = xImovel.Complemento
sBairro = xImovel.Bairro
sCidade = xImovel.Cidade
sUF = xImovel.UF
sCep = xImovel.Cep

If tTipoEnd = Imobiliario Then
    If nTipoEnd = 0 Then
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Localizacao
    ElseIf nTipoEnd = 1 Then
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Cadastrocidadao
    Else
        xImovel.RetornaEndereco nCodReduz, Imobiliario, Entrega
    End If
Else
    xImovel.RetornaEndereco nCodReduz, tTipoEnd, Entrega
End If
sEnderecoEnt = xImovel.Endereco
nNumEnt = Val(xImovel.Numero)
sComplementoEnt = xImovel.Complemento
sBairroEnt = xImovel.Bairro
sCidadeEnt = xImovel.Cidade
sUFEnt = xImovel.UF
sCepEnt = xImovel.Cep

txtNome.Text = sNome
If Len(sDoc) = 11 Then
    sDoc = Format(sDoc, "000\.000\.000-00")
ElseIf Len(sDoc) = 14 Then
    sDoc = Format(sDoc, "00\.000\.000/0000-00")
End If
txtDoc.Text = sDoc
txtQuadra.Text = sQuadra
txtLote.Text = sLote
txtInscricao.Text = sInscricao

txtEndereco.Text = sEndereco & ", " & nNum & " " & sComplemento
txtBairro.Text = sBairro
txtCidade.Text = sCidade
txtUF.Text = sUF
txtCep.Text = sCep

txtEnderecoEnt.Text = sEnderecoEnt & ", " & nNumEnt & " " & sComplementoEnt
txtBairroEnt.Text = sBairroEnt
txtCidadeEnt.Text = sCidadeEnt
txtUFEnt.Text = sUFEnt
txtCepEnt.Text = sCepEnt

Liberado

If Trim(txtDoc.Text) = "" Or Val(txtDoc.Text) = 0 Then
    MsgBox "CPF/CNPJ obrigatório para emissão de guia.", vbCritical, "Erro"
    Exit Sub
End If

If sEndereco = "" Then
    MsgBox "Endereço obrigatório para emissão de guia.", vbCritical, "Erro"
    Exit Sub
End If


frmEmissaoGuia2.show vbModal

Exit Sub
SemCadastro:
Liberado
MsgBox "Inscrição não cadastrada.", vbCritical, "Erro"


End Sub
