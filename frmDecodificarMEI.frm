VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDecodificarMEI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Decodificação de arquivos do MEI"
   ClientHeight    =   8550
   ClientLeft      =   9150
   ClientTop       =   3675
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames frProgress 
      Height          =   1155
      Left            =   2190
      Top             =   3210
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2037
      FrameColor      =   255
      FillColor       =   4210688
      TextBoxColor    =   8454016
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
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
      ThemeColor      =   3
      ColorFrom       =   192
      ColorTo         =   8438015
      Begin Tributacao.XP_ProgressBar pBar 
         Height          =   165
         Left            =   150
         TabIndex        =   8
         Top             =   780
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   291
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
         Color           =   255
         Scrolling       =   1
      End
      Begin VB.Label lblFileNumber 
         BackStyle       =   0  'Transparent
         Caption         =   "Importando Empresa 0 de 0"
         ForeColor       =   &H0080FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   90
         Width           =   4305
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Caption         =   "ARRECADA08.ret"
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   4155
      End
   End
   Begin prjChameleon.chameleonButton cmdGerarIM 
      Height          =   345
      Left            =   7530
      TabIndex        =   7
      ToolTipText     =   "Criar uma inscrição municipal para esta empresa"
      Top             =   540
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "G&erar IM"
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
      MICON           =   "frmDecodificarMEI.frx":0000
      PICN            =   "frmDecodificarMEI.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   7485
      Left            =   90
      TabIndex        =   5
      Top             =   1020
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   13203
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmDecodificarMEI.frx":0107
   End
   Begin VB.ComboBox cmbReg 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1620
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   570
      Width           =   4635
   End
   Begin VB.TextBox txtArq 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   5985
   End
   Begin prjChameleon.chameleonButton cmdArq 
      Height          =   345
      Left            =   150
      TabIndex        =   1
      Top             =   90
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Arquivo"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmDecodificarMEI.frx":0189
      PICN            =   "frmDecodificarMEI.frx":01A5
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
      Height          =   345
      Left            =   7530
      TabIndex        =   2
      ToolTipText     =   "Executar a operação selecionada"
      Top             =   120
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Executar"
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
      MICON           =   "frmDecodificarMEI.frx":0260
      PICN            =   "frmDecodificarMEI.frx":027C
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
      Height          =   345
      Left            =   6330
      TabIndex        =   6
      ToolTipText     =   "Gravar os Dados"
      Top             =   540
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MICON           =   "frmDecodificarMEI.frx":031B
      PICN            =   "frmDecodificarMEI.frx":0337
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
      Caption         =   "Lista de Empresas.:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1425
   End
   Begin VB.Menu mnuGravar 
      Caption         =   "&Gravar"
      Visible         =   0   'False
      Begin VB.Menu mnuEmpresa 
         Caption         =   "Apenas esta empresa"
      End
      Begin VB.Menu mnuTodos 
         Caption         =   "Todas as empresas"
      End
   End
End
Attribute VB_Name = "frmDecodificarMEI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tTabelaCnae
    Codigo As String
    Nome As String
End Type

Private Type tTabela
    Codigo As Integer
    Nome As String
End Type

Private Type tNatureza
    Codigo As Integer
    Nome As String
End Type

Private Type tQualificacao
    Codigo As Integer
    Nome As String
End Type

Private Type tMunicipio
    Codigo As Integer
    Nome As String
    UF As String
End Type


Private Type tMei
    id As Integer
    Tipo As Integer
    Cnpj As String
    Nome As String
    Fantasia As String
    Recibo As String
    IdSolicitacao As String
    AtoOficio As String
    Matriz As String
    SIMPLES As String
    CodNatureza As String
    Natureza As String
    Nire As String
    Capital As Double
    CnpjMatriz As String
    NomeResp As String
    CpfResp As String
    CodQualificacao As String
    Qualificacao As String
    Endereco As String
    EnderecoTipo As String
    EnderecoCodigo As Integer
    EnderecoNumero As String
    EnderecoComplemento As String
    Bairro As String
    BairroCodigo As Integer
    CodMunicipio As String
    Municipio As String
    MunicipoCodigo As Integer
    UF As String
    Cep As String
    DDD As String
    TelefoneNF As String
    Telefone1 As String
    Telefone2 As String
    Fax As String
    Email As String
    Permanencia_Livro As String
    Opção_Livro As String
    Opção_Doc As String
    Processo_Eletronico As String
    Utilização_ECF As String
    Abrangencia As String
    OrigemEndereco As String
    Endereco_Resp As String
    Endereco_RespTipo As String
    Endereco_RespCodigo As String
    Endereco_RespNumero As String
    Endereco_RespComplemento As String
    Bairro_Resp As String
    Bairro_RespCodigo As Integer
    CodMunicipio_Resp As String
    Municipio_Resp As String
    Municipio_RespCodigo As Integer
    UF_Resp As String
    Cep_Resp As String
    Telefone_Resp As String
    Fax_Resp As String
    Email_Resp As String
    Cnae As String
    DataAbertura As String
    CodigoCidadao As Long
End Type

Private Type tEvento
    id As Integer
    Tipo As Integer
    Codigo As String
    Nome As String
    Data As String
End Type

Private Type tCnae
    id As Integer
    Tipo As Integer
    Cnae As String
    Nome As String
End Type

Dim aMei() As tMei, aEvento() As tEvento, aTabela() As tTabela, aNatureza() As tNatureza, aQualificacao() As tQualificacao, aMunicipio() As tMunicipio, aCnae() As tCnae, aTabelaCnae() As tTabelaCnae

Private Sub cmbReg_Click()
Dim sCNPJ As String, nPos As Integer, sNome As String, i As Integer

Rtb.Text = ""
If cmbReg.ListIndex = -1 Then Exit Sub
sNome = cmbReg.Text

For nPos = 1 To UBound(aMei)
    With aMei(nPos)
        If .Nome = sNome Then
             Rtb.SelColor = vbBlack
             Rtb.SelText = "CPF/CNPJ: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Cnpj & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Razão Social: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .Nome & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Nome Fantasia: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .Fantasia & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Matriz: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .Matriz & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "CNPJ Matriz: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .CnpjMatriz & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Optante do Simples: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .SIMPLES & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Recibo solicitação: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .Recibo & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Identificação da solicitação: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .IdSolicitacao & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Nº de controle do Ato Ofício: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .AtoOficio & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Natureza jurídica: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & Format(Val(.CodNatureza), "000-0") & " - " & .Natureza & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Nº Ident. do Registro da Empresa (NIRE): "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & .Nire & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Capital Social: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " R$ " & Format(.Capital, "#0.00") & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Nome do Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .NomeResp & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "CPF do Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .CpfResp & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Qualificação do Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = " " & Format(Val(.CodQualificacao), "00") & " - " & .Qualificacao & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Endereço: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .EnderecoCodigo & " - " & .EnderecoTipo & " " & .Endereco & ", " & .EnderecoNumero & " " & .EnderecoComplemento & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Bairro: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .BairroCodigo & " - " & .Bairro & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Município: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = Format(.MunicipoCodigo, "0000") & " - " & .Municipio & "/" & .UF & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Cep: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Cep & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Telefone1: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Telefone1
             Rtb.SelColor = vbBlack
             Rtb.SelText = "  Telefone2: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Telefone2
             Rtb.SelColor = vbBlack
             Rtb.SelText = "  Fax: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Fax & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Email: "
             Rtb.SelColor = vbBlue
             Rtb.SelUnderline = True
             Rtb.SelText = .Email & vbCrLf
             Rtb.SelUnderline = False
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Permanência Livros Fiscais: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Permanencia_Livro & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Opção Livros Eletrônicos: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Opção_Livro & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Opção Documentos Eletrônicos: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Opção_Doc & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Processamento Eletrônico Dados: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Processo_Eletronico & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Utilização ECF: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Utilização_ECF & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Abrangência Municipal: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Abrangencia & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Origem Endereço Responsável : "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .OrigemEndereco & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Endereço Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .EnderecoCodigo & " - " & .Endereco_RespTipo & " " & .Endereco_Resp & ", " & .Endereco_RespNumero & " " & .Endereco_RespComplemento & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Bairro Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Bairro_RespCodigo & " - " & .Bairro_Resp & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Município Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Municipio_Resp & "/" & .UF_Resp & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Cep Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Cep_Resp & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Telefone Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Telefone_Resp
             Rtb.SelColor = vbBlack
             Rtb.SelText = "  Fax Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Fax_Resp & vbCrLf
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Email Responsável: "
             Rtb.SelColor = vbBlue
             Rtb.SelUnderline = True
             Rtb.SelText = .Email_Resp & vbCrLf
             Rtb.SelUnderline = False
             Rtb.SelColor = vbBlack
             Rtb.SelText = "Cnae Fiscal: "
             Rtb.SelColor = vbBlue
             Rtb.SelText = .Cnae & vbCrLf
             Rtb.SelColor = Vinho
             Rtb.SelText = "Cnae Secundários: " & vbCrLf
             Rtb.SelText = "------------------------------" & vbCrLf
             Rtb.SelColor = vbBlue
             
             For i = 1 To UBound(aCnae)
                 With aCnae(i)
                     If .id = nPos And .Tipo = 4 Then
                         Rtb.SelColor = vbBlack
                         Rtb.SelText = "Cnae: "
                         Rtb.SelColor = vbBlue
                         Rtb.SelText = .Cnae & .Nome & vbCrLf
                     End If
                 End With
             Next

             Rtb.SelColor = Vinho
             Rtb.SelText = vbCrLf & "---- Lista de Ocorrências ---- " & vbCrLf
             Rtb.SelText = "------------------------------------------------------- " & vbCrLf
             
             For i = 1 To UBound(aEvento)
                 With aEvento(i)
                     If .id = nPos And .Tipo = 1 Then
                         Rtb.SelColor = vbBlack
                         Rtb.SelText = "Data: "
                         Rtb.SelColor = vbBlue
                         Rtb.SelText = .Data & " "
                         Rtb.SelColor = vbBlack
                         Rtb.SelText = "Evento: "
                         Rtb.SelColor = vbBlue
                         Rtb.SelText = .Codigo & " - "
                         Rtb.SelColor = vbBlue
                         Rtb.SelText = .Nome & vbCrLf
                     End If
                 End With
             Next
        
            
           
        End If
    End With
Next

End Sub

Private Sub cmdArq_Click()
Dim fName As String, cc As cCommonDlg, ff As Long, sReg As String
cmbReg.Clear
Rtb.Text = ""
ff = FreeFile
Set cc = New cCommonDlg
If cc.VBGetOpenFileName(fName, "", True, False, False, False, "Texto[*.txt]", , sPathBin, "Selecione um arquivo so Simples ou MEI", , , , False) Then
    txtArq.Text = fName
    Open fName For Input As #ff
    Do While Not EOF(1)
        Line Input #ff, sReg
        Exit Do
    Loop
    Close #ff
    If Left(sReg, 6) <> "016587" And Left(sReg, 6) <> "046587" Then
        MsgBox "Arquivo inválido", vbCritical, "Erro"
        txtArq.Text = ""
    End If
End If
End Sub

Private Sub cmdExec_Click()
Dim fName As String, sReg As String, sNome As String, sCNPJ As String, sTipo As String, sRecibo As String, sIdSolicitacao As String
Dim Sql As String, RdoAux As rdoResultset, Item As Long, z As Integer, bFind As Boolean, nCodReduz As Long, sMatriz As String
Dim nPos As Long, nTot As Long, sAtoOficio As String, sCodEvento As String, i As Integer, sDataEvento As String, t As Integer
Dim sSimples As String, nId As Integer, sFantasia As String, sNatureza As String, sCodNatureza As String, sNire As String, sCapital As String
Dim sCNPJMatriz As String, sNomeResp As String, sCpfResp As String, sCodQualif As String, sQualif As String, sTipoEnd As String, sNomeEnd As String, nCodEndereco As Integer
Dim sNumImovel As String, sComplImovel As String, sCodMunicipio As String, sMunicipio As String, sUF As String, sBairro As String, nBairroCodigo As Integer, sDDD1 As String, sTelefone1 As String
Dim sDDD2 As String, sTelefone2 As String, sDDD3 As String, sFax As String, sEmail As String, sOpcao1 As String, sOpcao2 As String, sOpcao3 As String, sOpcao4 As String
Dim sOpcao5 As String, sAbrangencia As String, sOrigemEndereco As String, sTipoEndResp As String, sNomeEndResp As String, nCodEnderecoResp As Integer
Dim sNumImovelResp As String, sComplImovelResp As String, sCodMunicipioResp As String, sMunicipioResp As String, sUFResp As String, sBairroResp As String, nBairroRespCodigo As Integer, sDDDResp As String
Dim sDDDFaxResp As String, sFaxResp As String, sEmailResp As String, sTelefoneResp As String, sCepResp As String, sCnae As String, sCnaeSec As String
Dim nMunicipioCodigo As Integer, nMunicipioRespCodigo As Integer, nCodCidadao As Long
Ocupado
Rtb.Text = ""
cmbReg.Clear
fName = txtArq.Text
nTot = 0
nPos = 1
nId = 0
ReDim aEvento(0): ReDim aCnae(0)
If fName = "" Then
    Liberado
    MsgBox "Selecione um arquivo.", vbCritical, "Erro"
    Exit Sub
End If

Open fName For Input As #15
Do While Not EOF(15)
    Line Input #15, sReg
    sNome = Trim(Mid(sReg, 1342, 150))
    If IsNumeric(Left(sNome, 1)) Then
        nTot = nTot + 1
    End If
Loop
Close #15

frProgress.Visible = True
ReDim aMei(0)
Open fName For Input As #16
nPos = 1
Do While Not EOF(16)
    CallPb nPos, nTot
    Line Input #16, sReg
    sTipo = Left(sReg, 2)
    sCNPJ = Mid(sReg, 17, 14)
    sNome = Trim(Mid(sReg, 1342, 150))
    sFantasia = Trim(Mid(sReg, 1795, 55))
    sRecibo = Trim(Mid(sReg, 7, 10))
    sIdSolicitacao = Trim(Mid(sReg, 17, 14))
    sAtoOficio = Trim(Mid(sReg, 80, 25))
    sMatriz = IIf(Mid(sReg, 1341, 1) = "1", "SIM", "NÃO")
    sSimples = IIf(Mid(sReg, 6838, 1) = "S", "SIM", "NÃO")
    sCodNatureza = Trim(Mid(sReg, 1850, 4))
    sNatureza = Retorna_Natureza(Val(sCodNatureza))
    sNire = Trim(Mid(sReg, 1855, 11))
    sCapital = Trim(Mid(sReg, 1916, 14))
    sCNPJMatriz = Trim(Mid(sReg, 1930, 14))
    sNomeResp = Trim(Mid(sReg, 1948, 60))
    sCpfResp = Trim(Mid(sReg, 2008, 11))
    sCodQualif = Trim(Mid(sReg, 2019, 2))
    sQualif = Retorna_Qualificacao(Val(sCodQualif))
    sTipoEnd = Trim(Mid(sReg, 2092, 6))
    sNomeEnd = Trim(Mid(sReg, 2098, 60))
    sNumImovel = Trim(Mid(sReg, 2158, 6))
    sComplImovel = Trim(Mid(sReg, 2164, 156))
    sBairro = Trim(Mid(sReg, 2320, 50))
    sCodMunicipio = Trim(Mid(sReg, 2420, 4))
    sUF = Trim(Mid(sReg, 2424, 2))
    sMunicipio = Retorna_Municipio(sUF, Val(sCodMunicipio))
    sCep = Trim(Mid(sReg, 2426, 8))
    sDDD1 = Trim(Mid(sReg, 2642, 4))
    sTelefone1 = Trim(Mid(sReg, 2646, 8))
    sDDD2 = Trim(Mid(sReg, 2654, 4))
    sTelefone2 = Trim(Mid(sReg, 2658, 8))
    sDDD3 = Trim(Mid(sReg, 2666, 4))
    sFax = Trim(Mid(sReg, 2670, 8))
    sEmail = Trim(Mid(sReg, 2678, 115))
    sOpcao1 = IIf(Mid(sReg, 3241, 1) = "S", "SIM", "NÃO")
    sOpcao2 = IIf(Mid(sReg, 3242, 1) = "S", "SIM", "NÃO")
    sOpcao3 = IIf(Mid(sReg, 3243, 1) = "S", "SIM", "NÃO")
    sOpcao4 = IIf(Mid(sReg, 3244, 1) = "S", "SIM", "NÃO")
    sOpcao5 = IIf(Mid(sReg, 3245, 1) = "S", "SIM", "NÃO")
    sAbrangencia = IIf(Mid(sReg, 3486, 1) = "S", "SIM", "NÃO")
    sOrigemEndereco = Trim(Mid(sReg, 3525, 1))
    If Val(sOrigemEndereco) = 1 Then
        sOrigemEndereco = "COLETADO PELO USUÁRIO"
    ElseIf Val(sOrigemEndereco) = 2 Then
        sOrigemEndereco = "RECUPERADO DA BASE CPF"
    ElseIf Val(sOrigemEndereco) = 3 Then
        sOrigemEndereco = "RECUPERADO DA BASE CNPJ"
    End If
    sTipoEndResp = Trim(Mid(sReg, 3526, 6))
    sNomeEndResp = Trim(Mid(sReg, 3532, 60))
    sNumImovelResp = Trim(Mid(sReg, 3592, 6))
    sComplImovelResp = Trim(Mid(sReg, 3598, 156))
    sBairroResp = Trim(Mid(sReg, 3754, 50))
    sCodMunicipioResp = Trim(Mid(sReg, 3854, 4))
    sUFResp = Trim(Mid(sReg, 3866, 2))
    sMunicipioResp = Retorna_Municipio(sUFResp, Val(sCodMunicipioResp))
    sCepResp = Trim(Mid(sReg, 3858, 8))
    sDDDResp = Trim(Mid(sReg, 3868, 4))
    sTelefoneResp = Trim(Mid(sReg, 3872, 8))
    sDDDFaxResp = Trim(Mid(sReg, 3880, 4))
    sFaxResp = Trim(Mid(sReg, 3884, 8))
    sEmailResp = Trim(Mid(sReg, 3892, 115))
    sCnae = Trim(Mid(sReg, 1421, 7))
    
    Sql = "select * from cidade where siglauf='" & sUF & "' and desccidade like '%" & sMunicipio & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nMunicipioCodigo = RdoAux!CodCidade
    Else
        nMunicipioCodigo = 0
    End If
    RdoAux.Close
    
    Sql = "select * from cidade where siglauf='" & sUF & "' and desccidade like '%" & sMunicipioResp & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nMunicipioRespCodigo = RdoAux!CodCidade
    Else
        nMunicipioRespCodigo = 0
    End If
    RdoAux.Close
    
    
    
    
    Sql = "select * from logradouro where nomelogradouro='" & sNomeEnd & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nCodEndereco = RdoAux!CodLogradouro
    Else
        nCodEndereco = 0
    End If
    RdoAux.Close

    Sql = "select * from logradouro where nomelogradouro='" & sNomeEndResp & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nCodEnderecoResp = RdoAux!CodLogradouro
    Else
        nCodEnderecoResp = 0
    End If
    RdoAux.Close

    Sql = "select * from bairro where siglauf='" & sUF & "' and codcidade=" & nMunicipioCodigo & " and descbairro like '%" & sBairro & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nBairroCodigo = RdoAux!CodBairro
    Else
        nBairroCodigo = 0
    End If
    RdoAux.Close

    Sql = "select * from bairro where siglauf='" & sUF & "' and codcidade=" & nMunicipioCodigo & " and descbairro like '%" & sBairroResp & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        nBairroRespCodigo = RdoAux!CodBairro
    Else
        nBairroRespCodigo = 0
    End If
    RdoAux.Close


    
    bFind = False
    For z = 1 To UBound(aMei)
        If aMei(z).Nome = sNome Then
            bFind = True
            Exit For
        End If
    Next
    If sTipo = "01" Then
        If Not bFind Then
            nId = nId + 1
            ReDim Preserve aMei(UBound(aMei) + 1)
            aMei(UBound(aMei)).id = nId
            aMei(UBound(aMei)).Tipo = Val(sTipo)
            aMei(UBound(aMei)).Nome = sNome
            aMei(UBound(aMei)).Fantasia = sFantasia
            aMei(UBound(aMei)).Matriz = sMatriz
            aMei(UBound(aMei)).SIMPLES = sSimples
            aMei(UBound(aMei)).CodNatureza = sCodNatureza
            aMei(UBound(aMei)).Natureza = sNatureza
            aMei(UBound(aMei)).Nire = sNire
            aMei(UBound(aMei)).Capital = Val(sCapital) / 100
            If Mid(sCNPJMatriz, 9, 3) = "000" Then
                aMei(UBound(aMei)).CnpjMatriz = Format(sCNPJMatriz, "0#\.###\.###/####-##")
            Else
                If sCNPJMatriz <> "" Then
                    aMei(UBound(aMei)).CnpjMatriz = Format(sCNPJMatriz, "00#\.###\.###-##")
                Else
                    aMei(UBound(aMei)).CnpjMatriz = "(Dados em branco)"
                End If
            End If
            If Mid(sCNPJ, 9, 3) = "000" Then
                aMei(UBound(aMei)).Cnpj = Format(sCNPJ, "0#\.###\.###/####-##")
            Else
                If sCNPJ <> "" Then
                    aMei(UBound(aMei)).Cnpj = Format(sCNPJ, "00#\.###\.###-##")
                Else
                    aMei(UBound(aMei)).Cnpj = aMei(UBound(aMei)).CnpjMatriz
                End If
            End If
            aMei(UBound(aMei)).CpfResp = Format(sCpfResp, "00#\.###\.###-##")
            aMei(UBound(aMei)).NomeResp = sNomeResp
            aMei(UBound(aMei)).CodQualificacao = sCodQualif
            aMei(UBound(aMei)).Qualificacao = sQualif
            aMei(UBound(aMei)).EnderecoTipo = sTipoEnd
            aMei(UBound(aMei)).Endereco = sNomeEnd
            aMei(UBound(aMei)).EnderecoCodigo = nCodEndereco
            aMei(UBound(aMei)).EnderecoNumero = sNumImovel
            aMei(UBound(aMei)).EnderecoComplemento = sComplImovel
            aMei(UBound(aMei)).Qualificacao = sQualif
            aMei(UBound(aMei)).Bairro = sBairro
            aMei(UBound(aMei)).BairroCodigo = nBairroCodigo
            aMei(UBound(aMei)).Municipio = sMunicipio
            aMei(UBound(aMei)).MunicipoCodigo = nMunicipioCodigo
            aMei(UBound(aMei)).UF = sUF
            aMei(UBound(aMei)).Cep = Format(sCep, "00000-000")
            aMei(UBound(aMei)).DDD = sDDD1
            aMei(UBound(aMei)).TelefoneNF = sTelefone1
            aMei(UBound(aMei)).Telefone1 = sDDD1 & "-" & sTelefone1
            aMei(UBound(aMei)).Telefone2 = sDDD2 & "-" & sTelefone2
            aMei(UBound(aMei)).Fax = sDDD3 & "-" & sFax
            aMei(UBound(aMei)).Email = sEmail
            aMei(UBound(aMei)).Capital = IIf(Len(sCapital) > 0, RemovePonto(sCapital), 0)
            aMei(UBound(aMei)).Permanencia_Livro = sOpcao1
            aMei(UBound(aMei)).Opção_Livro = sOpcao2
            aMei(UBound(aMei)).Opção_Doc = sOpcao3
            aMei(UBound(aMei)).Processo_Eletronico = sOpcao4
            aMei(UBound(aMei)).Utilização_ECF = sOpcao5
            aMei(UBound(aMei)).Abrangencia = sAbrangencia
            aMei(UBound(aMei)).OrigemEndereco = sOrigemEndereco
            aMei(UBound(aMei)).Endereco_RespTipo = sTipoEndResp
            aMei(UBound(aMei)).Endereco_Resp = sNomeEndResp
            aMei(UBound(aMei)).Endereco_RespCodigo = nCodEnderecoResp
            aMei(UBound(aMei)).Endereco_RespNumero = sNumImovelResp
            aMei(UBound(aMei)).Endereco_RespComplemento = sComplImovelResp
            aMei(UBound(aMei)).Bairro_Resp = sBairroResp
            aMei(UBound(aMei)).Bairro_RespCodigo = nBairroRespCodigo
            aMei(UBound(aMei)).Municipio_Resp = sMunicipioResp
            aMei(UBound(aMei)).Municipio_RespCodigo = nMunicipioRespCodigo
            aMei(UBound(aMei)).UF_Resp = sUFResp
            aMei(UBound(aMei)).Cep_Resp = Format(sCepResp, "00000-000")
            aMei(UBound(aMei)).Telefone_Resp = sDDDResp & "-" & sTelefoneResp
            aMei(UBound(aMei)).Fax_Resp = sDDDFaxResp & "-" & sFaxResp
            aMei(UBound(aMei)).Email_Resp = sEmailResp
            aMei(UBound(aMei)).CodigoCidadao = nCodCidadao
        End If
    Else
        aMei(UBound(aMei)).id = nId
        aMei(UBound(aMei)).Cnae = Format(sCnae, "0000-0/00") & " (" & Retorna_Cnae(sCnae) & ")"
    End If
    
    aMei(UBound(aMei)).Recibo = sRecibo
    aMei(UBound(aMei)).IdSolicitacao = sIdSolicitacao
    aMei(UBound(aMei)).AtoOficio = sAtoOficio
    
    t = 0
    For i = 1245 To 1268 Step 3
        sCodEvento = Trim(Mid(sReg, i, 3))
        
        sDataEvento = Trim(Mid(sReg, 1269 + (t * 8), 8))
        If sCodEvento <> "" Then
            AdicionaEvento sCodEvento, sDataEvento, UBound(aMei), Val(sTipo)
        End If
        
        If sCodEvento = "101" Then
            aMei(UBound(aMei)).DataAbertura = Right(sDataEvento, 2) & "/" & Mid(sDataEvento, 5, 2) & "/" & Left(sDataEvento, 4)
        End If
        
        t = t + 1
    Next
    
    t = 0
    For i = 1460 To 2152 Step 7
        sCnaeSec = Trim(Mid(sReg, i, 7))
        If sCnaeSec <> "" And Left(sReg, 2) = "04" Then
            AdicionaCnae Format(sCnaeSec, "0000-0/00"), " (" & Retorna_Cnae(sCnaeSec) & ")", UBound(aMei), Val(sTipo)
        End If
        t = t + 1
    Next
    If (Not IsNumeric(Left(sNome, 1))) Then
        lblFileName.Caption = sNome
        lblFileNumber.Caption = "Importando Empresa " & nPos & " de " & Int(nTot)
        nPos = nPos + 1
    End If
    
proximo:
Loop
Close #16
frProgress.Visible = False

For nPos = 1 To UBound(aMei)
    bFind = False
    For z = 0 To cmbReg.ListCount - 1
        If cmbReg.List(z) = aMei(nPos).Nome Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        cmbReg.AddItem aMei(nPos).Nome
    End If
Next

If cmbReg.ListCount > 0 Then cmbReg.ListIndex = 0
Liberado
Exit Sub
Erro:

If rdoErrors(1).Number = 2627 Then
    Resume Next
Else
    Liberado
    MsgBox rdoErrors(1).Description
End If

Exit Sub

End Sub

Private Sub AdicionaEvento(Codigo As String, Data As String, id As Integer, Tipo As Integer)
Dim nSize As Integer

nSize = UBound(aEvento) + 1
ReDim Preserve aEvento(nSize)
aEvento(nSize).id = id
aEvento(nSize).Tipo = Tipo
aEvento(nSize).Codigo = Codigo
aEvento(nSize).Data = Right(Data, 2) & "/" & Mid(Data, 5, 2) & "/" & Left(Data, 4)
aEvento(nSize).Nome = Retorna_Mei(Val(Codigo))

End Sub

Private Sub AdicionaCnae(Cnae As String, Nome As String, id As Integer, Tipo As Integer)
Dim nSize As Integer

nSize = UBound(aCnae) + 1
ReDim Preserve aCnae(nSize)
aCnae(nSize).id = id
aCnae(nSize).Tipo = Tipo
aCnae(nSize).Cnae = Cnae
aCnae(nSize).Nome = Nome

End Sub


Private Sub cmdGravar_Click()
PopupMenu mnuGravar, , cmdGravar.Left, cmdGravar.Top + cmdGravar.Height

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaTabela
CarregaNatureza
CarregaQualificacao
CarregaMunicipio
CarregaCnae
End Sub

Private Sub Fonte(Alinhamento As RichTextLib.SelAlignmentConstants, Cor As Long, Size As Integer, Negrito As Boolean, Italico As Boolean, Sublinhado As Boolean)

With Rtb
    .SelAlignment = Alinhamento
    .SelColor = Cor
    .SelFontSize = Size
    .SelBold = Negrito
    .SelUnderline = Sublinhado
    .SelItalic = Italico
End With

End Sub

Private Sub Negrito()
Rtb.SelBold = True
End Sub

Private Sub Normal()
Rtb.SelBold = False
End Sub

Private Sub CarregaTabela()
Dim Sql As String, RdoAux As rdoResultset

ReDim aTabela(0)
Sql = "select codigo,nome from evento_mei order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aTabela(UBound(aTabela) + 1)
        aTabela(UBound(aTabela)).Codigo = !Codigo
        aTabela(UBound(aTabela)).Nome = !Nome
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaNatureza()
Dim Sql As String, RdoAux As rdoResultset

ReDim aNatureza(0)
Sql = "select codigo,nome from natureza_juridica order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aNatureza(UBound(aNatureza) + 1)
        aNatureza(UBound(aNatureza)).Codigo = !Codigo
        aNatureza(UBound(aNatureza)).Nome = !Nome
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaCnae()
Dim Sql As String, RdoAux As rdoResultset

ReDim aTabelaCnae(0)
Sql = "select cnae,descricao from cnae order by cnae"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aTabelaCnae(UBound(aTabelaCnae) + 1)
        aTabelaCnae(UBound(aTabelaCnae)).Codigo = !Cnae
        aTabelaCnae(UBound(aTabelaCnae)).Nome = !descricao
       .MoveNext
    Loop
   .Close
End With

End Sub


Private Sub CarregaQualificacao()
Dim Sql As String, RdoAux As rdoResultset

ReDim aQualificacao(0)
Sql = "select codigo,nome from qualificacao_responsavel order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aQualificacao(UBound(aQualificacao) + 1)
        aQualificacao(UBound(aQualificacao)).Codigo = !Codigo
        aQualificacao(UBound(aQualificacao)).Nome = !Nome
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaMunicipio()
Dim Sql As String, RdoAux As rdoResultset

ReDim aMunicipio(0)
Sql = "select codigo,nome,uf from municipio order by uf,codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aMunicipio(UBound(aMunicipio) + 1)
        aMunicipio(UBound(aMunicipio)).Codigo = !Codigo
        aMunicipio(UBound(aMunicipio)).Nome = !Nome
        aMunicipio(UBound(aMunicipio)).UF = !UF
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Function Retorna_Municipio(UF As String, Codigo As Integer) As String
Dim sRet As String, x As Integer, bFind As Boolean

bFind = False
For x = 1 To UBound(aMunicipio)
    If aMunicipio(x).UF = UF And aMunicipio(x).Codigo = Codigo Then
        bFind = True
        Exit For
    End If
Next

If bFind Then
    sRet = aMunicipio(x).Nome
Else
    sRet = "Município não cadastrado!"
End If

Retorna_Municipio = sRet
End Function

Private Function Retorna_Cnae(Codigo As String) As String
Dim sRet As String, x As Integer, bFind As Boolean, sTmp As String

sTmp = RetornaNumero(Codigo)
bFind = False
For x = 1 To UBound(aTabelaCnae)
    If aTabelaCnae(x).Codigo = sTmp Then
        bFind = True
        Exit For
    End If
Next

If bFind Then
    sRet = aTabelaCnae(x).Nome
Else
    sRet = "Cnae não cadastrado!"
End If

Retorna_Cnae = sRet
End Function


Private Function Retorna_Mei(Codigo As Integer) As String
Dim sRet As String, x As Integer, bFind As Boolean

bFind = False
For x = 1 To UBound(aTabela)
    If aTabela(x).Codigo = Codigo Then
        bFind = True
        Exit For
    End If
Next

If bFind Then
    sRet = aTabela(x).Nome
Else
    sRet = "Evento não cadastrado!"
End If

Retorna_Mei = sRet
End Function

Private Function Retorna_Natureza(Codigo As Integer) As String
Dim sRet As String, x As Integer, bFind As Boolean

bFind = False
For x = 1 To UBound(aNatureza)
    If aNatureza(x).Codigo = Codigo Then
        bFind = True
        Exit For
    End If
Next

If bFind Then
    sRet = aNatureza(x).Nome
Else
    sRet = "Natureza não cadastrada!"
End If

Retorna_Natureza = sRet
End Function

Private Function Retorna_Qualificacao(Codigo As Integer) As String
Dim sRet As String, x As Integer, bFind As Boolean

bFind = False
For x = 1 To UBound(aQualificacao)
    If aQualificacao(x).Codigo = Codigo Then
        bFind = True
        Exit For
    End If
Next

If bFind Then
    sRet = aQualificacao(x).Nome
Else
    sRet = "Qualificacao não cadastrada!"
End If

Retorna_Qualificacao = sRet
End Function



Private Sub cmdGerarIM_Click()
Dim sCNPJ As String, nPos As Integer, sNome As String, i As Integer, Sql As String, RdoAux As rdoResultset
Dim MinCod As Long, MaxCod As Long, nCodCidadao As Long, sAtivExtenso As String, sCnae As String, sSecao As String
Dim nDivisao As Integer, nGrupo As Integer, sClasse As String, nClasse As Integer, nSubClasse As Integer

If cmbReg.ListIndex = -1 Then Exit Sub
sNome = cmbReg.Text

For nPos = 1 To UBound(aMei)
    With aMei(nPos)
        If .Nome = sNome Then
            sCNPJ = RetornaNumero(.CnpjMatriz)
            Sql = "select * from mobiliario where cnpj='" & sCNPJ & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    MsgBox "CNPJ já cadastrado (IM: " & !codigomob & ").", vbCritical, "Não é possível gravar."
                   .Close
                    Exit Sub
                End If
               .Close
            End With
            
            Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB>100000 and CODIGOMOB<200000 ORDER BY CODIGOMOB"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                Do Until .EOF
                   If MinCod = 0 Then
                      MinCod = !codigomob
                   Else
                      MaxCod = !codigomob
                      If MaxCod - MinCod > 1 Then
                          MaxCod = MinCod + 1
                          Exit Do
                      Else
                          MinCod = MaxCod
                      End If
                   End If
                  .MoveNext
                Loop
               .Close
            End With
            Sql = "insert mobiliario (codigomob,razaosocial,nomefantasia,codlogradouro,numero,complemento,siglauf,codcidade,cnpj,codbairro,capitalsocial,dataabertura,fonecontato,emailcontato,numregistroresp,ddd_nf,telefone_nf,email_nf)"
            Sql = Sql & " values(" & MaxCod & ",'" & Mask(.Nome) & "'," & IIf(.Fantasia = "", "Null", "'" & Mask(.Fantasia) & "'") & "," & .EnderecoCodigo & "," & RetornaNumero(.EnderecoNumero) & ","
            Sql = Sql & IIf(.EnderecoComplemento = "", "Null", "'" & Mask(.EnderecoComplemento) & "'") & ",'" & .UF & "'," & .MunicipoCodigo & ",'" & RetornaNumero(.CnpjMatriz) & "'," & IIf(.BairroCodigo = 0, "Null", .BairroCodigo) & ","
            Sql = Sql & Virg2Ponto(CStr(.Capital)) & "," & IIf(Not IsDate(.DataAbertura), "Null", "'" & Format(.DataAbertura, ("mm/dd/yyyy")) & "'") & "," & IIf(.Telefone1 = "", "Null", "'" & Mask(.Telefone1) & "'") & ","
            Sql = Sql & IIf(.Email = "", "Null", "'" & Mask(.Email) & "'") & "," & IIf(.Nire = "", "Null", "'" & Mask(.Nire) & "'") & "," & IIf(.DDD = "", "Null", "'" & .DDD & "'") & "," & IIf(.TelefoneNF = "", "Null", "'" & .TelefoneNF & "'") & "," & IIf(.Email = "", "Null", "'" & Mask(.Email) & "'") & ")"
            cn.Execute Sql, rdExecDirect
            
            Sql = "select * from cidadao where cpf='" & RetornaNumero(.CpfResp) & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux.RowCount > 0 Then
                nCodCidadao = RdoAux!CodCidadao
            Else
                nCodCidadao = 0
            End If
            RdoAux.Close
            
            If nCodCidadao > 0 Then
                Sql = "insert mobiliarioproprietario(codmobiliario,codcidadao,principal) values(" & MaxCod & "," & nCodCidadao & ",1)"
                cn.Execute Sql, rdExecDirect
            End If
            
            If .Cnae <> "" Then
                sCnae = Left(.Cnae, 9)
                sSecao = ""
                nDivisao = Val(Left(sCnae, 2))
                nGrupo = Val(Mid(sCnae, 3, 1))
                sClasse = Mid(sCnae, 4, 3)
                sClasse = Left(sClasse, 1) & Right(sClasse, 1)
                nClasse = Val(sClasse)
                nSubClasse = Val(Right(sCnae, 2))
                
                Sql = "INSERT MOBILIARIOCNAE(CODMOBILIARIO,SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,PRINCIPAL,CNAE) VALUES("
                Sql = Sql & MaxCod & ",'" & sSecao & "'," & nDivisao & "," & nGrupo & "," & nClasse & "," & nSubClasse & "," & 1 & ",'" & sCnae & "')"
                cn.Execute Sql, rdExecDirect
                    
                sAtivExtenso = Retorna_Cnae(sCnae) & ";"
            End If
            
            
            
            For i = 1 To UBound(aCnae)
                With aCnae(i)
                    If .id = nPos And .Tipo = 4 Then
                        sCnae = Left(.Cnae, 9)
                        sSecao = ""
                        nDivisao = Val(Left(sCnae, 2))
                        nGrupo = Val(Mid(sCnae, 3, 1))
                        sClasse = Mid(sCnae, 4, 3)
                        sClasse = Left(sClasse, 1) & Right(sClasse, 1)
                        nClasse = Val(sClasse)
                        nSubClasse = Val(Right(sCnae, 2))
                        Sql = "INSERT MOBILIARIOCNAE(CODMOBILIARIO,SECAO,DIVISAO,GRUPO,CLASSE,SUBCLASSE,PRINCIPAL,CNAE) VALUES("
                        Sql = Sql & MaxCod & ",'" & sSecao & "'," & nDivisao & "," & nGrupo & "," & nClasse & "," & nSubClasse & "," & 0 & ",'" & sCnae & "')"
                        cn.Execute Sql, rdExecDirect
                            
                        sAtivExtenso = sAtivExtenso & Retorna_Cnae(sCnae) & ";"
                    End If
                End With
            Next
            
            Sql = "update mobiliario set ativextenso='" & UCase(Mask(sAtivExtenso)) & "' where codigomob=" & MaxCod
            cn.Execute Sql, rdExecDirect
            
            MsgBox "Inscrição Municipal: " & MaxCod & " criada com sucesso.", vbInformation, "Atenção"
            
        End If
    End With
Next

End Sub

Private Sub mnuEmpresa_Click()
    If cmbReg.ListIndex > -1 Then
        Rtb.SaveFile sPathBin & "\" & cmbReg.Text & ".Rtf"
        MsgBox "Arquivo salvo em " & sPathBin & "\" & cmbReg.Text & ".Rtf", vbInformation, "Gravando dados"
    End If
End Sub

Private Sub mnuTodos_Click()
Dim x As Integer
If cmbReg.ListCount > 0 Then
    For x = 0 To cmbReg.ListCount - 1
        cmbReg.ListIndex = x
        Rtb.SaveFile sPathBin & "\" & cmbReg.Text & ".Rtf"
        
    Next
    MsgBox "Todos os arquivos foram salvos em " & sPathBin, vbInformation, "Gravando dados"
End If

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
DoEvents

End Sub

