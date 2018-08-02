VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmExporta 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integração com o Sistema de ISS Eletrônico"
   ClientHeight    =   2235
   ClientLeft      =   9240
   ClientTop       =   2505
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   7170
   Begin VB.CommandButton Command3 
      Caption         =   "NF"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3960
      TabIndex        =   28
      Top             =   2430
      Width           =   870
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2880
      TabIndex        =   27
      Top             =   2475
      Width           =   870
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Int"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1710
      TabIndex        =   26
      Top             =   2475
      Width           =   765
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup"
      Enabled         =   0   'False
      Height          =   285
      Left            =   225
      TabIndex        =   24
      Top             =   2460
      Width           =   1125
   End
   Begin VB.CommandButton cmdCanceladas 
      Caption         =   "Cancel"
      Height          =   240
      Left            =   180
      TabIndex        =   22
      Top             =   3375
      Width           =   1005
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1890
      Width           =   4740
   End
   Begin prjChameleon.chameleonButton cmdExport 
      Height          =   1005
      Left            =   180
      TabIndex        =   17
      ToolTipText     =   "Exporta o arquivo de dados do GTI para o ISS Eletrônico"
      Top             =   405
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1773
      BTYPE           =   14
      TX              =   "Exportar Guias"
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
      FOCUSR          =   0   'False
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmExporta.frx":0000
      PICN            =   "frmExporta.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar Pb 
      Height          =   240
      Left            =   135
      TabIndex        =   16
      Top             =   1575
      Width           =   4515
      _ExtentX        =   7964
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
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6660
      TabIndex        =   15
      Top             =   1890
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6165
      TabIndex        =   14
      Top             =   1890
      Width           =   375
   End
   Begin VB.ListBox lstISS 
      Height          =   1425
      Left            =   5460
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   360
      Width           =   1605
   End
   Begin VB.ListBox lstOpc 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1155
      ItemData        =   "frmExporta.frx":02F4
      Left            =   1590
      List            =   "frmExporta.frx":0307
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   3630
      Width           =   2775
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Exportar"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1620
      TabIndex        =   4
      Top             =   3120
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Importar"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2820
      TabIndex        =   3
      Top             =   3120
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Top             =   5085
      Width           =   1515
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   360
      Left            =   6120
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   5490
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "&Sair"
      ENAB            =   0   'False
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
      MICON           =   "frmExporta.frx":038A
      PICN            =   "frmExporta.frx":03A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   360
      Left            =   4740
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Emitir Relatório"
      Top             =   5490
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Executar"
      ENAB            =   0   'False
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmExporta.frx":0414
      PICN            =   "frmExporta.frx":0430
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdImport 
      Height          =   1005
      Left            =   1200
      TabIndex        =   18
      ToolTipText     =   "Importa as guias geradas pelo sistema de ISS Eletrônico"
      Top             =   405
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1773
      BTYPE           =   14
      TX              =   "Importar Guias"
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
      FOCUSR          =   0   'False
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmExporta.frx":04CF
      PICN            =   "frmExporta.frx":04EB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdImportSMov 
      Height          =   1005
      Left            =   2220
      TabIndex        =   19
      ToolTipText     =   "Importar arquivo sem movimento"
      Top             =   405
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1773
      BTYPE           =   14
      TX              =   "Importar S/Movim."
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
      FOCUSR          =   0   'False
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmExporta.frx":07F8
      PICN            =   "frmExporta.frx":0814
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdImportNF 
      Height          =   1005
      Left            =   3240
      TabIndex        =   20
      ToolTipText     =   "Importa as Notas Fiscais cadastradas no Sistema de ISS Eletrônico"
      Top             =   405
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1773
      BTYPE           =   14
      TX              =   " Importar Notas"
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
      FOCUSR          =   0   'False
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmExporta.frx":091E
      PICN            =   "frmExporta.frx":093A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdImportarSenha 
      Height          =   1005
      Left            =   4275
      TabIndex        =   23
      ToolTipText     =   "Importar as senhas do Sistema de ISS Eletrônico"
      Top             =   405
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   1773
      BTYPE           =   14
      TX              =   " Importar Simples"
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
      FOCUSR          =   0   'False
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmExporta.frx":0A3A
      PICN            =   "frmExporta.frx":0A56
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblArq 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   255
      Left            =   180
      TabIndex        =   25
      Top             =   2790
      Width           =   1875
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresas com Retenção"
      Height          =   255
      Left            =   5295
      TabIndex        =   13
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Registros.....:"
      Height          =   225
      Index           =   0
      Left            =   4440
      TabIndex        =   10
      Top             =   3945
      Width           =   1665
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total de Concluidos...:"
      Height          =   225
      Index           =   1
      Left            =   4440
      TabIndex        =   9
      Top             =   4260
      Width           =   1665
   End
   Begin VB.Label lblRegTot 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6150
      TabIndex        =   8
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Label lblRegPerc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6150
      TabIndex        =   7
      Top             =   4260
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Exportado..........:"
      Height          =   225
      Index           =   2
      Left            =   4440
      TabIndex        =   6
      Top             =   3630
      Width           =   1665
   End
   Begin VB.Label lblTipo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6150
      TabIndex        =   5
      Top             =   3660
      Width           =   1125
   End
End
Attribute VB_Name = "frmExporta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, ax As String, bCancel As Boolean
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nVVP As Double, nVVT As Double, nAreaTotal As Double
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
Private Type ATIVIDADES
    nCodigo As Long
    nSeq As Integer
    sEstimado As String
    sFixo As String
    sISSEletronico As String
    sTipoIss As String
End Type

Private Type ISSELETRO
    Identificacao As String
    Numero As String
    Inscricao As String
    Sequencia As String
    Ano As String
    Mes As String
    Aliquota As Double
    Tipo As String
    DataVencto As String
    ValorPrincipal As Double
    ValorAcrescimo As Double
    DataExporta As String
    Filler As String
End Type

Private Type ISSELETRONOVO
    nTipoLinha As Integer '1-cabeçalho,2-detalhe,3-rodapé
    sDataGeracao As String
    sHoraGeracao As String
    nNumeroDaGuia As Long
    nSequencia As Integer
    nTipoDeEmissao As Integer 'Link to Tabela TipoEmissao
    sSimplesNacional As String 'S ou N
    nExercicio As Integer
    nMes As Integer
    nAliquota As Double
    nCodReduz As Long
    sInscricaoT As String
    sRazaoSocialT As String
    sCPFT As String
    sCNPJT As String
    sEnderecoT As String
    sComplEndT As String
    sInscricaoP As String
    sRazaoSocialP As String
    sCPFP As String
    sCNPJP As String
    sEnderecoP As String
    sComplEndP As String
    nValorMovimento As Double
    nValorImposto As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    sDataEmissao As String
    sDataVencimento As String
    sUsuario As String
    sAtivIss As String
    sAtivSeq As String
    nStatus As Integer '0-gerada,1-impressa,4-paga,9-cancelada
    nQtdeLinhas As Long
    nSomaTotalGuias As Double
    nQtdeSemMov As Integer
    nER As Integer '0-emitida,1-recebida (somente para guia avulsa)
End Type

Private Type GUIAAVULSA
    nDoc As Long
    nCodReduz As Long
End Type

Private Type DATASIMPLES
    sDataIni As String
    sDataFim As String
End Type

Private Type NOTAS
    IdentificaPrestador As String
    TipoPrestador As Integer
    TipoNota As Integer
    NumeroNota As Double
    Serie As String
    DataEmissao As String
    MesRef As Integer
    AnoRef As Integer
    StatusNota As Integer
    DataCancel As String
    Natureza As String
    ValorTotal As Double
    ValorServico As Double
    ValorImposto As Double
    Recolhimento As Integer
    Atividade As Double
    Aliquota As Double
    RazaoPrestador As String
    CidadePrestador As String
    UFPrestador As String
    LocalPrestador As String
    IdentificaTomador As String
    TipoTomador As String
    RazaoTomador As String
    CidadeTomador As String
    UFTomador As String
    LocalTomador As String
    NumGuia As Long
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
Dim aSimples() As DATASIMPLES

Private Sub cmdBackup_Click()
Dim sPathBackup As String, strLinha As String, nRowCount As Long, nPos As Long
Dim amibc_im As Long, amibc_ano As Integer, amibc_mes As Integer, amibc_dt As String
Dim amibc_tm As Long, amibc_dt_ger As String, amibc_tm_ger As Long, amibc_banco As Integer
Dim amibc_cnpj As String, amibc_tp_doc As Integer, amibc_versao As Integer
Dim amibcv_im As Long, amibcv_ano As Integer, amibcv_mes As Integer, amibcv_grupo As Integer
Dim amibcv_sub_grupo As Integer, amibcv_conta As String, amibcv_ativ As Integer, amibcv_saldo As Double
Dim amibcv_aliquota As Double, amibcv_vl_imp As Double, amibcv_guia As String, amibcv_ativ_seq As Integer
Dim amide_im As Long, amide_ativ As Long, amide_mes As Integer, amide_ano As Integer, amide_vl_total As Double
Dim amide_vl_imp As Double, amide_guia As String, amide_aliquota As Double, amide_ativ_seq As Integer, amide_vl_fonte As Double
Dim amide_vl_munic As Double, amide_vl_estim As Double, amide_vl_liquido As Double, amide_dt_incl As String
Dim amies_im As Long, amies_ano As Integer, amies_mes As Integer, amies_ativ As Integer, amies_ativ_seq As Integer
Dim amies_vl_serv As Double, amies_vl_reduc As Double, amies_vl_isento As Double, amies_vl_munic As Double
Dim amies_aliq As Double, amies_vl_fonte As Double, amies_vl_total As Double, amies_guia As String, amies_dt_incl As String
Dim amigu_num As String, amigu_seq As Integer, amigu_id As String, amigu_tipo As Integer, amigu_dt_venc As String, amigu_dt_emiss As String
Dim amigu_vl_serv As Double, amigu_vl_princ As Double, amigu_vl_corr As Double, amigu_vl_multa As Double, amigu_vl_juros As Double
Dim amigu_vl_outros As Double, amigu_dt_pagto As String, amigu_pg_princ As Double, amigu_pg_adic As Double, amigu_status As Integer, amigu_tp_guia As Integer
Dim amine_id As String, amine_tipo As Integer, amine_nb As Integer, amine_serie As Integer, amine_tp As Integer, amine_end As String
Dim amine_st As Integer, amine_bairro As String, amine_cep As String, amine_mail As String, amine_rps As Long, amine_rps_tp As Integer
Dim amine_rps_serie As String, amine_time As Long, amine_rps_dt As Date
Dim aminea_id As String, aminea_tipo As Integer, aminea_nb As Integer, aminea_serie As Integer, aminea_tp As Integer, aminea_seq
Dim aminea_ativ As Integer, aminea_ativ_seq As Integer, aminea_aliquota As Double, aminea_vl_serv As Double, aminea_vl_imp As Double
Dim amineo_id As String, amineo_tipo As Integer, amineo_nb As Integer, amineo_serie As Integer, amineo_tp As Integer, amineo_seq, amineo_obs As String
Dim aminf_usr As String, aminf_id As String, aminf_tipo As Integer, aminf_nb As String, aminf_serie As Integer, aminf_tp As Integer, aminf_dt_emiss As String
Dim aminf_mes As Integer, aminf_ano As Integer, aminf_st As Integer, aminf_dt_cancel As String, aminf_nat As Integer, aminf_vl_total As Double
Dim aminf_vl_serv As Double, aminf_ativ As String, aminf_ativ_seq As Integer, aminf_aliquota As Double, aminf_vl_imp As Double
Dim aminf_rec_imp As Integer, aminf_prest_razao As String, aminf_prest_cidade As String, aminf_prest_uf As String, aminf_prest_fora As Integer
Dim aminf_tom_id As String, aminf_tom_razao As String, aminf_tom_cidade As String, aminf_tom_uf As String, aminf_tom_fora As Integer
Dim aminf_guia As String, aminf_ativ_emp As Integer, aminf_fecha As String, aminf_dt_incl As String, aminf_sim_nac As Integer, aminf_guia_seq As Integer
Dim aminfc_id As String, aminfc_tipo As Integer, aminfc_nb As String, aminfc_serie As Integer, aminfc_tp As Integer, aminfc_dt_emiss As String
Dim aminfc_mes As Integer, aminfc_ano As Integer, aminfc_st As Integer, aminfc_dt_cancel As String, aminfc_nat As Integer, aminfc_vl_total As Double
Dim aminfc_vl_serv As Double, aminfc_ativ As String, aminfc_ativ_seq As Integer, aminfc_aliquota As Double, aminfc_vl_imp As Double, aminfc_rec_imp As Integer
Dim aminfc_prest_razao As String, aminfc_prest_cidad As String, aminfc_prest_uf As String, aminfc_prest_fora As Integer, aminfc_tom_id As String
Dim aminfc_tom_tipo As Integer, aminfc_tom_razao As String, aminfc_tom_cidade As String, aminfc_tom_uf As String, aminfc_tom_fora As Integer
Dim aminfc_guia As String, aminfc_ativ_emp As Integer, aminfc_fecha As String, aminfc_dt_incl As String, aminfc_sim_nac As Integer, aminfc_guia_seq As Integer
Dim amiqn_guia As String, amiqn_prest_id As String, amiqn_prest_tp As Integer, amiqn_prest_razao As String, amiqn_prest_fora As Integer, amiqn_tom_id As String
Dim amiqn_tom_tp As Integer, amiqn_tom_razao As String, amiqn_tom_fora As Integer, amiqn_mes As Integer, amiqn_ano As Integer, amiqn_tipo As Integer, amiqn_dt_incl As String
Dim amist_im As Long, amist_mes As Integer, amist_ano As Integer, amist_st As Integer, amist_dt_incl As Date

GoTo fim

sPathBackup = "C:\Trabalho\GTI\Diversos\tabelas e catalog\"
GoTo AMIST

Sql = "delete from obvius_amibc"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
Open sPathBackup & "amibc.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        amibc_im = Val(RemoveSpace(Mid(strLinha, 1, 12)))
        amibc_ano = Val(RemoveSpace(Mid(strLinha, 16, 5)))
        amibc_mes = Val(RemoveSpace(Mid(strLinha, 25, 2)))
        amibc_dt = RevertDate(RemoveSpace(Mid(strLinha, 31, 10)))
        amibc_tm = Val(RemoveSpace(Mid(strLinha, 45, 7)))
        amibc_dt_ger = RevertDate(RemoveSpace(Mid(strLinha, 56, 10)))
        amibc_tm_ger = Val(RemoveSpace(Mid(strLinha, 70, 7)))
        amibc_banco = Val(Mid(strLinha, 81, 3))
        amibc_cnpj = RemoveSpace(Mid(strLinha, 88, 18))
        amibc_tp_doc = Val(Mid(strLinha, 109, 1))
        amibc_versao = Val(Mid(strLinha, 116, 1))
        Sql = "insert obvius_amibc(amibc_im,amibc_ano,amibc_mes,amibc_dt,amibc_tm,amibc_dt_ger,amibc_tm_ger,amibc_banco,"
        Sql = Sql & "amibc_cnpj,amibc_tp_doc,amibc_versao) values(" & amibc_im & "," & amibc_ano & "," & amibc_mes & ",'"
        Sql = Sql & Format(amibc_dt, "mm/dd/yyyy") & "'," & amibc_tm & ",'" & Format(amibc_dt_ger, "mm/dd/yyyy") & "'," & amibc_tm_ger & "," & amibc_banco & ",'" & amibc_cnpj & "',"
        Sql = Sql & amibc_tp_doc & "," & amibc_versao & ")"
        cn.Execute Sql, rdExecDirect
        DoEvents
   Loop
Close #FF1

AMIBCV:
lblArq.Caption = "amibcv.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amibcv.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0


Sql = "delete from obvius_amibcv"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amibcv.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amibcv_im = Val(RemoveSpace(Mid(strLinha, 1, 12)))
        amibcv_ano = Val(RemoveSpace(Mid(strLinha, 16, 5)))
        amibcv_mes = Val(RemoveSpace(Mid(strLinha, 25, 2)))
        amibcv_grupo = Val(RemoveSpace(Mid(strLinha, 30, 16)))
        amibcv_sub_grupo = Val(RemoveSpace(Mid(strLinha, 49, 16)))
        amibcv_conta = RemoveSpace(Mid(strLinha, 68, 12))
        amibcv_ativ = Val(RemoveSpace(Mid(strLinha, 83, 11)))
        amibcv_saldo = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 97, 25))))
        amibcv_aliquota = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 125, 6))))
        amibcv_vl_imp = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 134, 25))))
        amibcv_guia = RemoveSpace(Mid(strLinha, 163, 15))
        amibcv_ativ_seq = Val(Mid(strLinha, 181, 4))
        Sql = "insert obvius_amibcv(amibcv_im,amibcv_ano,amibcv_mes,amibcv_grupo,amibcv_sub_grupo,amibcv_conta,amibcv_ativ,amibcv_saldo,"
        Sql = Sql & "amibcv_aliquota,amibcv_vl_imp,amibcv_guia,amibcv_ativ_seq) values(" & amibcv_im & "," & amibcv_ano & "," & amibcv_mes & ","
        Sql = Sql & amibcv_grupo & "," & amibcv_sub_grupo & ",'" & amibcv_conta & "'," & amibcv_ativ & "," & Virg2Ponto(CStr(amibcv_saldo)) & "," & Virg2Ponto(CStr(amibcv_aliquota)) & ","
        Sql = Sql & Virg2Ponto(CStr(amibcv_vl_imp)) & "," & Val(amibcv_guia) & "," & amibcv_ativ_seq & ")"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1

GoTo fim
AMIDE:
lblArq.Caption = "amide.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amide.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0


Sql = "delete from obvius_amide"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amide.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amide_im = Val(RemoveSpace(Mid(strLinha, 1, 12)))
        amide_ativ = Val(RemoveSpace(Mid(strLinha, 15, 11)))
        amide_mes = Val(RemoveSpace(Mid(strLinha, 30, 2)))
        amide_ano = Val(RemoveSpace(Mid(strLinha, 36, 5)))
        amide_vl_total = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 44, 15))))
        amide_vl_imp = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 62, 15))))
        amide_guia = RemoveSpace(Mid(strLinha, 81, 15))
        amide_aliquota = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 99, 6))))
        amide_ativ_seq = Val(RemoveSpace(Mid(strLinha, 108, 4)))
        amide_vl_fonte = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 115, 15))))
        amide_vl_munic = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 133, 15))))
        amide_vl_estim = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 151, 15))))
        amide_vl_liquido = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 169, 15))))
        amide_dt_incl = RevertDate(RemoveSpace(Mid(strLinha, 188, 10)))
        
        Sql = "insert obvius_amide(amide_im,amide_ativ,amide_mes,amide_ano,amide_vl_total,amide_vl_imp,amide_guia,amide_aliquota,"
        Sql = Sql & "amide_ativ_seq,amide_vl_fonte,amide_vl_munic,amide_vl_estim,amide_vl_liquido,amide_dt_incl) values(" & amide_im & "," & amide_ativ & "," & amide_mes & ","
        Sql = Sql & amide_ano & "," & Virg2Ponto(CStr(amide_vl_total)) & "," & Virg2Ponto(CStr(amide_vl_imp)) & "," & amide_guia & "," & Virg2Ponto(CStr(amide_aliquota)) & ","
        Sql = Sql & amide_ativ_seq & "," & Virg2Ponto(CStr(amide_vl_fonte)) & "," & Virg2Ponto(CStr(amide_vl_munic)) & ","
        If RemoveSpace(Mid(strLinha, 188, 10)) = "" Then
            Sql = Sql & Virg2Ponto(CStr(amide_vl_estim)) & "," & Virg2Ponto(CStr(amide_vl_liquido)) & "," & "Null" & ")"
        Else
            Sql = Sql & Virg2Ponto(CStr(amide_vl_estim)) & "," & Virg2Ponto(CStr(amide_vl_liquido)) & ",'" & Format(amide_dt_incl, "mm/dd/yyyy") & "')"
        End If
        
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1
'Exit Sub

AMIES:
lblArq.Caption = "amies.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amies.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0


Sql = "delete from obvius_amies"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amies.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amies_im = Val(RemoveSpace(Mid(strLinha, 1, 12)))
        amies_ano = Val(RemoveSpace(Mid(strLinha, 16, 5)))
        amies_mes = Val(RemoveSpace(Mid(strLinha, 25, 2)))
        amies_ativ = Val(RemoveSpace(Mid(strLinha, 30, 11)))
        If amies_ativ = 0 Then GoTo ProximoAmies
        amies_ativ_seq = Val(RemoveSpace(Mid(strLinha, 44, 4)))
        amies_vl_serv = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 51, 15))))
        amies_vl_reduc = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 69, 15))))
        amies_vl_isento = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 87, 15))))
        amies_vl_munic = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 105, 15))))
        amies_aliq = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 123, 6))))
        amies_vl_fonte = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 132, 15))))
        amies_vl_total = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 150, 15))))
        amies_guia = RemoveSpace(Mid(strLinha, 169, 15))
        amies_dt_incl = RevertDate(RemoveSpace(Mid(strLinha, 188, 10)))
        
        Sql = "insert obvius_amies(amies_im,amies_ano,amies_mes,amies_ativ,amies_ativ_seq,amies_vl_serv,amies_vl_reduc,amies_vl_isento,"
        Sql = Sql & "amies_vl_munic,amies_aliq,amies_vl_fonte,amies_vl_total,amies_guia,amies_dt_incl) values(" & amies_im & "," & amies_ano & "," & amies_mes & ","
        Sql = Sql & amies_ativ & "," & amies_ativ_seq & "," & Virg2Ponto(CStr(amies_vl_serv)) & "," & Virg2Ponto(CStr(amies_vl_reduc)) & "," & Virg2Ponto(CStr(amies_vl_isento)) & ","
        Sql = Sql & Virg2Ponto(CStr(amies_vl_munic)) & "," & Virg2Ponto(CStr(amies_aliq)) & "," & Virg2Ponto(CStr(amies_vl_fonte)) & "," & Virg2Ponto(CStr(amies_vl_total)) & ","
        Sql = Sql & amies_guia & ","
        If Val(amies_dt_incl) = 0 Then
            Sql = Sql & "Null" & ")"
        Else
            Sql = Sql & "'" & Format(amies_dt_incl, "mm/dd/yyyy") & "')"
        End If
        
        cn.Execute Sql, rdExecDirect
ProximoAmies:
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1
'Exit Sub

AMIGU:
lblArq.Caption = "amigu.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amigu.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0


Sql = "delete from obvius_amigu"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amigu.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amigu_num = Val(RemoveSpace(Mid(strLinha, 2, 17)))
        amigu_seq = Val(RemoveSpace(Mid(strLinha, 20, 4)))
        amigu_id = RemoveSpace(Mid(strLinha, 27, 19))
        amigu_tipo = Val(RemoveSpace(Mid(strLinha, 49, 1)))
        amigu_dt_venc = RevertDate(RemoveSpace(Mid(strLinha, 54, 10)))
        If Not IsDate(amigu_dt_venc) Then amigu_dt_venc = "Null"
        
        amigu_dt_emiss = RevertDate(RemoveSpace(Mid(strLinha, 68, 10)))
        If Not IsDate(amigu_dt_emiss) Then amigu_dt_emiss = "Null"
        
        If RemoveSpace(Mid(strLinha, 82, 14)) <> "" Then
            amigu_vl_serv = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 81, 15))))
        Else
            amigu_vl_serv = 0
        End If
        
        If RemoveSpace(Mid(strLinha, 100, 14)) <> "" Then
            amigu_vl_princ = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 99, 15))))
        Else
            amigu_vl_princ = 0
        End If
        
        If RemoveSpace(Mid(strLinha, 118, 14)) <> "" Then
            amigu_vl_corr = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 117, 15))))
        Else
            amigu_vl_corr = 0
        End If
        
        If RemoveSpace(Mid(strLinha, 136, 14)) <> "" Then
            amigu_vl_multa = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 135, 15))))
        Else
            amigu_vl_multa = 0
        End If
        
        If RemoveSpace(Mid(strLinha, 154, 14)) <> "" Then
            amigu_vl_juros = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 153, 15))))
        Else
            amigu_vl_juros = 0
        End If
        
        If RemoveSpace(Mid(strLinha, 172, 14)) <> "" Then
            amigu_vl_outros = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 171, 15))))
        Else
            amigu_vl_outros = 0
        End If
            
        amigu_dt_pagto = RevertDate(RemoveSpace(Mid(strLinha, 190, 10)))
        If Not IsDate(amigu_dt_pagto) Then
            amigu_dt_pagto = "Null"
            amigu_pg_princ = 0
            amigu_pg_adic = 0
        Else
            amigu_pg_princ = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 204, 15))))
            amigu_pg_adic = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 222, 15))))
        End If
        amigu_status = Val(RemoveSpace(Mid(strLinha, 239, 2)))
        amigu_tp_guia = Val(RemoveSpace(Mid(strLinha, 244, 3)))
        
        Sql = "insert obvius_amigu(amigu_num,amigu_seq,amigu_id,amigu_tipo,amigu_dt_venc,amigu_dt_emiss,amigu_vl_serv,amigu_vl_princ,"
        Sql = Sql & "amigu_vl_corr,amigu_vl_multa,amigu_vl_juros,amigu_vl_outros,amigu_dt_pagto,amigu_pg_princ,amigu_pg_adic,amigu_status,amigu_tp_guia) values(" & amigu_num & "," & amigu_seq & ",'" & amigu_id & "'," & amigu_tipo
        If amigu_dt_venc <> "Null" Then
            Sql = Sql & ",'" & Format(amigu_dt_venc, "mm/dd/yyyy") & "','" & Format(amigu_dt_emiss, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(amigu_vl_serv)) & "," & Virg2Ponto(CStr(amigu_vl_princ)) & ","
        Else
            Sql = Sql & "," & "Null" & "," & "Null" & "," & Virg2Ponto(CStr(amigu_vl_serv)) & "," & Virg2Ponto(CStr(amigu_vl_princ)) & ","
        End If
        
        Sql = Sql & Virg2Ponto(CStr(amigu_vl_corr)) & "," & Virg2Ponto(CStr(amigu_vl_multa)) & "," & Virg2Ponto(CStr(amigu_vl_juros)) & "," & Virg2Ponto(CStr(amigu_vl_outros)) & ","
        If amigu_dt_pagto <> "Null" Then
            Sql = Sql & "'" & Format(amigu_dt_pagto, "mm/dd/yyyy") & "'," & Virg2Ponto(CStr(amigu_pg_princ)) & "," & Virg2Ponto(CStr(amigu_pg_adic)) & "," & amigu_status & "," & amigu_tp_guia & ")"
        Else
            Sql = Sql & "Null" & "," & 0 & "," & 0 & "," & amigu_status & "," & amigu_tp_guia & ")"
        End If
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1


AMINE:
lblArq.Caption = "amine.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amine.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0

Sql = "delete from obvius_amine"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amine.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amine_id = RemoveSpace(Mid(strLinha, 20, 19))
        amine_tipo = Val(RemoveSpace(Mid(strLinha, 42, 1)))
        amine_nb = Val(RemoveSpace(Mid(strLinha, 46, 19)))
        amine_serie = Val(RemoveSpace(Mid(strLinha, 68, 1)))
        amine_tp = Val(RemoveSpace(Mid(strLinha, 72, 1)))
        amine_end = Trim(Mid(strLinha, 76, 200))
        amine_st = Val(RemoveSpace(Mid(strLinha, 279, 1)))
        amine_bairro = Trim(Mid(strLinha, 283, 30))
        amine_cep = RemoveSpace(Mid(strLinha, 316, 11))
        amine_mail = Trim(Mid(strLinha, 330, 80))
        amine_rps = Val(RemoveSpace(Mid(strLinha, 413, 19)))
        amine_rps_tp = Val(RemoveSpace(Mid(strLinha, 435, 1)))
        amine_rps_serie = RemoveSpace(Mid(strLinha, 439, 5))
        amine_time = Val(RemoveSpace(Mid(strLinha, 447, 8)))
        If RemoveSpace(Mid(strLinha, 459, 10)) <> "0" And RemoveSpace(Mid(strLinha, 459, 10)) <> "" Then
            amine_rps_dt = RevertDate(RemoveSpace(Mid(strLinha, 459, 10)))
        Else
            amine_rps_dt = CDate("01/01/1900")
        End If
        
        Sql = "insert obvius_amine(amine_id,amine_tipo,amine_nb,amine_serie,amine_tp,amine_end,amine_st,amine_bairro,"
        Sql = Sql & "amine_cep,amine_mail,amine_rps,amine_rps_tp,amine_rps_serie,amine_time,amine_rps_dt) values('" & amine_id & "'," & amine_tipo & "," & amine_nb & ","
        Sql = Sql & amine_serie & "," & amine_tp & ",'" & Mask(amine_end) & "'," & amine_st & ",'" & Mask(amine_bairro) & "','"
        Sql = Sql & amine_cep & "','" & amine_mail & "'," & amine_rps & "," & amine_rps_tp & ",'" & Left(amine_rps_serie, 1) & "'," & amine_time & ","
        If amine_rps_dt <> "01/01/1900" And Year(CDate(amine_rps_dt)) > 1900 Then
            Sql = Sql & "'" & Format(amine_rps_dt, "mm/dd/yyyy") & "')"
        Else
            Sql = Sql & "Null" & ")"
        End If
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1

AMINEA:
lblArq.Caption = "aminea.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "aminea.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0

Sql = "delete from obvius_aminea"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "aminea.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        aminea_id = RemoveSpace(Mid(strLinha, 20, 19))
        aminea_tipo = Val(RemoveSpace(Mid(strLinha, 42, 1)))
        aminea_nb = Val(RemoveSpace(Mid(strLinha, 46, 19)))
        aminea_serie = Val(RemoveSpace(Mid(strLinha, 68, 1)))
        aminea_tp = Val(RemoveSpace(Mid(strLinha, 72, 1)))
        aminea_seq = Val(RemoveSpace(Mid(strLinha, 76, 3)))
        aminea_ativ = Val(RemoveSpace(Mid(strLinha, 82, 11)))
        aminea_ativ_seq = Val(RemoveSpace(Mid(strLinha, 96, 4)))
        aminea_aliquota = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 103, 6))))
        aminea_vl_serv = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 112, 15))))
        aminea_vl_imp = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 130, 15))))
        
        Sql = "insert obvius_aminea(aminea_id,aminea_tipo,aminea_nb,aminea_serie,aminea_tp,aminea_seq,aminea_ativ,aminea_ativ_seq,"
        Sql = Sql & "aminea_aliquota,aminea_vl_serv,aminea_vl_imp) values('" & aminea_id & "'," & aminea_tipo & "," & aminea_nb & ","
        Sql = Sql & aminea_serie & "," & aminea_tp & "," & aminea_seq & "," & aminea_ativ & "," & aminea_ativ_seq & ","
        Sql = Sql & Virg2Ponto(CStr(aminea_aliquota)) & "," & Virg2Ponto(CStr(aminea_vl_serv)) & "," & Virg2Ponto(CStr(aminea_vl_serv)) & ")"
       
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1

AMINEO:
lblArq.Caption = "amineo.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amineo.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0

Sql = "delete from obvius_amineo"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amineo.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amineo_id = RemoveSpace(Mid(strLinha, 20, 19))
        amineo_tipo = Val(RemoveSpace(Mid(strLinha, 42, 1)))
        amineo_nb = Val(RemoveSpace(Mid(strLinha, 46, 19)))
        amineo_serie = Val(RemoveSpace(Mid(strLinha, 68, 1)))
        amineo_tp = Val(RemoveSpace(Mid(strLinha, 72, 1)))
        amineo_seq = Val(RemoveSpace(Mid(strLinha, 76, 3)))
        amineo_obs = Trim(Mid(strLinha, 82, 250))
        
        Sql = "insert obvius_amineo(amineo_id,amineo_tipo,amineo_nb,amineo_serie,amineo_tp,amineo_seq,amineo_obs) values('" & amineo_id & "'," & amineo_tipo & "," & amineo_nb & ","
        Sql = Sql & amineo_serie & "," & amineo_tp & "," & amineo_seq & ",'" & Mask(amineo_obs) & "')"
       
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1
GoTo fim

AMINF:
lblArq.Caption = "aminf.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "aminf.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0

Sql = "delete from obvius_aminf"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "aminf.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
        
        aminf_usr = RemoveSpace(Mid(strLinha, 1, 16))
        aminf_id = RemoveSpace(Mid(strLinha, 20, 19))
        aminf_tipo = Val(RemoveSpace(Mid(strLinha, 42, 1)))
        aminf_nb = RemoveSpace(Mid(strLinha, 46, 19))
        aminf_serie = Val(RemoveSpace(Mid(strLinha, 68, 1)))
        aminf_tp = Val(RemoveSpace(Mid(strLinha, 72, 1)))
        If RemoveSpace(Mid(strLinha, 77, 10)) <> "0" And RemoveSpace(Mid(strLinha, 77, 10)) <> "" Then
            aminf_dt_emiss = RevertDate(RemoveSpace(Mid(strLinha, 77, 10)))
        Else
            aminf_dt_emiss = "01/01/1900"
        End If
        aminf_mes = Val(RemoveSpace(Mid(strLinha, 91, 2)))
        aminf_ano = Val(RemoveSpace(Mid(strLinha, 97, 5)))
        aminf_st = Val(RemoveSpace(Mid(strLinha, 105, 1)))
        If RemoveSpace(Mid(strLinha, 110, 10)) <> "0" And RemoveSpace(Mid(strLinha, 110, 10)) <> "" Then
            aminf_dt_cancel = RevertDate(RemoveSpace(Mid(strLinha, 110, 10)))
        Else
            aminf_dt_cancel = "01/01/1900"
        End If
        aminf_nat = Val(RemoveSpace(Mid(strLinha, 123, 1)))
        If RemoveSpace(Mid(strLinha, 127, 15)) <> "?" Then
            aminf_vl_total = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 127, 15))))
            aminf_vl_serv = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 145, 15))))
            aminf_ativ = Trim(Mid(strLinha, 163, 11))
            aminf_ativ_seq = Val(RemoveSpace(Mid(strLinha, 177, 4)))
            aminf_aliquota = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 184, 6))))
            aminf_vl_imp = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 193, 15))))
            aminf_rec_imp = Val(RemoveSpace(Mid(strLinha, 211, 1)))
        Else
            aminf_vl_total = 0
            aminf_vl_serv = 0
            aminf_ativ = 0
            aminf_ativ_seq = 0
            aminf_aliquota = 0
            aminf_vl_imp = 0
            aminf_rec_imp = 0
        End If
        aminf_prest_razao = Trim(Mid(strLinha, 215, 100))
        aminf_prest_cidade = Trim(Mid(strLinha, 318, 40))
        aminf_prest_uf = Trim(Mid(strLinha, 361, 2))
        aminf_prest_fora = Val(RemoveSpace(Mid(strLinha, 367, 1)))
        aminf_tom_id = RemoveSpace(Mid(strLinha, 370, 19))
        aminf_tom_tipo = Val(RemoveSpace(Mid(strLinha, 392, 1)))
        aminf_tom_razao = Trim(Mid(strLinha, 396, 100))
        aminf_tom_cidade = Trim(Mid(strLinha, 499, 40))
        aminf_tom_uf = Trim(Mid(strLinha, 542, 2))
        aminf_tom_fora = Val(RemoveSpace(Mid(strLinha, 547, 1)))
        aminf_guia = RemoveSpace(Mid(strLinha, 552, 15))
        aminf_ativ_emp = Val(RemoveSpace(Mid(strLinha, 570, 1)))
        If RemoveSpace(Mid(strLinha, 575, 10)) <> "0" And RemoveSpace(Mid(strLinha, 575, 10)) <> "" Then
            aminf_fecha = RevertDate(RemoveSpace(Mid(strLinha, 575, 10)))
        Else
            aminf_fecha = "01/01/1900"
        End If
        If RemoveSpace(Mid(strLinha, 589, 10)) <> "0" And RemoveSpace(Mid(strLinha, 589, 10)) <> "" Then
            aminf_dt_incl = RevertDate(RemoveSpace(Mid(strLinha, 589, 10)))
        Else
            aminf_dt_incl = "01/01/1900"
        End If
        aminf_sim_nac = Val(RemoveSpace(Mid(strLinha, 603, 1)))
        If RemoveSpace(Mid(strLinha, 607, 4)) <> "?" Then
            aminf_guia_seq = Val(RemoveSpace(Mid(strLinha, 607, 4)))
        Else
            aminf_guia_seq = 0
        End If
        On Error Resume Next
        Sql = "insert obvius_aminf(aminf_usr,aminf_id,aminf_tipo,aminf_nb,aminf_serie,aminf_tp,aminf_dt_emiss,aminf_mes,aminf_ano,aminf_st,aminf_dt_cancel,"
        Sql = Sql & "aminf_nat,aminf_vl_total,aminf_vl_serv,aminf_ativ,aminf_ativ_seq,aminf_aliquota,aminf_vl_imp,aminf_rec_imp,aminf_prest_razao,"
        Sql = Sql & "aminf_prest_cidade,aminf_prest_uf,aminf_prest_fora,aminf_tom_id,aminf_tom_tipo,aminf_tom_razao,aminf_tom_cidade,"
        Sql = Sql & "aminf_tom_uf,aminf_tom_fora,aminf_guia,aminf_ativ_emp,aminf_fecha,aminf_dt_incl,aminf_sim_nac,aminf_guia_seq) values('"
        Sql = Sql & aminf_usr & "','" & aminf_id & "'," & aminf_tipo & ",'" & aminf_nb & "'," & aminf_serie & "," & aminf_tp & ",'" & Format(aminf_dt_emiss, "mm/dd/yyyy") & "',"
        Sql = Sql & aminf_mes & "," & aminf_ano & "," & aminf_st & ",'" & Format(aminf_dt_cancel, "mm/dd/yyyy") & "'," & aminf_nat & "," & Virg2Ponto(CStr(aminf_vl_total)) & ","
        Sql = Sql & Virg2Ponto(CStr(aminf_vl_serv)) & ",'" & aminf_ativ & "'," & aminf_ativ_seq & "," & Virg2Ponto(CStr(aminf_aliquota)) & ","
        Sql = Sql & Virg2Ponto(CStr(aminf_vl_imp)) & "," & aminf_rec_imp & ",'" & Mask(aminf_prest_razao) & "','" & aminf_prest_cidade & "','"
        Sql = Sql & aminf_prest_uf & "'," & aminf_prest_fora & ",'" & aminf_tom_id & "'," & aminf_tom_tipo & ",'" & Mask(aminf_tom_razao) & "','"
        Sql = Sql & Mask(aminf_tom_cidade) & "','" & aminf_tom_uf & "'," & aminf_tom_fora & ",'" & aminf_guia & "'," & aminf_ativ_emp & ",'" & Format(aminf_fecha, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(aminf_dt_incl, "mm/dd/yyyy") & "'," & aminf_sim_nac & "," & aminf_guia_seq & ")"
        
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1
GoTo fim

AMINFC:
lblArq.Caption = "aminfc.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "aminfc.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0

Sql = "delete from obvius_aminfc"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "aminfc.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        aminfc_id = RemoveSpace(Mid(strLinha, 20, 19))
        aminfc_tipo = Val(RemoveSpace(Mid(strLinha, 42, 1)))
        aminfc_nb = Val(RemoveSpace(Mid(strLinha, 46, 19)))
        aminfc_serie = Val(RemoveSpace(Mid(strLinha, 68, 1)))
        aminfc_tp = Val(RemoveSpace(Mid(strLinha, 72, 1)))
        aminfc_dt_emiss = RevertDate(RemoveSpace(Mid(strLinha, 77, 10)))
        aminfc_mes = Val(RemoveSpace(Mid(strLinha, 91, 2)))
        aminfc_ano = Val(RemoveSpace(Mid(strLinha, 97, 5)))
        aminfc_st = Val(RemoveSpace(Mid(strLinha, 105, 1)))
        If RemoveSpace(Mid(strLinha, 110, 10)) <> "0" And RemoveSpace(Mid(strLinha, 110, 10)) <> "" Then
            aminfc_dt_cancel = RevertDate(RemoveSpace(Mid(strLinha, 110, 10)))
        Else
            aminfc_dt_cancel = "01/01/1900"
        End If
        aminfc_nat = Val(RemoveSpace(Mid(strLinha, 123, 1)))
        aminfc_vl_total = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 127, 15))))
        aminfc_vl_serv = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 145, 15))))
        aminfc_ativ = Trim(Mid(strLinha, 163, 11))
        aminfc_ativ_seq = Val(RemoveSpace(Mid(strLinha, 177, 4)))
        aminfc_aliquota = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 184, 6))))
        aminfc_vl_imp = CDbl(Ponto2Virg(RemoveSpace(Mid(strLinha, 193, 15))))
        aminfc_rec_imp = Val(RemoveSpace(Mid(strLinha, 211, 1)))
        aminfc_prest_razao = Trim(Mid(strLinha, 215, 100))
        aminfc_prest_cidade = Trim(Mid(strLinha, 318, 40))
        aminfc_prest_uf = Trim(Mid(strLinha, 361, 2))
        aminfc_prest_fora = Val(RemoveSpace(Mid(strLinha, 367, 1)))
        aminfc_tom_id = RemoveSpace(Mid(strLinha, 370, 19))
        aminfc_tom_tipo = Val(RemoveSpace(Mid(strLinha, 392, 1)))
        aminfc_tom_razao = Trim(Mid(strLinha, 396, 100))
        aminfc_tom_cidade = Trim(Mid(strLinha, 499, 40))
        aminfc_tom_uf = Trim(Mid(strLinha, 542, 2))
        aminfc_tom_fora = Val(RemoveSpace(Mid(strLinha, 547, 1)))
        aminfc_guia = RemoveSpace(Mid(strLinha, 552, 15))
        aminfc_ativ_emp = Val(RemoveSpace(Mid(strLinha, 570, 1)))
        If RemoveSpace(Mid(strLinha, 575, 10)) <> "0" And RemoveSpace(Mid(strLinha, 575, 10)) <> "" Then
            aminfc_fecha = RevertDate(RemoveSpace(Mid(strLinha, 575, 10)))
        Else
            aminfc_fecha = "01/01/1900"
        End If
        If RemoveSpace(Mid(strLinha, 589, 10)) <> "0" And RemoveSpace(Mid(strLinha, 589, 10)) <> "" Then
            aminfc_dt_incl = RevertDate(RemoveSpace(Mid(strLinha, 589, 10)))
        Else
            aminfc_dt_incl = "01/01/1900"
        End If
        aminfc_sim_nac = Val(RemoveSpace(Mid(strLinha, 603, 1)))
        If RemoveSpace(Mid(strLinha, 607, 4)) <> "?" Then
            aminfc_guia_seq = Val(RemoveSpace(Mid(strLinha, 607, 4)))
        Else
            aminfc_guia_seq = 0
        End If
        
        Sql = "insert obvius_aminfc(aminfc_id,aminfc_tipo,aminfc_nb,aminfc_serie,aminfc_tp,aminfc_dt_emiss,aminfc_mes,aminfc_ano,aminfc_st,aminfc_dt_cancel,"
        Sql = Sql & "aminfc_nat,aminfc_vl_total,aminfc_vl_serv,aminfc_ativ,aminfc_ativ_seq,aminfc_aliquota,aminfc_vl_imp,aminfc_rec_imp,aminfc_prest_razao,"
        Sql = Sql & "aminfc_prest_cidade,aminfc_prest_uf,aminfc_prest_fora,aminfc_tom_id,aminfc_tom_tipo,aminfc_tom_razao,aminfc_tom_cidade,"
        Sql = Sql & "aminfc_tom_uf,aminfc_tom_fora,aminfc_guia,aminfc_ativ_emp,aminfc_fecha,aminfc_dt_incl,aminfc_sim_nac,aminfc_guia_seq) values('"
        Sql = Sql & aminfc_id & "'," & aminfc_tipo & ",'" & aminfc_nb & "'," & aminfc_serie & "," & aminfc_tp & ",'" & Format(aminfc_dt_emiss, "mm/dd/yyyy") & "',"
        Sql = Sql & aminfc_mes & "," & aminfc_ano & "," & aminfc_st & ",'" & Format(aminfc_dt_cancel, "mm/dd/yyyy") & "'," & aminfc_nat & "," & Virg2Ponto(CStr(aminfc_vl_total)) & ","
        Sql = Sql & Virg2Ponto(CStr(aminfc_vl_serv)) & ",'" & aminfc_ativ & "'," & aminfc_ativ_seq & "," & Virg2Ponto(CStr(aminfc_aliquota)) & ","
        Sql = Sql & Virg2Ponto(CStr(aminfc_vl_imp)) & "," & aminfc_rec_imp & ",'" & Mask(aminfc_prest_razao) & "','" & aminfc_prest_cidade & "','"
        Sql = Sql & aminfc_prest_uf & "'," & aminfc_prest_fora & ",'" & aminfc_tom_id & "'," & aminfc_tom_tipo & ",'" & Mask(aminfc_tom_razao) & "','"
        Sql = Sql & Mask(aminfc_tom_cidade) & "','" & aminfc_tom_uf & "'," & aminfc_tom_fora & ",'" & aminfc_guia & "'," & aminfc_ativ_emp & ",'" & Format(aminfc_fecha, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(aminfc_dt_incl, "mm/dd/yyyy") & "'," & aminfc_sim_nac & "," & aminfc_guia_seq & ")"
        
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1

AMIQN:
lblArq.Caption = "amiqn.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amiqn.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0


Sql = "delete from obvius_amiqn"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amiqn.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amiqn_guia = RemoveSpace(Mid(strLinha, 2, 15))
        amiqn_prest_id = RemoveSpace(Mid(strLinha, 20, 19))
        amiqn_prest_tp = Val(RemoveSpace(Mid(strLinha, 43, 1)))
        amiqn_prest_razao = Trim(Mid(strLinha, 47, 100))
        amiqn_prest_fora = Val(RemoveSpace(Mid(strLinha, 151, 1)))
        amiqn_tom_id = RemoveSpace(Mid(strLinha, 155, 19))
        amiqn_tom_tp = Val(RemoveSpace(Mid(strLinha, 178, 1)))
        amiqn_tom_razao = Trim(Mid(strLinha, 182, 100))
        amiqn_tom_fora = Val(RemoveSpace(Mid(strLinha, 286, 1)))
        amiqn_mes = Val(RemoveSpace(Mid(strLinha, 291, 2)))
        amiqn_ano = Val(RemoveSpace(Mid(strLinha, 297, 5)))
        amiqn_tipo = Val(Mid(strLinha, 306, 1))
        If RemoveSpace(Mid(strLinha, 311, 10)) <> "0" And RemoveSpace(Mid(strLinha, 311, 10)) <> "" Then
            amiqn_dt_incl = RevertDate(RemoveSpace(Mid(strLinha, 311, 10)))
        Else
            amiqn_dt_incl = "01/01/1900"
        End If
        
        Sql = "insert obvius_amiqn(amiqn_guia,amiqn_prest_id,amiqn_prest_tp,amiqn_prest_razao,amiqn_prest_fora,amiqn_tom_id,amiqn_tom_tp,amiqn_tom_razao,"
        Sql = Sql & "amiqn_tom_fora,amiqn_mes,amiqn_ano,amiqn_tipo,amiqn_dt_incl) values('" & amiqn_guia & "','" & amiqn_prest_id & "'," & amiqn_prest_tp & ",'"
        Sql = Sql & Mask(amiqn_prest_razao) & "'," & amiqn_prest_fora & ",'" & amiqn_tom_id & "'," & amiqn_tom_tp & ",'" & Mask(amiqn_tom_razao) & "'," & amiqn_tom_fora & ","
        Sql = Sql & amiqn_mes & "," & amiqn_ano & "," & amiqn_tipo & ",'" & Format(amiqn_dt_incl, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1


AMIST:
lblArq.Caption = "amist.txt"
lblArq.Refresh
FF1 = FreeFile()
Open sPathBackup & "amist.txt" For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    nRowCount = nRowCount + 1
Loop
Close #FF1
Pb.value = 0


Sql = "delete from obvius_amist"
cn.Execute Sql, rdExecDirect
FF1 = FreeFile()
nPos = 1
Open sPathBackup & "amist.txt" For Input As FF1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
  
        amist_im = Val(RemoveSpace(Mid(strLinha, 1, 12)))
        amist_mes = Val(RemoveSpace(Mid(strLinha, 16, 2)))
        amist_ano = Val(RemoveSpace(Mid(strLinha, 22, 5)))
        amist_st = Val(RemoveSpace(Mid(strLinha, 31, 1)))
        If Mid(strLinha, 35, 1) <> "?" Then
            amist_dt_incl = RevertDate(RemoveSpace(Mid(strLinha, 36, 10)))
        Else
            amist_dt_incl = "01/01/1900"
        End If
        
        Sql = "insert obvius_amist(amist_im,amist_mes,amist_ano,amist_st,amist_dt_incl) values(" & amist_im & "," & amist_mes & "," & amist_ano & ","
        Sql = Sql & amist_st & ",'" & Format(amist_dt_incl, "mm/dd/yyyy") & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
        DoEvents
   Loop
Close #FF1





fim:
lblArq.Caption = "FIM"
lblArq.Refresh

End Sub

Private Function RemoveSpace(sPalavra As String) As String
RemoveSpace = Replace(sPalavra, " ", "")
End Function

Private Function RevertDate(sDate As String) As String
RevertDate = Right(sDate, 2) & "/" & Mid(sDate, 5, 2) & "/" & Left(sDate, 4)
End Function

Private Sub cmdCanceladas_Click()
Dim nNumDoc As Long, strLinha As String, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nCod As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer, nStatus As Integer
Exit Sub
FF1 = FreeFile()
Open sPathBin & "\CANCELADAS.TXT" For Input As FF1
   Do While Not EOF(1)
        nCod = 0
        nAno = 0
        nLanc = 0
        nSeq = 0
        nParc = 0
        nCompl = 0
        Line Input #1, strLinha
        If IsNumeric(Mid(strLinha, 6, 7)) Then
            nNumDoc = CLng(Mid(strLinha, 6, 7))
            Sql = "SELECT * FROM PARCELADOCUMENTO WHERE NUMDOCUMENTO=" & nNumDoc
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    nCod = !CODREDUZIDO
                    nAno = !AnoExercicio
                    nLanc = !CodLancamento
                    nSeq = !SeqLancamento
                    nParc = !NumParcela
                    nCompl = !CODCOMPLEMENTO
                    Sql = "SELECT STATUSLANC FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCod & " AND ANOEXERCICIO=" & nAno & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        nStatus = !statuslanc
                        If nStatus = 3 Then
                            Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5 WHERE CODREDUZIDO=" & nCod & " AND ANOEXERCICIO=" & nAno & " AND "
                            Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                            cn.Execute Sql, rdExecDirect
                        End If
                       .Close
                    End With
                End If
               .Close
            End With
        End If
   Loop
On Error Resume Next
Close #FF1
MsgBox "fim"
End Sub

Private Sub cmdImportarSenha_Click()

Dim Sql As String, fName As String, cc As cCommonDlg, aName() As String, nFile As Integer
Dim nCodReduz As Long, sCNPJ As String, strLinha As String, sDataIni As String
Dim sDataFim As String, sDataOpcao As String, nAliquota As Integer, sUsuario As String, sDataAlt As String
Dim aRecord() As String, dDataIni As Date, nPos As Long, nTot As Long

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , False, , , "Documento de Texto|*.txt", , App.Path & "\Bin", "Selecione o arquivo do simples nacional", , Me.hwnd, OFN_HIDEREADONLY, False

If fName = "" Then Exit Sub
Ocupado

nTot = 1
nPos = 0
FF1 = FreeFile()
Open fName For Input As FF1
Do While Not EOF(1)
    Line Input #FF1, strLinha
    If Mid(strLinha, 7, 1) <> ";" And nTot = 1 Then
        Close #FF1
        Liberado
        MsgBox "Arquivo inválido.", vbCritical, "ERRO"
        Exit Sub
    End If
    nTot = nTot + 1
Loop
Close #FF1
Pb.value = 0


Open fName For Input As #1
Do While Not EOF(1)
    If nPos Mod 10 = 0 Then
        CallPb nPos, nTot
    End If
    Line Input #1, strLinha
    
    aRecord = Split(strLinha, ";")
    nCodReduz = CLng(aRecord(0))
    If Val(nCodReduz) < 100000 Or Val(nCodReduz) > 300000 Then GoTo proximo

    sCNPJ = aRecord(1)
    sDataIni = aRecord(2)
    sDataIni = Right(sDataIni, 2) & "/" & Mid(sDataIni, 5, 2) & "/" & Left(sDataIni, 4)
    
    sDataFim = aRecord(3)
    If sDataFim <> "0" Then
        sDataFim = Right(sDataFim, 2) & "/" & Mid(sDataFim, 5, 2) & "/" & Left(sDataFim, 4)
    End If
    
    sDataOpcao = aRecord(4)
    nAliq = Val(aRecord(5))
    sUsuario = aRecord(6)
    sDataOpcao = aRecord(7)
    
    Sql = "select * from simplesimporta where codreduz=" & nCodReduz & " and dataini='" & sDataIni & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        Sql = "INSERT SIMPLESIMPORTA(CODREDUZ,CNPJ,DATAINI,DATAFIM,DATAOPCAO,ALIQUOTA,USUARIO,DATAALT) VALUES("
        Sql = Sql & nCodReduz & ",'" & sCNPJ & "','" & sDataIni & "','" & sDataFim & "','" & sDataOpcao & "'," & nAliq & ",'"
        Sql = Sql & sUsuario & "','" & sDataOpcao & "')"
        cn.Execute Sql, rdExecDirect
        
        Sql = "select * from periodosn where codigo=" & nCodReduz & " order by dataini"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If RdoAux2.RowCount > 0 Then
                Do Until .EOF
                    dDataIni = !dataini
                    If IsNull(!datafim) Then
                        If DateDiff("d", dDataIni, CDate(sDataIni)) > 0 Then
                            Sql = "update periodosn set datafim='" & Format(CDate(sDataIni), "mm/dd/yyyy") & "' where codigo=" & nCodReduz & " and dataini='" & Format(!dataini, "mm/dd/yyyy") & "'"
                            cn.Execute Sql, rdExecDirect
                        End If
                    End If
                   .MoveNext
                Loop
                If DateDiff("d", dDataIni, CDate(sDataIni)) <> 0 Then
                    On Error Resume Next
                    Sql = "INSERT PERIODOSN(CODIGO,DATAINI,DATAFIM,USUARIO) VALUES(" & nCodReduz & ",'" & Format(CDate(sDataIni), "mm/dd/yyyy") & "',"
                    If sDataFim = "0" Then
                        Sql = Sql & "Null" & ",'Importado ISS')"
                    Else
                        Sql = Sql & "'" & Format(CDate(sDataFim), "mm/dd/yyyy") & "','Importado ISS')"
                    End If
                    cn.Execute Sql, rdExecDirect
                    On Error GoTo 0
                End If
            Else
                'não existem registros, então é só gravar direto
                Sql = "INSERT PERIODOSN(CODIGO,DATAINI,DATAFIM,USUARIO) VALUES(" & nCodReduz & ",'" & Format(CDate(sDataIni), "mm/dd/yyyy") & "',"
                If sDataFim = "0" Then
                    Sql = Sql & "Null" & ",'Importado ISS')"
                Else
                    Sql = Sql & "'" & Format(CDate(sDataFim), "mm/dd/yyyy") & "','Importado ISS')"
                End If
                cn.Execute Sql, rdExecDirect
            End If
           .Close
        End With
        
        RdoAux.Close
    Else
        RdoAux.Close
    End If
    nPos = nPos + 1
    DoEvents
proximo:
Loop
Close #1
Pb.value = 0
Liberado
MsgBox "Simples nacional importado!", vbInformation, "Informação"

End Sub

Private Sub cmdImportNF_Click()
Dim aNota() As NOTAS, nPos As Long, aRecord() As String, nRowCount As Long
Dim Sql As String, RdoAux As rdoResultset, nNumFiles As Integer, sFileName As String
Dim fName As String, cc As cCommonDlg, aName() As String, nFile As Integer
txtMsg.Text = ""
Set cc = New cCommonDlg
'cc.VBGetOpenFileName fName, , , True, , , "Documento de Texto|*.txt;*.prt", , App.Path & "\Bin", "Selecione um ou mais arquivos para importação", , Me.hwnd, OFN_HIDEREADONLY, False
cc.VBGetOpenFileName fName, , , , , , "Importação de NF|*.txt", , App.Path & "\Bin", "Selecione o arquivo de notas para importação", , Me.hwnd, OFN_HIDEREADONLY, False
aName = Split(fName, " ")

If UBound(aName) = -1 Then Exit Sub
nRowCount = 0
If UBound(aName) < 2 Then
    nNumFiles = 0
Else
    nNumFiles = 1
End If

For nFile = nNumFiles To UBound(aName)
    If nFile > 0 Then
        txtMsg.Text = "Importando arquivo: " & nFile & " de " & UBound(aName) & " - " & aName(nFile)
    Else
        txtMsg.Text = "Importando arquivo: 1 de 1 - " & aName(nFile)
    End If
    Me.Refresh
    Ocupado
    FF1 = FreeFile()
    ReDim aNota(0): nPos = 1
    If nNumFiles = 0 Then
        sFileName = aName(nFile)
    Else
        sFileName = aName(0) & aName(nFile)
    End If
    
    Open sFileName For Input As FF1
    Do While Not EOF(1)
        Line Input #FF1, strLinha
        nRowCount = nRowCount + 1
    Loop
    Close #FF1
    Pb.value = 0
    
    Open sFileName For Input As FF1
    Do While Not EOF(1)
        Line Input #FF1, strLinha
        If nPos Mod 30 = 0 Then
            CallPb nPos, nRowCount
        End If
        If UCase(Right(aName(nFile), 3)) = "TXT" Then
            aRecord = Split(strLinha, "#")
            
            If UBound(aRecord) < 26 Then
                MsgBox "O arquivo " & sFileName & " é inválido.", vbCritical, "Erro de Importação"
                Close #FF1
                GoTo NEXTFILE
            End If
            
            ReDim Preserve aNota(nPos)
            aNota(nPos).IdentificaPrestador = aRecord(0)
            aNota(nPos).TipoPrestador = Val(aRecord(1))
            aNota(nPos).TipoNota = Val(aRecord(2))
            aNota(nPos).NumeroNota = Val(aRecord(3))
            aNota(nPos).Serie = aRecord(4)
            aNota(nPos).DataEmissao = Right(aRecord(5), 2) & "/" & Mid(aRecord(5), 5, 2) & "/" & Left(aRecord(5), 4)
            aNota(nPos).MesRef = Val(aRecord(6))
            aNota(nPos).AnoRef = Val(aRecord(7))
            aNota(nPos).StatusNota = Val(aRecord(8))
            If aRecord(9) = "0" Then
                aNota(nPos).DataCancel = "Null"
            Else
                aNota(nPos).DataCancel = Right(aRecord(9), 2) & "/" & Mid(aRecord(9), 5, 2) & "/" & Left(aRecord(9), 4)
            End If
            aNota(nPos).Natureza = aRecord(10)
            aNota(nPos).ValorTotal = CDbl(aRecord(11)) / 100
            aNota(nPos).ValorServico = CDbl(aRecord(12)) / 100
            aNota(nPos).ValorImposto = CDbl(aRecord(13)) / 100
            aNota(nPos).Recolhimento = Val(aRecord(14))
            aNota(nPos).Atividade = Val(aRecord(15))
            aNota(nPos).Aliquota = CDbl(aRecord(16)) / 100
            aNota(nPos).RazaoPrestador = aRecord(17)
            aNota(nPos).CidadePrestador = aRecord(18)
            aNota(nPos).UFPrestador = aRecord(19)
            aNota(nPos).LocalPrestador = aRecord(20)
            aNota(nPos).IdentificaTomador = aRecord(21)
            aNota(nPos).TipoTomador = aRecord(22)
            aNota(nPos).RazaoTomador = aRecord(23)
            aNota(nPos).CidadeTomador = aRecord(24)
            aNota(nPos).UFTomador = aRecord(25)
            aNota(nPos).LocalTomador = aRecord(26)
            'If aNota(nPos).NumeroNota = 4788 Then MsgBox "teste"
            If Len(aRecord(33)) > 7 Then
                aNota(nPos).NumGuia = Right(aRecord(33), 7)
            Else
                aNota(nPos).NumGuia = 0
            End If
            
        Else
            If Left(strLinha, 5) = "AMINF" Or strLinha = "" Then
                GoTo PROXIMANF
            End If
            If Left(strLinha, 2) <> "IM" Then
                GoTo PROXIMANF
            End If
            ReDim Preserve aNota(nPos)
            aNota(nPos).IdentificaPrestador = Mid(strLinha, 18, 14)
            aNota(nPos).TipoPrestador = Val(Mid(strLinha, 33, 1))
            aNota(nPos).TipoNota = Val(Mid(strLinha, 71, 1))
            aNota(nPos).NumeroNota = Mid(strLinha, 44, 14)
            aNota(nPos).Serie = Mid(strLinha, 59, 1)
            aNota(nPos).DataEmissao = Mid(strLinha, 92, 2) & "/" & Mid(strLinha, 90, 2) & "/" & Mid(strLinha, 86, 4)
            aNota(nPos).MesRef = Mid(strLinha, 102, 2)
            aNota(nPos).AnoRef = Mid(strLinha, 110, 4)
            aNota(nPos).StatusNota = Mid(strLinha, 115, 1)
            If Trim(Mid(strLinha, 131, 8)) = "0" Then
                aNota(nPos).DataCancel = "Null"
            Else
                aNota(nPos).DataCancel = Mid(strLinha, 137, 2) & "/" & Mid(strLinha, 135, 2) & "/" & Mid(strLinha, 131, 4)
            End If
            aNota(nPos).Natureza = Mid(strLinha, 140, 1)
            aNota(nPos).ValorTotal = Val(Mid(strLinha, 151, 13))
            aNota(nPos).ValorServico = Val(Mid(strLinha, 165, 13))
            aNota(nPos).ValorImposto = Val(Mid(strLinha, 219, 13))
            aNota(nPos).Recolhimento = Val(Mid(strLinha, 233, 1))
            aNota(nPos).Atividade = Val(Mid(strLinha, 181, 8))
            aNota(nPos).Aliquota = Val(Mid(strLinha, 213, 6))
            aNota(nPos).RazaoPrestador = Mid(strLinha, 247, 100)
            aNota(nPos).CidadePrestador = Mid(strLinha, 348, 40)
            aNota(nPos).UFPrestador = Mid(strLinha, 389, 2)
            aNota(nPos).LocalPrestador = Mid(strLinha, 404, 1)
            aNota(nPos).IdentificaTomador = Mid(strLinha, 421, 14)
            aNota(nPos).TipoTomador = Val(Mid(strLinha, 436, 1))
            aNota(nPos).RazaoTomador = Mid(strLinha, 451, 100)
            aNota(nPos).CidadeTomador = Mid(strLinha, 552, 40)
            aNota(nPos).UFTomador = Mid(strLinha, 593, 2)
            aNota(nPos).LocalTomador = Mid(strLinha, 606, 1)
            aNota(nPos).NumGuia = Mid(strLinha, 626, 7)
        End If
PROXIMANF:
        nPos = nPos + 1
    Loop
    
    On Error Resume Next
    Close #FF1
    On Error GoTo 0
    
    For nPos = 0 To UBound(aNota)
        If nPos Mod 30 = 0 Then
            CallPb CLng(nPos), CLng(UBound(aNota))
        End If
        With aNota(nPos)
            If .IdentificaPrestador <> "" Then
                Sql = "SELECT identificaprestador FROM nfisseletro2 where identificaprestador='" & Trim(.IdentificaPrestador) & "' and numeronota='" & .NumeroNota & "' and serie='" & .Serie & "'"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount = 0 Then
                    Sql = "INSERT nfisseletro2(identificaprestador,tipoprestador,tiponota,numeronota,serie,dataemissao,mesref,anoref,statusnota,datacancel,natureza,valortotal,"
                    Sql = Sql & "valorservico,valorimposto,recolhimento,atividade,aliquota,razaoprestador,cidadeprestador,ufprestador,localprestador,identificatomador,"
                    Sql = Sql & "tipotomador,razaotomador,cidadetomador,uftomador,LocalTomador,numdoc) values('" & Trim(.IdentificaPrestador) & "'," & .TipoPrestador & "," & .TipoNota & ","
                    Sql = Sql & .NumeroNota & ",'" & .Serie & "','" & Format(.DataEmissao, "mm/dd/yyyy") & "'," & .MesRef & "," & .AnoRef & "," & .StatusNota & "," & IIf(.DataCancel = "Null" Or .DataCancel = "00/00/0000", "Null", "'" & Format(.DataCancel, "mm/dd/yyyy") & "'") & ",'" & .Natureza & "',"
                    Sql = Sql & Virg2Ponto(CStr(.ValorTotal)) & "," & Virg2Ponto(CStr(.ValorServico)) & "," & Virg2Ponto(CStr(.ValorImposto)) & "," & .Recolhimento & "," & .Atividade & "," & Virg2Ponto(CStr(.Aliquota)) & ",'" & Mask(Trim(.RazaoPrestador)) & "','"
                    Sql = Sql & Mask(Trim(.CidadePrestador)) & "','" & .UFPrestador & "','" & .LocalPrestador & "','" & Trim(.IdentificaTomador) & "','" & .TipoTomador & "','" & Mask(Trim(.RazaoTomador)) & "','"
                    Sql = Sql & Mask(Trim(.CidadeTomador)) & "','" & .UFTomador & "','" & .LocalTomador & "'," & .NumGuia & ")"
                Else
                    If .NumGuia > 0 Then
                        Sql = "update nfisseletro2 set numdoc=" & .NumGuia & " where identificaprestador='" & Trim(.IdentificaPrestador) & "' and numeronota='" & .NumeroNota & "' and serie='" & .Serie & "'"
                    End If
                    
                End If
                cn.Execute Sql, rdExecDirect
                RdoAux.Close
            End If
        End With
    Next
NEXTFILE:
Next 'next file
Pb.value = 0
Liberado
txtMsg.Text = ""
Me.Refresh
MsgBox "Fim da importação dos arquivos.", vbInformation, "Informação"

End Sub

Private Sub cmdAdd_Click()
Dim z As Variant, x As Integer

z = InputBox("Digite o código da empresa.", "Atenção")
If Val(z) = 0 Then Exit Sub

If Val(z) < 500000 Or Val(z) >= 600000 Then
    MsgBox "Código fora da faixa.", vbCritical, "Atenção"
    Exit Sub
End If

For x = 0 To lstISS.ListCount - 1
    If Val(z) = Val(lstISS.List(x)) Then
        MsgBox "Código já existe na lista.", vbCritical, "Atenção"
        Exit Sub
    End If
Next

Sql = "SELECT CODCIDADAO FROM CIDADAO WHERE CODCIDADAO=" & Val(z)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Código não cadastrado.", vbCritical, "Atenção"
        Exit Sub
    End If
   .Close
End With

Sql = "INSERT ISSRETIDO (CODREDUZIDO) VALUES(" & Val(z) & ")"
cn.Execute Sql, rdExecDirect
CarregaLista
MsgBox "Código Inserido.", vbInformation, "Informação"

End Sub

Private Sub cmdCancel_Click()

If MsgBox("Deseja cancelar a operação.", vbQuestion + vbYesNo + vbDefaultButton2, "Cancelamento !!!") = vbYes Then
    bCancel = True
End If

End Sub

Private Sub cmdDel_Click()
If lstISS.ListIndex > -1 Then
    If MsgBox("Excluir a empresa " & lstISS.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "DELETE FROM ISSRETIDO WHERE CODREDUZIDO=" & lstISS.Text
        cn.Execute Sql, rdExecDirect
        CarregaLista
    End If
End If
End Sub

Private Sub cmdExport_Click()
bConsist = True
ExportaConsist
End Sub

Private Sub cmdImport_Click()
Dim fName As String, cc As cCommonDlg, aName() As String, nFile As Integer

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Importação de guias|*.txt", , App.Path & "\Bin", "Selecione o arquivo de guias para importação", , Me.hwnd, OFN_HIDEREADONLY, False
aName = Split(fName, ";")

'If Dir(sPathBin & "\GUIAS.TXT") = "" Then
'    MsgBox "O Arquivo 'GUIAS.TXT' não foi localizado, certifique-se de ter copiado " & vbCrLf & "para o diretório 'BIN' do diretório raiz do sistema.", vbCritical, "Atenção"
 ''   Exit Sub
'End If
If UBound(aName) = 0 Then
    FF1 = FreeFile()
    Open aName(0) For Input As FF1
       Do While Not EOF(1)
            Line Input #1, strLinha
            If Left(strLinha, 1) = 2 Then
                Close #FF1
                ImportaConsistNovo (aName(0))
                Exit Do
            ElseIf Left(strLinha, 1) = 0 Then
                Close #FF1
                ImportaConsist2
                Exit Do
            End If
       Loop
    On Error Resume Next
    Close #FF1
Else
    MsgBox "Operação cancelada ou arquivo inválido.", vbCritical, "Atenção"
End If

End Sub

Private Sub cmdImportSMov_Click()
Dim aNota() As ISSELETRO, aRecord() As String
Dim Sql As String, RdoAux As rdoResultset, x As Integer
Dim fName As String, cc As cCommonDlg, aName() As String, nFile As Integer

If Dir(sPathBin & "\SEMMOV.CSV") = "" Then
    MsgBox "O Arquivo 'SEMMOV.CSV' não foi localizado, certifique-se de ter copiado " & vbCrLf & "para o diretório 'BIN' do diretório raiz do sistema.", vbCritical, "Atenção"
    Exit Sub
End If

txtMsg.Text = ""

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Documento sem Movimento|SEMMOV.CSV", , App.Path & "\Bin", "Selecione o arquivo sem movimento para importação", , Me.hwnd, OFN_HIDEREADONLY, False
aName = Split(fName, ";")
Ocupado
If UBound(aName) = 0 Then
    txtMsg.Text = "Importando arquivo sem movimento"
    FF1 = FreeFile()
    ReDim aNota(0): nPos = 1
    Open aName(0) For Input As FF1
       Do While Not EOF(1)
            Line Input #1, strLinha
            aRecord = Split(strLinha, ";")
            ReDim Preserve aNota(nPos)
            aNota(nPos).Identificacao = aRecord(0)
            aNota(nPos).Ano = aRecord(3)
            Select Case UCase(aRecord(2))
                Case "JANEIRO"
                    aNota(nPos).Mes = 1
                Case "FEVEREIRO"
                    aNota(nPos).Mes = 2
                Case "MARÇO", "MARCO"
                    aNota(nPos).Mes = 3
                Case "ABRIL"
                    aNota(nPos).Mes = 4
                Case "MAIO"
                    aNota(nPos).Mes = 5
                Case "JUNHO"
                    aNota(nPos).Mes = 6
                Case "JULHO"
                    aNota(nPos).Mes = 7
                Case "AGOSTO"
                    aNota(nPos).Mes = 8
                Case "SETEMBRO"
                    aNota(nPos).Mes = 9
                Case "OUTUBRO"
                    aNota(nPos).Mes = 10
                Case "NOVEMBRO"
                    aNota(nPos).Mes = 11
                Case "DEZEMBRO"
                    aNota(nPos).Mes = 12
            End Select
            nPos = nPos + 1
       Loop
    On Error Resume Next
    Close #FF1
    On Error GoTo 0
    
    For x = 1 To UBound(aNota)
        CallPb CLng(x), CLng(UBound(aNota))

        With aNota(x)
            Sql = "SELECT * FROM NFISSELETROSMOV WHERE CODIGO=" & .Identificacao & "  AND ANO=" & .Ano & " AND MES=" & .Mes
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
            If RdoAux.RowCount = 0 Then
                Sql = "INSERT NFISSELETROSMOV(CODIGO,ANO,MES,SMOV) VALUES(" & .Identificacao & "," & .Ano & "," & .Mes & ",1)"
                cn.Execute Sql, rdExecDirect
            End If
            RdoAux.Close
        End With
    Next
    txtMsg.Text = ""
    Liberado
    MsgBox "Arquivo importado com sucesso.", vbInformation, "Atenção"
Else
    MsgBox "Operação cancelada ou arquivo inválido.", vbCritical, "Atenção"
End If
Liberado
End Sub

Private Sub cmdOK_Click()
Dim nExporta As Integer

If Opt(0).value = True Then
    bCancel = False
    For nExporta = 0 To lstOpc.ListCount - 1
        If lstOpc.Selected(nExporta) = True Then
            lblTipo.Caption = Left(lstOpc.List(nExporta), 2)
            Exporta Val(lblTipo.Caption)
        End If
    Next
Else
    If lstOpc.ListIndex = 4 Then
        'ImportaConsistNovo
    Else
        Importa
    End If
End If
Pb.Color = vbWhite
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim Sql As String, RdoAux As rdoResultset, amigu_id As String, amigu_guia As String, amigu_tipo As Integer, amigu_seq As Integer
Dim amigu_codreduzido As Long, amigu_numdoc As Long, RdoAux2 As rdoResultset, aminf_tp As Integer, aminf_tipo As Integer, aminf_nb As String, aminf_serie As Integer
Dim nCodReduz As Single

Sql = "select amigu_num,amigu_seq,amigu_id,amigu_tipo from obvius_amigu  order by amigu_num"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        amigu_codreduzido = 0
        amigu_tipo = !amigu_tipo
        amigu_id = !amigu_id
        amigu_seq = !amigu_seq
                
        If amigu_tipo = 2 Then
            If CLng(amigu_id) < 600000 Then
                amigu_codreduzido = amigu_id
            Else
                GoTo proximo
            End If
        Else
            If Len(amigu_id) = 11 Then
                Sql = "select codcidadao,cpf from cidadao where cpf='" & amigu_id & "'"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                If RdoAux2.RowCount > 0 Then
                    amigu_id = RdoAux2!CodCidadao
                    RdoAux2.Close
                Else
                    'MsgBox "teste"
                End If
            Else
                Sql = "select codcidadao,cnpj from cidadao where cnpj='" & amigu_id & "'"
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                If RdoAux2.RowCount > 0 Then
                    amigu_id = RdoAux2!CodCidadao
                    RdoAux2.Close
                Else
                    RdoAux2.Close
                    Sql = "select codigomob,cnpj from mobiliario where cnpj='" & amigu_id & "'"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
                    If RdoAux2.RowCount > 0 Then
                        amigu_id = RdoAux2!codigomob
                    Else
'                        MsgBox "teste|"
                    End If
                    
                End If
            End If
        End If
        
        amigu_guia = !amigu_num
        amigu_numdoc = CLng(Right(amigu_guia, 7))
        
        Sql = "update obvius_amigu set amigu_codreduzido=" & amigu_codreduzido & ",amigu_numdoc=" & amigu_numdoc & " where amigu_num='" & amigu_guia & "' and "
        Sql = Sql & "amigu_seq=" & amigu_seq
        cn.Execute Sql, rdExecDirect
proximo:
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Sub Command2_Click()
Dim Sql As String, RdoAux As rdoResultset, aminf_usr As String, aminf_usrid As Long, aminf_id As String, aminf_guia As String, aminf_tom_id As String
Dim aminf_codreduzido As Long, aminf_numdoc As Long, RdoAux2 As rdoResultset, aminf_tp As Integer, aminf_tipo As Integer, aminf_nb As String, aminf_serie As Integer
Dim nCodReduz As Single

Sql = "select aminf_usr,aminf_id,aminf_tipo,aminf_nb,aminf_serie, aminf_guia,aminf_tp,aminf_tom_id from obvius_aminf order by aminf_id,aminf_guia"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
    On Error Resume Next
        aminf_usr = !aminf_usr
        aminf_usrid = CLng(Right(aminf_usr, 6))
        aminf_tipo = !aminf_tipo
        aminf_nb = !aminf_nb
        aminf_serie = !aminf_serie
        aminf_tp = !aminf_tp
        aminf_id = !aminf_id
        aminf_tom_id = !aminf_tom_id
                
        If aminf_usrid = aminf_id Then
            nCodReduz = aminf_id
            aminf_codreduzido = Val(aminf_id)
        Else
            nCodReduz = aminf_tom_id
            aminf_codreduzido = Val(aminf_tom_id)
        End If
        
     '   aminf_codreduzido = 0
'        If nCodReduz > 100000 And nCodReduz < 600000 Then
'           ' aminf_codreduzido = Val(aminf_id)
'        Else
'            If Len(nCodReduz) = 11 Then
'                Sql = "select codigomob,cpf from mobiliario where cpf='" & aminf_id & "'"
'                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
'                If RdoAux2.RowCount > 0 Then
'                    aminf_codreduzido = RdoAux2!CODIGOMOB
'                Else
'                    RdoAux2.Close
'                    Sql = "select codcidadao,cpf from cidadao where cpf='" & aminf_id & "'"
'                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
'                    If RdoAux2.RowCount > 0 Then
'                        aminf_codreduzido = RdoAux2!CodCidadao
'                    End If
'                End If
'            Else
'                Sql = "a!"
'            End If
'        End If
        
        
        aminf_guia = !aminf_guia
        If aminf_guia <> "0" Then
            aminf_numdoc = CLng(Right(aminf_guia, 7))
        Else
            aminf_numdoc = 0
        End If
        
        Sql = "update obvius_aminf set aminf_codreduzido=" & aminf_codreduzido & ",aminf_numdoc=" & aminf_numdoc & " where aminf_usr='" & aminf_usr & "' and "
        Sql = Sql & "aminf_id='" & aminf_id & "' and aminf_tipo=" & aminf_tipo & " and aminf_nb='" & !aminf_nb & "' and aminf_serie=" & !aminf_serie & " and "
        Sql = Sql & "aminf_tp=" & !aminf_tp
        cn.Execute Sql, rdExecDirect
        DoEvents
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"

End Sub

Private Sub Command3_Click()
'importar as NF



End Sub

Private Sub Form_Load()

Centraliza Me
CarregaLista
End Sub

Private Sub CarregaLista()
lstISS.Clear
Sql = "SELECT CODREDUZIDO FROM ISSRETIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstISS.AddItem !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Exporta(nIndex As Integer)
Dim nPos As Long, nTot As Long, sDoc As String, sDataNascto As String, nSexo As Integer, z As Integer, bConsist As Boolean
Dim sLayout As String, nCodCid As Long, sNomeProp As String, sCidadeEntrega As String, sUFEntrega As String, sCepEntrega As String
Dim sEndereco As String, nNumero As Integer, sComplemento As String, nTipoEnd As Integer
Dim sBairro As String, sCidade As String, sUF As String, sCep As String, sEnderecoEntrega As String, nNumEntrega As Integer, sBairroEntrega As String
Dim sCodProp1 As String, sCodProp2 As String, sNomeProp1 As String, sNomeProp2 As String, sCPF1 As String, sCPF2 As String, sEND1 As String, sEND2 As String
Dim sCIDADE1 As String, sCIDADE2 As String, sUF1 As String, sUF2 As String, sTipoTrib As String, sTipoInc As String

Ocupado
'INICIALIZA
nPos = 0: bConsist = False
Pb.value = 0
'VERIFICA OPÇÃO
Select Case nIndex
    Case 0 'CADASTRO DE DÉBITOS
        Open sPathBin & "\DEBITO.TXT" For Output As #1
            sLayout = "00"
            Sql = "SELECT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, "
            Sql = Sql & "debitoparcela.codcomplemento, debitoparcela.statuslanc, debitoparcela.datavencimento, debitoparcela.datadebase, debitoparcela.numerolivro,"
            Sql = Sql & "debitoparcela.paginalivro, debitoparcela.numcertidao, debitoparcela.datainscricao, debitoparcela.dataajuiza, debitotributo.codtributo,"
            Sql = Sql & "debitotributo.valortributo , TRIBUTO.desctributo, TRIBUTO.abrevtributo FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND "
            Sql = Sql & "debitoparcela.anoexercicio = debitotributo.anoexercicio AND debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
            Sql = Sql & "debitoparcela.numparcela = debitotributo.numparcela AND debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN tributo ON debitotributo.codtributo = tributo.codtributo "
            Sql = Sql & "Where (debitoparcela.AnoExercicio BETWEEN 2002 AND 2009) And (debitoparcela.statuslanc = 3) And (debitoparcela.CodLancamento <> 20) And (debitoparcela.NumParcela > 0) AND DEBITOTRIBUTO.CODTRIBUTO<>3 AND DATAINSCRICAO IS NOT NULL AND DATAAJUIZA IS NULL"
            'Sql = Sql & " AND DEBITOPARCELA.CODREDUZIDO<100000" 'SOMENTE IPTU
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                nTot = .RowCount: lblRegTot.Caption = nTot
               'PREPARA MATRIZ DE DADOS PARA CÁLCULO
'                LoadMatrix
                Do Until .EOF
                    nPos = nPos + 1
                    If nPos Mod 50 = 0 Then
                       CallPb nPos, CLng(nTot)
                       lblRegPerc.Caption = nPos
                    End If
                   'CARREGA AS VARIÁVEIS
                    sTipoTrib = 0
                    Select Case !CODREDUZIDO
                        Case Is < 100000
                            sTipoInc = 1
                        Case Is >= 100000 And !CODREDUZIDO < 500000
                            sTipoInc = 2
                        Case Is >= 500000
                            sTipoInc = 3
                    End Select
                    ax = sLayout & ";" & sTipoTrib & ";" & FillLeft(!CodTributo, 6) & ";" & FillSpace(!desctributo, 50) & ";" & sTipoInc & ";" & FillLeft(!CODREDUZIDO, 6) & ";" & FillLeft(!CODREDUZIDO, 6) & ";"
                    ax = ax & !AnoExercicio & ";" & Format(!datainscricao, "dd/mm/yyyy") & ";" & FillLeft(Virg2Ponto(Format(!ValorTributo, "0000000.00")), 10) & ";" & FillSpace(Val(SubNull(!numerolivro)), 3) & ";"
                    ax = ax & FillLeft(Val(Left$(SubNull(!paginalivro), 3)), 3) & ";" & FillLeft(!NumParcela, 3) & ";" & Format(!datainscricao, "dd/mm/yyyy") & ";" & Format(!DataVencimento, "dd/mm/yyyy") & ";"
                    ax = ax & Format(!DataVencimento, "dd/mm/yyyy") & ";" & Format("0", "0000000.00") & ";" & FillLeft(Val(Left$(SubNull(!numcertidao), 6)), 6) & ";" & Year(!datainscricao) & ";"
                    ax = ax & Format(!datainscricao, "dd/mm/yyyy") & ";" & IIf(IsNull(!dataajuiza), 0, 1) & ";" & FillLeft(0, 6) & ";" & FillLeft(0, 4) & ";" & "00/00/0000"
                    Print #1, ax
                   .MoveNext
                Loop
               .Close
            End With
        Close #1
    Case 1 'CADASTRO IMOBILIÁRIO
        Open sPathBin & "\IMOBILIARIO.TXT" For Output As #1
            sLayout = "01"
            'CARREGA O CÓDIGO DE TODOS OS IMÓVEIS QUE POSSUAM DEBITOS ATRASADOS
            Sql = "SELECT debitoparcela.codreduzido, vwCnsImovel.distrito, vwCnsImovel.setor, vwCnsImovel.quadra, vwCnsImovel.lote, vwCnsImovel.seq,"
            Sql = Sql & "vwCnsImovel.unidade, vwCnsImovel.subunidade,vwCnsImovel.ee_tipoend ,vwCnsImovel.codlogr, vwCnsImovel.abrevtipolog, vwCnsImovel.abrevtitlog,vwCnsImovel.codtipolog,vwCnsImovel.codtitlog,"
            Sql = Sql & "vwCnsImovel.NomeLogradouro , vwCnsImovel.Li_Num, vwCnsImovel.Li_Compl, vwCnsImovel.Dt_AreaTerreno FROM debitoparcela INNER JOIN "
            Sql = Sql & "vwCnsImovel ON debitoparcela.codreduzido = vwCnsImovel.codreduzido Where (debitoparcela.statuslanc = 3) AND DATAINSCRICAO IS NOT NULL AND DATAAJUIZA IS NULL And (debitoparcela.CodLancamento <> 20) And (debitoparcela.NumParcela > 0) And (debitoparcela.AnoExercicio >= 2002) "
 '           Sql = Sql & " AND DEBITOPARCELA.CODREDUZIDO<100000" 'SOMENTE IPTU
            Sql = Sql & "GROUP BY debitoparcela.codreduzido, vwCnsImovel.abrevtipolog,vwCnsImovel.codtipolog,vwCnsImovel.codtitlog, vwCnsImovel.abrevtitlog, vwCnsImovel.nomelogradouro, vwCnsImovel.distrito,"
            Sql = Sql & "vwCnsImovel.setor, vwCnsImovel.quadra, vwCnsImovel.lote, vwCnsImovel.seq, vwCnsImovel.unidade, vwCnsImovel.subunidade,vwCnsImovel.ee_tipoend,vwCnsImovel.codlogr, vwCnsImovel.abrevtipolog, vwCnsImovel.abrevtitlog, vwCnsImovel.nomelogradouro, vwCnsImovel.li_num,"
            Sql = Sql & "vwCnsImovel.Li_Compl , vwCnsImovel.Dt_AreaTerreno Having (debitoparcela.CODREDUZIDO < 100000) ORDER BY debitoparcela.codreduzido"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                nTot = .RowCount: lblRegTot.Caption = nTot
                'PREPARA MATRIZ DE DADOS PARA CÁLCULO
                LoadMatrix
                Do Until .EOF
                    'CALCULA OS VALORES VENAIS
                    CalculoIndividual !CODREDUZIDO
                    nPos = nPos + 1
                    If nPos Mod 50 = 0 Then
                       CallPb nPos, CLng(nTot)
                       lblRegPerc.Caption = nPos
                    End If
                    'PROPRIETARIO
                    Sql = "SELECT * FROM VWCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & !CODREDUZIDO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        If .RowCount > 0 Then
                            nCodCid = !CodCidadao
                            sNomeProp = !nomecidadao
                        Else
                            nCodCid = 0
                            sNomeProp = ""
                        End If
                    End With
                    'ENDEREÇO DE ENTREGA
                    nTipoEnd = !Ee_TipoEnd
                    If nTipoEnd = 0 Then 'ENDEREÇO DO IMOVEL
                        sEnderecoEntrega = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " nº " & !Li_Num
                        sCidadeEntrega = "JABOTICABAL"
                        sUFEntrega = "SP"
                        sCepEntrega = RetornaCEP(!CodLogr, !Li_Num)
                    ElseIf nTipoEnd = 1 Then 'ENDEREÇO DO PROPRIETARIO
                        If nCodCid > 0 Then
                            Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG, "
                            Sql = Sql & "vwLOGRADOURO.NOMELOGRADOURO, cidadao.numimovel, cidade.desccidade, cidadao.nomelogradouro AS nomelogradouro2,"
                            Sql = Sql & "Cidadao.SiglaUF FROM cidadao INNER JOIN cidade ON cidadao.siglauf = cidade.siglauf AND cidadao.codcidade = cidade.codcidade LEFT OUTER JOIN "
                            Sql = Sql & "vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO where CODCIDADAO=" & nCodCid
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                If Val(SubNull(!CodLogradouro)) > 0 Then
                                    sEnderecoEntrega = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " nº " & !NUMIMOVEL
                                Else
                                    sEnderecoEntrega = !NOMELOGRADOURO2 & " nº " & !NUMIMOVEL
                                End If
                                sCidadeEntrega = !descCidade
                                sUFEntrega = !SiglaUF
                               .Close
                            End With
                        Else
                            sEnderecoEntrega = ""
                            sCidadeEntrega = ""
                            sUFEntrega = ""
                        End If
                    ElseIf nTipoEnd = 2 Then 'ENDEREÇO DE ENTREGA
                        Sql = "SELECT endentrega.codreduzido, endentrega.ee_codlog, endentrega.ee_nomelog, endentrega.ee_numimovel, endentrega.ee_complemento, endentrega.ee_uf, "
                        Sql = Sql & "endentrega.ee_cidade, endentrega.ee_bairro, endentrega.ee_cep, endentrega.ee_loteamento, endentrega.ee_descbairro, cidade.desccidade,"
                        Sql = Sql & "vwLOGRADOURO.AbrevTipoLog,vwLOGRADOURO.AbrevTitLog , vwLOGRADOURO.NOMETITLOG, vwLOGRADOURO.NomeLogradouro FROM endentrega INNER JOIN cidade ON endentrega.ee_uf = cidade.siglauf AND "
                        Sql = Sql & "endentrega.ee_cidade = cidade.codcidade LEFT OUTER JOIN vwLOGRADOURO ON endentrega.ee_codlog = vwLOGRADOURO.CODLOGRADOURO WHERE CODREDUZIDO = " & RdoAux!CODREDUZIDO
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            If .RowCount > 0 Then
                                If !Ee_CodLog = 0 Then
                                    sEnderecoEntrega = !Ee_NomeLog & " nº " & !Ee_NumImovel
                                Else
                                    sEnderecoEntrega = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " nº " & !Ee_NumImovel
                                End If
                                sCidadeEntrega = !descCidade
                                sUFEntrega = !Ee_Uf
                                If Not IsNull(!Ee_Cep) Then
                                    sCepEntrega = !Ee_Cep
                                Else
                                    sCepEntrega = RetornaCEP(!Ee_CodLog, !Ee_NumImovel)
                                End If
                            Else
                            End If
                           .Close
                        End With
                    End If
                    'GRAVA NO ARQUIVO
                    ax = sLayout & ";" & FillLeft(!CODREDUZIDO, 6) & ";" & !Distrito & ";" & FillSpace(!Setor, 5) & ";" & FillSpace(!Quadra, 4) & ";"
                    ax = ax & FillSpace(!Lote, 4) & ";" & FillSpace(!Seq, 2) & ";" & FillLeft(!Unidade, 3) & ";" & FillLeft(!SubUnidade, 3) & ";" & FillLeft(!CodLogr, 6) & ";"
                    ax = ax & FillSpace(!CODTIPOLOG, 15) & ";" & FillSpace(Val(SubNull(IIf(!CODTITLOG = 9999, 0, !CODTITLOG))), 15) & ";" & FillSpace(!NomeLogradouro, 50) & ";"
                    ax = ax & FillLeft(!Li_Num, 4) & ";" & FillSpace(!Li_Compl, 15) & ";" & FillSpace(RetornaNumero(RetornaCEP(!CodLogr, !Li_Num)), 8) & ";"
                    ax = ax & FillLeft(Virg2Ponto(Format(!Dt_AreaTerreno, "0000000.00")), 10) & ";" & FillLeft(Virg2Ponto(Format(nAreaTotal, "0000000.00")), 10) & ";"
                    ax = ax & FillLeft(Virg2Ponto(Format(nVVT, "0000000.00")), 10) & ";" & FillLeft(Virg2Ponto(Format(nVVP, "0000000.00")), 10) & ";"
                    ax = ax & FillSpace(sNomeProp, 70) & ";" & FillSpace(sEnderecoEntrega, 100) & ";" & FillSpace(sCidadeEntrega, 50) & ";"
                    ax = ax & FillSpace(sUFEntrega, 2) & ";" & FillSpace(RetornaNumero(sCepEntrega), 8)
                    Print #1, ax
                   .MoveNext
                Loop
               .Close
            End With
        Close #1
    Case 2 'CADASTRO MOBILIÁRIO
        Open sPathBin & "\MOBILIARIO.TXT" For Output As #1
            sLayout = "02"
            'CARREGA O CÓDIGO DE TODOS AS EMPRESAS QUE POSSUAM DEBITOS ATRASADOS
            Sql = "SELECT vwCNSMOBILIARIO.codigomob, vwCNSMOBILIARIO.razaosocial, vwCNSMOBILIARIO.nomefantasia, vwCNSMOBILIARIO.codlogradouro, "
            Sql = Sql & "vwCNSMOBILIARIO.CODTIPOLOG,vwCNSMOBILIARIO.CODTITLOG,vwCNSMOBILIARIO.ABREVTIPOLOG, vwCNSMOBILIARIO.ABREVTITLOG, vwCNSMOBILIARIO.NOMELOGRADOURO, vwCNSMOBILIARIO.numero,"
            Sql = Sql & "vwCNSMOBILIARIO.complemento, vwCNSMOBILIARIO.codcidade, vwCNSMOBILIARIO.siglauf, vwCNSMOBILIARIO.cep,vwCNSMOBILIARIO.dataabertura,"
            Sql = Sql & "vwCNSMOBILIARIO.dataencerramento, vwCNSMOBILIARIO.CODATIVIDADE,vwCNSMOBILIARIO.cnpj, vwCNSMOBILIARIO.cpf,vwCNSMOBILIARIO.ativextenso, vwCNSMOBILIARIO.descuf, "
            Sql = Sql & "vwCNSMOBILIARIO.desccidade, vwCNSMOBILIARIO.EECODLOGR,vwCNSMOBILIARIO.EENOMELOGR, vwCNSMOBILIARIO.EENUMERO, vwCNSMOBILIARIO.EECOMPL,"
            Sql = Sql & "vwCNSMOBILIARIO.EEUF,vwCNSMOBILIARIO.EECODCIDADE, vwCNSMOBILIARIO.EECEP, vwCNSMOBILIARIO.EEDESCCIDADE, vwCNSMOBILIARIO.NOMELOGR,"
            Sql = Sql & "cidade.desccidade AS desccidade FROM vwCNSMOBILIARIO INNER JOIN debitoparcela ON vwCNSMOBILIARIO.codigomob = debitoparcela.codreduzido INNER JOIN "
            Sql = Sql & "cidade ON vwCNSMOBILIARIO.siglauf = cidade.siglauf AND vwCNSMOBILIARIO.codcidade = cidade.codcidade Where (debitoparcela.AnoExercicio >= 2002) And "
            Sql = Sql & "(debitoparcela.CodLancamento <> 20) And (debitoparcela.NumParcela > 0) And (debitoparcela.statuslanc = 3) AND DATAINSCRICAO IS NOT NULL AND DATAAJUIZA IS NULL GROUP BY vwCNSMOBILIARIO.codigomob, "
            Sql = Sql & "vwCNSMOBILIARIO.razaosocial, vwCNSMOBILIARIO.nomefantasia, vwCNSMOBILIARIO.codlogradouro,vwCNSMOBILIARIO.CODTIPOLOG,vwCNSMOBILIARIO.CODTITLOG,vwCNSMOBILIARIO.ABREVTIPOLOG, vwCNSMOBILIARIO.ABREVTITLOG, "
            Sql = Sql & "vwCNSMOBILIARIO.NOMELOGRADOURO, vwCNSMOBILIARIO.numero,vwCNSMOBILIARIO.complemento, vwCNSMOBILIARIO.codcidade, vwCNSMOBILIARIO.siglauf, vwCNSMOBILIARIO.cep,"
            Sql = Sql & "vwCNSMOBILIARIO.dataabertura, vwCNSMOBILIARIO.dataencerramento, vwCNSMOBILIARIO.CODATIVIDADE, vwCNSMOBILIARIO.cnpj, vwCNSMOBILIARIO.cpf,vwCNSMOBILIARIO.ativextenso, vwCNSMOBILIARIO.descuf, "
            Sql = Sql & "vwCNSMOBILIARIO.desccidade, vwCNSMOBILIARIO.EECODLOGR,vwCNSMOBILIARIO.EENOMELOGR, vwCNSMOBILIARIO.EENUMERO, vwCNSMOBILIARIO.EECOMPL, vwCNSMOBILIARIO.EEUF,"
            Sql = Sql & "vwCNSMOBILIARIO.EECODCIDADE, vwCNSMOBILIARIO.EECEP, vwCNSMOBILIARIO.EEDESCCIDADE, vwCNSMOBILIARIO.NOMELOGR,Cidade.desccidade"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                nTot = .RowCount: lblRegTot.Caption = nTot
                Do Until .EOF
                    nPos = nPos + 1
                    If nPos Mod 50 = 0 Then
                       CallPb nPos, CLng(nTot)
                       lblRegPerc.Caption = nPos
                    End If
                    If !codigomob > 190000 And !codigomob < 200000 Then GoTo PROXIMOMOB
'                    If !CODIGOMOB = 107532 Then MsgBox "TESTE"
                    If Not IsNull(!CPF) And Val(SubNull(!CPF)) > 0 Then
                        sDoc = Format(Val(RetornaNumero(!CPF)), "000\.000\.000-00")
                    ElseIf (Not IsNull(!Cnpj)) And (Val(SubNull(!Cnpj)) > 0) Then
                        sDoc = Format(Val(RetornaNumero(!Cnpj)), "000\.000\.000/0000-00")
                    Else
                        sDoc = 0
                    End If
                    sDoc = IIf(Val(sDoc) = 0, FillSpace(" ", 19), FillLeft(sDoc, 19))
                    If Not IsNull(!descCidade) Then
                        sCidade = !descCidade
                    Else
                        sCidade = !desccidade2
                    End If
                    z = 1
                    'BUSCA OS PROPRIETÁRIOS
                    Sql = "SELECT CODCIDADAO FROM MOBILIARIOPROPRIETARIO WHERE CODMOBILIARIO=" & !codigomob
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        sCodProp1 = "": sCodProp2 = "": sNomeProp1 = "": sNomeProp2 = "": sCPF1 = "": sCPF2 = ""
                        sEND1 = "": sEND2 = "": sCIDADE1 = "": sCIDADE2 = "": sUF1 = "": sUF2 = ""
                        Do Until .EOF
                            Sql = "SELECT vwCIDADAO.*, cidade.desccidade AS NOMECIDADE2 FROM vwCIDADAO LEFT OUTER JOIN "
                            Sql = Sql & "cidade ON vwCIDADAO.siglauf = cidade.siglauf AND vwCIDADAO.codcidade = cidade.codcidade "
                            Sql = Sql & "Where CodCidadao = " & !CodCidadao
                            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux3
                                If .RowCount > 0 Then
                                    If z = 1 Then
                                        sCodProp1 = CStr(!CodCidadao)
                                        sNomeProp1 = !nomecidadao
                                        If Not IsNull(!CPF) Then
                                            sCPF1 = Format(Val(RetornaNumero(!CPF)), "000\.000\.000-00")
                                        End If
                                        If Not IsNull(!NomeLogradouro) Then
                                            sEND1 = !NomeLogradouro & " Nº " & !NUMIMOVEL
                                        Else
                                            If Not IsNull(!NOMELOGRADOURO2) Then
                                                sEND1 = !NOMELOGRADOURO2 & " Nº " & !NUMIMOVEL
                                            Else
                                                sEND1 = ""
                                            End If
                                        End If
                                        sCIDADE1 = SubNull(!NOMECIDADE2)
                                        sUF1 = SubNull(!SiglaUF)
                                    ElseIf z = 2 Then
                                        sCodProp2 = CStr(!CodCidadao)
                                        sNomeProp2 = !nomecidadao
                                        If Not IsNull(!CPF) Then
                                            sCPF2 = Format(Val(RetornaNumero(!CPF)), "000\.000\.000-00")
                                        End If
                                        If Not IsNull(!NomeLogradouro) Then
                                            sEND2 = !NomeLogradouro & " Nº " & !NUMIMOVEL
                                        Else
                                            If Not IsNull(!NOMELOGRADOURO2) Then
                                                sEND2 = !NOMELOGRADOURO2 & " Nº " & !NUMIMOVEL
                                            Else
                                                sEND2 = ""
                                            End If
                                        End If
                                        sCIDADE2 = SubNull(!NOMECIDADE2)
                                        sUF2 = SubNull(!SiglaUF)
                                    End If
                                End If
                               .Close
                            End With
                            'SOMENTE OS 2 PRIMEIROS PROPRIETARIOS SERÃO GRAVADOS
                            If z > 2 Then Exit Do
                            z = z + 1
                           .MoveNext
                        Loop
                       .Close
                    End With
                    
                    ax = sLayout & ";" & !codigomob & ";" & sDoc & ";" & FillSpace(!RazaoSocial, 70) & ";" & FillSpace(SubNull(!NOMEFANTASIA), 70) & ";" & IIf(IsNull(!DATAABERTURA), FillSpace("", 10), Format(!DATAABERTURA, "dd/mm/yyyy")) & ";"
                    ax = ax & IIf(IsNull(!dataencerramento), FillSpace("", 10), Format(!dataencerramento, "dd/mm/yyyy")) & ";" & FillSpace(sCidade, 50) & ";" & FillSpace(!SiglaUF, 2) & ";" & FillLeft(!CodLogradouro, 6) & ";"
                    ax = ax & FillSpace(Val(SubNull(IIf(!CODTIPOLOG = 9999, 0, !CODTIPOLOG))), 15) & ";" & FillSpace(Val(SubNull(IIf(!CODTITLOG = 9999, 0, !CODTITLOG))), 15) & ";" & FillSpace(IIf(IsNull(!NomeLogradouro), SubNull(!NomeLogr), !NomeLogradouro), 50) & ";"
                    ax = ax & FillLeft(Val(SubNull(!Numero)), 4) & ";" & FillSpace(SubNull(!Complemento), 15) & ";" & FillSpace(RetornaNumero(RetornaCEP(!CodLogradouro, !Numero)), 8) & ";" & FillLeft(!codatividade, 6) & ";"
                    ax = ax & FillSpace(SubNull(!ativextenso), 50) & ";" & FillLeft(sCodProp1, 6) & ";" & FillSpace(sNomeProp1, 70) & ";" & FillSpace(sCPF1, 11) & ";" & FillSpace(sEND1, 100) & ";" & FillSpace(sCIDADE1, 50) & ";" & FillSpace(sUF1, 2) & ";" & FillLeft(sCodProp2, 6) & ";"
                    ax = ax & FillSpace(sNomeProp2, 70) & ";" & FillSpace(sCPF2, 11) & ";" & FillSpace(sEND2, 100) & ";" & FillSpace(sCIDADE2, 50) & ";" & FillSpace(sUF2, 2) & ";"
                    Print #1, ax
PROXIMOMOB:
                   .MoveNext
                Loop
            End With
        Close #1
    Case 3 'CADASTRO DE CIDADÃO
        Open sPathBin & "\CIDADAO.TXT" For Output As #1
            sLayout = "03"
            'CARREGA O CÓDIGO DE TODOS OS CIDADÕES QUE POSSUAM DEBITOS ATRASADOS
            Sql = "SELECT vwCIDADAO.codcidadao, vwCIDADAO.nomecidadao, vwCIDADAO.cpf, vwCIDADAO.cnpj, vwCIDADAO.codlogradouro, vwCIDADAO.nomelogradouro, "
            Sql = Sql & "vwCIDADAO.codtitlog, vwCIDADAO.abrevtitlog, vwCIDADAO.codtipolog, vwCIDADAO.abrevtipolog,"
            Sql = Sql & "vwCIDADAO.NOMELOGRADOURO2 , vwCIDADAO.numimovel, vwCIDADAO.complemento FROM debitoparcela INNER JOIN vwCIDADAO ON debitoparcela.codreduzido = vwCIDADAO.codcidadao "
            Sql = Sql & "Where (debitoparcela.AnoExercicio >= 2002) And (debitoparcela.CodLancamento <> 20) And (debitoparcela.NumParcela > 0) And (debitoparcela.statuslanc = 3) AND (DATAINSCRICAO IS NOT NULL) AND (DATAAJUIZA IS NULL) "
            Sql = Sql & "GROUP BY debitoparcela.codreduzido, vwCIDADAO.codcidadao, vwCIDADAO.nomecidadao, vwCIDADAO.cpf, vwCIDADAO.cnpj, vwCIDADAO.codlogradouro,vwCIDADAO.nomelogradouro, "
            Sql = Sql & "vwCIDADAO.codtitlog, vwCIDADAO.abrevtitlog, vwCIDADAO.codtipolog, vwCIDADAO.abrevtipolog, vwCIDADAO.NOMELOGRADOURO2, "
            Sql = Sql & "vwCIDADAO.numimovel, vwCIDADAO.complemento HAVING (codreduzido BETWEEN 500000 AND 800000) ORDER BY CODREDUZIDO"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                nTot = .RowCount: lblRegTot.Caption = nTot
                Do Until .EOF
                    nPos = nPos + 1
                    If nPos Mod 50 = 0 Then
                       CallPb nPos, CLng(nTot)
                       lblRegPerc.Caption = nPos
                    End If
                    If Not IsNull(!CPF) Then
                        sDoc = RetornaNumero(!CPF)
                    ElseIf Not IsNull(!Cnpj) Then
                        sDoc = RetornaNumero(!Cnpj)
                    Else
                        sDoc = 0
                    End If
                    sDataNascto = "00/00/0000"
                    nSexo = 1
                    If Not IsNull(!NomeLogradouro) Then
                        sEndereco = !NomeLogradouro
                    Else
                        If Not IsNull(!NOMELOGRADOURO2) Then
                            sEndereco = !NOMELOGRADOURO2
                        Else
                            sEndereco = ""
                        End If
                    End If
                    'GRAVA NO ARQUIVO
                    ax = sLayout & ";" & FillLeft(!CodCidadao, 6) & ";" & FillSpace(sDoc, 11) & ";" & FillSpace(!nomecidadao, 70) & ";" & sDataNascto & ";"
                    ax = ax & FillLeft(CStr(nSexo), 2) & ";" & FillLeft(Val(SubNull(!CodLogradouro)), 6) & ";" & FillSpace(Val(SubNull(!CODTIPOLOG)), 15) & ";"
                    ax = ax & FillSpace(Val(SubNull(IIf(!CODTITLOG = 9999, 0, !CODTITLOG))), 15) & ";" & FillSpace(sEndereco, 50) & ";" & FillLeft(Val(SubNull(!NUMIMOVEL)), 4) & ";"
                    ax = ax & FillSpace(SubNull(!Complemento), 15) & ";" & FillSpace(RetornaNumero(RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!NUMIMOVEL)))), 8)
                    Print #1, ax
                   .MoveNext
                Loop
               .Close
            End With
        Close #1
    Case 4 'SISTEMA CONSIST
        bConsist = True
        ExportaConsist
End Select
Liberado
Pb.value = 100: lblRegPerc.Caption = nTot
'If Not bConsist Then
'    x = Shell("NOTEPAD" & " " & sPathBin & "\DEBITO.TXT", vbNormalFocus)
'End If
'MsgBox "fim"
End Sub

Private Sub ExportaConsist()
Dim ax As String, aCodigos() As Long, x As Integer, Y As Integer, nPos As Long, nCPF As Byte, sDoc As String, sDataEncerra As String, RdoAux2 As rdoResultset, sTmp As String, sDataSimples As String
Dim sEnd As String, sCep As String, aAtiv(24) As ATIVIDADES, t As Integer, sAtiv As String, aAliq(24) As Double, bAchou As Boolean, nSimples As Single, nRegEspecial As Single

Open sPathBin & "\PREFJABC.TXT" For Output As #1

Print #1, "[INICIO-ANO]"
Print #1, Format(Year(Now), "0000")
Print #1, "[FIM-ANO]"

Print #1, "[INICIO-SELIC]"
Print #1, "0120100000000000000020056"
Print #1, "0220100000000000000020056"
Print #1, "0320100000000000000020056"
Print #1, "0420100000000000000020056"
Print #1, "0520100000000000000020056"
Print #1, "0620100000000000000020056"
Print #1, "0720100000000000000020056"
Print #1, "0820100000000000000020056"
Print #1, "0920100000000000000020056"
Print #1, "1020100000000000000020056"
Print #1, "1120100000000000000020056"
Print #1, "1220100000000000000020056"
Print #1, "0120110000000000000020999"
Print #1, "0220110000000000000020999"
Print #1, "0320110000000000000020999"
Print #1, "0420110000000000000020999"
Print #1, "0520110000000000000020999"
Print #1, "0620110000000000000020999"
Print #1, "0720110000000000000020999"
Print #1, "0820110000000000000020999"
Print #1, "0920110000000000000020999"
Print #1, "1020110000000000000020999"
Print #1, "1120110000000000000020999"
Print #1, "1220110000000000000020999"
Print #1, "0120120000000000000022534"
Print #1, "0220120000000000000022534"
Print #1, "0320120000000000000022534"
Print #1, "0420120000000000000022534"
Print #1, "0520120000000000000022534"
Print #1, "0620120000000000000022534"
Print #1, "0720120000000000000022534"
Print #1, "0820120000000000000022534"
Print #1, "0920120000000000000022534"
Print #1, "1020120000000000000022534"
Print #1, "1120120000000000000022534"
Print #1, "1220120000000000000022534"
Print #1, "0120130000000000000023724"
Print #1, "0220130000000000000023724"
Print #1, "0320130000000000000023724"
Print #1, "0420130000000000000023724"
Print #1, "0520130000000000000023724"
Print #1, "0620130000000000000023724"
Print #1, "0720130000000000000023724"
Print #1, "0820130000000000000023724"
Print #1, "0920130000000000000023724"
Print #1, "1020130000000000000023724"
Print #1, "1120130000000000000023724"
Print #1, "1220130000000000000023724"
Print #1, "0120140000000000000025114"
Print #1, "0220140000000000000025114"
Print #1, "0320140000000000000025114"
Print #1, "0420140000000000000025114"
Print #1, "0520140000000000000025114"
Print #1, "0620140000000000000025114"
Print #1, "0720140000000000000025114"
Print #1, "0820140000000000000025114"
Print #1, "0920140000000000000025114"
Print #1, "1020140000000000000025114"
Print #1, "1120140000000000000025114"
Print #1, "1220140000000000000025114"
Print #1, "0120150000000000000026808"
Print #1, "0220150000000000000026808"
Print #1, "0320150000000000000026808"
Print #1, "0420150000000000000026808"
Print #1, "0520150000000000000026808"
Print #1, "0620150000000000000026808"
Print #1, "0720150000000000000026808"
Print #1, "0820150000000000000026808"
Print #1, "0920150000000000000026808"
Print #1, "1020150000000000000026808"
Print #1, "1120150000000000000026808"
Print #1, "1220150000000000000026808"
Print #1, "[FIM-SELIC]"
Print #1, ""

Print #1, ""
Print #1, "[INICIO-ATIVIDADE]"
Sql = "SELECT DISTINCT ATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE,ATIVIDADEISS.ITEM ,ATIVIDADEISS.RETIDO,TABELAISS.TIPOISS,TABELAISS.ALIQUOTA * 100 as aliquota,TABELAISS.DATA "
Sql = Sql & "FROM ATIVIDADEISS INNER JOIN TABELAISS ON ATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV WHERE CODATIVIDADE >=200 AND (atividadeiss.imprimir = 1) ORDER BY codatividade,DATA  "
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        If !TipoISS <> 11 Then
            ax = Format(!Item, "00000000") & FillSpace(!descatividade, 300) & Format(!Aliquota * 100, "0000") & Format(Day(!Data), "00") & Format(Month(!Data), "00") & Year(!Data) & !RETIDO
'            ax = Format(!codatividade, "00000000") & FillSpace("Item:" & !Item & " " & !descatividade, 300) & Format(!Aliquota * 100, "0000") & Format(Day(!Data), "00") & Format(Month(!Data), "00") & Year(!Data)
        Else
            ax = Format(!Item, "00000000") & FillSpace(!descatividade, 300) & Format(0, "0000") & Format(Day(!Data), "00") & Format(Month(!Data), "00") & Year(!Data) & !RETIDO
'            ax = Format(!codatividade, "00000000") & FillSpace("Item:" & !Item & " " & !descatividade, 300) & Format(0, "0000") & Format(Day(!Data), "00") & Format(Month(!Data), "00") & Year(!Data)
        End If
        Print #1, ax
       .MoveNext
    Loop
   .Close
End With
Print #1, "[FIM-ATIVIDADE]"
 
Print #1, ""
Print #1, "[INICIO-EMPRESA]"
ReDim aCodigos(0)
'CARREGA APENAS AS EMPRESAS VARIAVEL E ESTIMADO
'Sql = "SELECT DISTINCT codmobiliario From mobiliarioatividadeiss Where (CodTributo <> 11) And (codmobiliario > 100000)  ORDER BY CODMOBILIARIO"
Sql = "SELECT DISTINCT CODIGOMOB FROM VWCNSMOBILIARIO WHERE (DATAENCERRAMENTO > '10/1/2007'or dataencerramento is null) ORDER BY CODIGOMOB"
'Sql = "SELECT DISTINCT CODIGOMOB FROM VWCNSMOBILIARIO WHERE (DATAENCERRAMENTO > '10/1/2007'or dataencerramento is null) AND CODIGOMOB=112757 ORDER BY CODIGOMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCodigos(UBound(aCodigos) + 1)
        aCodigos(UBound(aCodigos)) = !codigomob
       .MoveNext
    Loop
   .Close
End With

'CARREGA DADOS DAS EMPRESAS
nPos = 0: Pb.value = 0
nTot = UBound(aCodigos): lblRegTot.Caption = nTot
For x = 1 To UBound(aCodigos)
    nPos = x
    CallPb nPos, CLng(nTot)
    lblRegPerc.Caption = nPos
'   Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & aCodigos(x) & " and  DATAENCERRAMENTO > '10/1/2007' "
'If aCodigos(x) = 100190 Then MsgBox "teste"
    DoEvents
    Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & aCodigos(x) & " and (DATAENCERRAMENTO > '10/1/2007'or dataencerramento is null)"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    With RdoAux
        If .RowCount = 0 Then GoTo proximo
        ReDim aSimples(6)
        For Y = 1 To 5
            aSimples(Y).sDataIni = String(8, " ")
            aSimples(Y).sDataFim = String(8, " ")
        Next
        
        nSimples = SNCheck(aCodigos(x))
        nRegEspecial = Val(SubNull(!REGESPECIAL))
        nCPF = IIf(IsNull(!Cnpj), 1, 0)
        sDoc = IIf(IsNull(!CPF), SubNull(!Cnpj), SubNull(!CPF))
        sDataEncerra = "        "
        If Not IsNull(!dataencerramento) Then
            sDataEncerra = Format(!dataencerramento, "yyyymmdd")
'            GoTo Proximo
        End If
        
'        If aCodigos(x) = 101018 Then
'            MsgBox "tedte2"
'        End If
        
       'SUSPENÇÃO
        Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !codigomob & " ORDER BY DATAEVENTO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If !CODTIPOEVENTO = 2 Then
                    sDataEncerra = Format(!DATAPROCEVENTO, "yyyymmdd")
                    'GoTo Proximo
                End If
            End If
           .Close
        End With
        
        If !CodCidade = 413 Then
            sCep = RetornaCEP(!CodLogradouro, !Numero)
        Else
            sCep = Left$(!Cep, 5) & "-" & Right$(!Cep, 3)
        End If
        
       'ENDERECO
        sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & ", " & !Numero & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
        sEnd = sEnd & !DescBairro & " - " & !descCidade & " - " & !SiglaUF & " - " & sCep
        sEnd = Left$(sEnd, 200)
        ax = Format(!codigomob, "00000000") & CStr(nCPF) & Format(Val(sDoc), "00000000000000") & FillSpace(!RazaoSocial, 100) & FillSpace(sDataEncerra, 8) & FillSpace(sEnd, 200)
       .Close
    End With
   'ATIVIDADES
    For t = 1 To 24
        aAtiv(t).nCodigo = 0
        aAtiv(t).nSeq = 0
        aAliq(t) = 0
        aAtiv(t).sEstimado = ""
        aAtiv(t).sFixo = ""
        aAtiv(t).sTipoIss = ""
    Next
    'Sql = "INSERT PERIODOIE(CODIGO,DATAINI) VALUES(" & aCodigos(x) & ",'01/01/2010')"
    'cn.Execute Sql, rdExecDirect
    
   'CARREGA ATIVIDADE DISTINTAS
    Sql = "SELECT DISTINCT CODTRIBUTO,VALORISS FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & aCodigos(x)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
    With RdoAux
        Do Until .EOF
            If !CodTributo = 11 Then
                aAliq(.AbsolutePosition) = 0
            Else
                aAliq(.AbsolutePosition) = !valoriss
            End If
           .MoveNext
        Loop
       .Close
    End With
    'CARREGA AS ATIVIDADES E SUA SEQUENCIA
   ' If aCodigos(x) = 112757 Then MsgBox "teste"
    Sql = "SELECT mobiliarioatividadeiss.codatividade, mobiliarioatividadeiss.valoriss, mobiliarioatividadeiss.codtributo, atividadeiss.isseletronico,item "
    Sql = Sql & "FROM mobiliarioatividadeiss LEFT OUTER JOIN atividadeiss ON mobiliarioatividadeiss.codatividade = atividadeiss.codatividade where codmobiliario=" & aCodigos(x)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
    With RdoAux
        Do Until .EOF
            For t = 1 To UBound(aAliq)
                If aAliq(t) = !valoriss Then
                    Exit For
                End If
            Next
            'aAtiv(.AbsolutePosition).nCodigo = !codatividade
            'If IsNull(!Item) Then
            '    GoTo proxativ
            'End If
            aAtiv(.AbsolutePosition).nCodigo = RetornaNumero(IIf(IsNull(!Item), 0, !Item))
 '           If !codatividade <= 200 And Trim(sDataEncerra) = "" Then
'                MsgBox aCodigos(x)
  '          End If
            aAtiv(.AbsolutePosition).nSeq = t
            aAtiv(.AbsolutePosition).sEstimado = IIf(!CodTributo = 12, "X", " ")
            aAtiv(.AbsolutePosition).sFixo = IIf(!CodTributo = 11, "X", " ")
            If nSimples = 1 Then
                aAtiv(.AbsolutePosition).sISSEletronico = IIf(!isseletronico = 1, "S", "N")
            Else
                aAtiv(.AbsolutePosition).sISSEletronico = "N"
            End If
            If !CodTributo = 11 Then
                aAtiv(.AbsolutePosition).sTipoIss = "F"
            ElseIf !CodTributo = 12 Then
                aAtiv(.AbsolutePosition).sTipoIss = "E"
            ElseIf !CodTributo = 13 Then
                aAtiv(.AbsolutePosition).sTipoIss = "V"
            Else
                aAtiv(.AbsolutePosition).sTipoIss = " "
            End If
proxativ:
           .MoveNext
        Loop
       .Close
    End With
    
    sAtiv = ""
    For t = 1 To 24
        sAtiv = sAtiv & Format(aAtiv(t).nCodigo, "00000000")
    Next
    For t = 1 To 24
        sAtiv = sAtiv & Format(aAtiv(t).nSeq, "00")
    Next
    
    For t = 1 To 24
        sAtiv = sAtiv & IIf(aAtiv(t).sEstimado = "", " ", aAtiv(t).sEstimado)
    Next
    ax = ax & sAtiv
    ax = ax & IIf(nSimples = 1, "X", " ")
    ax = ax & IIf(nRegEspecial = 1, "X", " ")
    sAtiv = ""
    For t = 1 To 24
        sAtiv = sAtiv & IIf(aAtiv(t).sFixo = "", " ", aAtiv(t).sFixo)
    Next
    ax = ax & sAtiv
    sAtiv = ""
    For t = 1 To 24
        sAtiv = sAtiv & IIf(aAtiv(t).sTipoIss = "", " ", aAtiv(t).sTipoIss)
    Next

    
    'CARREGA AS DATAS DO SIMPLES
    Sql = "SELECT * FROM PERIODOSN WHERE CODIGO=" & aCodigos(x) & " ORDER BY DATAINI DESC"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
    With RdoAux
        Do Until .EOF
            If .AbsolutePosition > 5 Then Exit Do
            aSimples(.AbsolutePosition).sDataIni = Format(!dataini, "dd/mm/yyyy")
            If Not IsNull(!datafim) Then
                aSimples(.AbsolutePosition).sDataFim = Format(!datafim, "dd/mm/yyyy")
            End If
           .MoveNext
        Loop
       .Close
    End With
    
    sDataSimples = ""
    For Y = 5 To 1 Step -1
        'sDataSimples = sDataSimples & Left(aSimples(Y).sDataIni, 2) & Mid(aSimples(Y).sDataIni, 4, 2) & Right(aSimples(Y).sDataIni, 4) & Left(aSimples(Y).sDataFim, 2) & Mid(aSimples(Y).sDataFim, 4, 2) & Right(aSimples(Y).sDataFim, 4)
        sDataSimples = sDataSimples & Right(aSimples(Y).sDataIni, 4) & Mid(aSimples(Y).sDataIni, 4, 2) & Left(aSimples(Y).sDataIni, 2) & Right(aSimples(Y).sDataFim, 4) & Mid(aSimples(Y).sDataFim, 4, 2) & Left(aSimples(Y).sDataFim, 2)
    Next
    ax = ax & sDataSimples & sAtiv
    Print #1, ax
        

proximo:
Next

'******* empresa retensao iss

'CARREGA DADOS DAS EMPRESAS COM RETENÇÃO NA FONTE
Sql = "SELECT CODREDUZIDO FROM ISSRETIDO ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 0: Pb.value = 0
    nTot = .RowCount: lblRegTot.Caption = nTot
    Do Until .EOF
        nPos = .AbsolutePosition
        CallPb nPos, CLng(nTot)
        lblRegPerc.Caption = nPos
        Sql = "SELECT * FROM VWCIDADAO WHERE CODCIDADAO=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        With RdoAux2
            If .RowCount > 0 Then
                nCPF = IIf(IsNull(!Cnpj), 1, 0)
                sDoc = IIf(IsNull(!CPF) Or Trim(!CPF) = "", Trim(SubNull(!Cnpj)), Trim(SubNull(!CPF)))
                sCep = ""
                If !CodCidade = 413 Then
                    If Not IsNull(!NUMIMOVEL) And !CodLogradouro > 0 Then
                        sCep = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
                    ElseIf Not IsNull(!NUMIMOVEL) And Not IsNull(!Cep) Then
                        sCep = !Cep
                    End If
                Else
                    If Not IsNull(!Cep) Then
                        sCep = Left$(!Cep, 5) & "-" & Right$(!Cep, 3)
                    End If
                End If
                If Not IsNull(!NomeLogradouro) Then
                   If !NomeLogradouro <> "" Then
                       sEnd = Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL)) & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
                   Else
                       sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
                   End If
                Else
                   sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
                End If
'                If Not IsNull(!NOMEBairro) Then
'                   sEnd = sEnd & SubNull(!NOMEBairro) & " - " & SubNull(!NomeCidade) & " - " & !NOMEUF & " - " & sCEP
'                Else
                   sEnd = sEnd & SubNull(!DescBairro) & " - " & SubNull(!descCidade) & " - " & !SiglaUF & " - " & sCep
                'End If
                sEnd = Left$(sEnd, 200)
                
                ax = Format(!CodCidadao, "00000000") & CStr(nCPF) & Format(Val(sDoc), "00000000000000") & FillSpace(!nomecidadao, 100) & FillSpace("", 8) & FillSpace(sEnd, 200) & String(240, "0") & String(26, " ")
                Print #1, ax
            End If
           .Close
        End With
PROXIMO2:
       .MoveNext
    Loop
End With

Print #1, "[FIM-EMPRESA]"
Print #1, ""
Print #1, "[INICIO-VENCIMENTOS]"
Print #1, "200701102007021220070312200704102007051020070611200707102007081020070910200710102007111620071217"
Print #1, "200801152008021520080317200804152008051520080616200807152008081520080915200810152008111720081215"
Print #1, "200901152009021620090316200904152009051520090615200907152009081720090915200910152009111620091215"
Print #1, "201001152010021720100315201004152010051720100615201007152010081620100915201010152010111620101215"
Print #1, "201101172011021520110315201104152011051620110615201107152011081520110915201110172011111620111215"
Print #1, "201201162012021520120315201204162012051520120615201207172012081520120917201210152012111620121217"
Print #1, "201301152013021520130315201304152013051520130617201307152013081520130916201310152013111820131216"
Print #1, "201401152014021720140317201404152014051520140616201407152014081520140915201410152014111720141215"
Print #1, "201501152015021620150315201504152015051520150615201507152015081720150915201510152015111620151215"
Print #1, "[FIM-VENCIMENTOS]"


Print #1, ""
Print #1, "[INICIO-GUIAS-PAGAS]"
Sql = "SELECT guiaisseletronico.numero, guiaisseletronico.datavencto, guiaisseletronico.valorprincipal, guiaisseletronico.valoracrescimo, "
Sql = Sql & "debitopago.DataPagamento , NumDocumento.ValorPago FROM numdocumento INNER JOIN guiaisseletronico ON numdocumento.numdocumento = guiaisseletronico.numero LEFT OUTER JOIN "
Sql = Sql & "debitopago ON guiaisseletronico.numero = debitopago.numdocumento Where (NumDocumento.ValorPago > 0)  AND (YEAR(debitopago.datapagamento) >= YEAR(GETDATE()) - 1) "
Sql = Sql & " ORDER BY guiaisseletronico.numero"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ax = Format(!Numero, "000000000000")
        If Not IsNull(!DataPagamento) Then
            ax = ax & Format(Year(!DataPagamento), "0000") & Format(Month(!DataPagamento), "00") & Format(Day(!DataPagamento), "00")
        Else
            ax = ax & Format(Year(!DataVencto), "0000") & Format(Month(!DataVencto), "00") & Format(Day(!DataVencto), "00")
        End If
        ax = ax & Format(!ValorPrincipal, "000000000.00")
        ax = ax & Format(!ValorAcrescimo, "000000000.00")
        Print #1, ax
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT DISTINCT damiss.dociss, debitopago.datapagamento, debitopago.valorpago FROM damiss INNER JOIN "
Sql = Sql & "debitopago ON damiss.docdam = debitopago.numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ax = Format(!dociss, "000000000000")
        ax = ax & Format(Year(!DataPagamento), "0000") & Format(Month(!DataPagamento), "00") & Format(Day(!DataPagamento), "00")
        ax = ax & Format(!ValorPago, "000000000.00")
        ax = ax & Format(0, "000000000.00")
        Print #1, ax
       .MoveNext
    Loop
   .Close
End With


Print #1, "[FIM-GUIAS-PAGAS]"
fim:
Print #1, ""
Print #1, "[INICIO-FERIADOS]"
Print #1, "[FIM-FERIADOS]"

Close #1
Liberado
MsgBox "Exportação finalizada com sucesso.", vbInformation, "Operação finalizada"
End Sub


Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents
If nPosF > 0 Then
    Pb.Color = &H808000
Else
    Pb.Color = vbWhite
End If
If nTotal = 0 Then nTotal = 1
If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Sub CalculoIndividual(nCodReduz As Long)
Dim nSomaTestada As Double, nAreaTerrenoReal As Double
Dim nUso As Integer, nTipo As Integer, nCat As Integer, nCodBairro As Integer
Dim bIsento As Boolean, nTestada1 As Double, x As Integer

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
    nAreaTerreno = !Dt_AreaTerreno
    nAreaTerrenoReal = nAreaTerreno
    nCodSituacao = !Dt_CodSituacao
    nCodPedologia = !Dt_CodPedol
    nCodTopografia = !Dt_CodTopog
    nCodAgrupamento = !CODAGRUPA
    bFracaoIdeal = IIf(!Dt_FracaoIdeal > 0, True, False)
    If bFracaoIdeal Then nAreaTerreno = !Dt_FracaoIdeal
    'TEM ÁREA?
    If Not IsNull(!SOMAAREA) Then
        nAreaTotal = !SOMAAREA
        bTemPredial = True
        nAreaPrincipal = FormatNumber(!SOMAAREA, 2)
    Else
        nAreaTotal = 0
        bTemPredial = False
        nAreaPrincipal = 0
    End If
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
    'FRAÇÃO IDEAL PARA CALCULO DE TESTADA
    '--Se houver Fracao Ideal o Comprimento da Testada e calculado por --> FRACAOIDEAL * TESTADA / AREA PRINCIPAL
    
    'BUSCA ÁREA PRINCIPAL
    Sql = "SELECT AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR,QTDEPAV FROM AREAS WHERE CODREDUZIDO = " & nCodReduz & "  AND TIPOAREA='P'"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux2
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
    '**************************
    '### FATOR PROFUNDIDADE ###
    '**************************
    If !Dt_CodUsoTerreno <> 6 Then 'gleba não tem profundidade
        '*** PROFUNDIDADE = AREA DO TERRENO / TESTADA PRINCIPAL DO LOTE
         nValorProfundidade = FormatNumber(nAreaTerrenoReal / nTestada1, 2)
        'LOCALIZAMOS PRIMEIRO O CODIGO DA PROFUNDIDADE A QUE PERTENCE O IMOVEL
        For x = 1 To UBound(aProf)
            If aProf(x).Distrito = !Distrito Then
               If nValorProfundidade >= aProf(x).Min And nValorProfundidade <= aProf(x).Max Then
                  Exit For
               ElseIf nValorProfundidade >= aProf(x).Min And aProf(x).Max = 0 Then
                  Exit For
               End If
            End If
        Next
        If x <= UBound(aProf) Then
            nCodProfundidade = aProf(x).Codigo
        Else
            nCodProfundidade = 1
        End If
        'PROCURAMOS AGORA O VALOR DO FATOR PROFUNDIDADE
        nFatorProfundidade = 0
        For x = 1 To UBound(aFatorF)
            If aFatorF(x).Distrito = !Distrito And aFatorF(x).Codigo = nCodProfundidade Then
               nFatorProfundidade = aFatorF(x).Fator
               Exit For
            End If
        Next
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
    nVVT = nValorVenalTerritorial
    'CÁLCULO VALOR VENAL PREDIAL
    '--VALOR VENAL PREDIAL = £(AREA CONSTRUIDA * PADRAO CONSTRUCAO DA AREA PRINCIPAL)
    If bTemPredial Then
        '**************************
        '### FATOR DISTRITO ###
        '**************************
        nFatorDistrito = aFatorD(!Distrito)
        '**************************
        '### FATOR CATEGORIA ###
        '**************************
        nValorVenalPredial = 0
        nFatorCategoria = 0
        For x = 1 To UBound(aFatorC)
            If aFatorC(x).Uso = nUso And aFatorC(x).Tipo = nTipo And aFatorC(x).Categoria = nCat Then
               nFatorCategoria = aFatorC(x).Fator
               Exit For
            End If
        Next
        nValorVenalPredial = nValorVenalPredial + (FormatNumber(nAreaPrincipal, 2) * FormatNumber(nFatorCategoria, 2))
       'FATOR CATEGORIA 98
        nValorVenalPredial = nValorVenalPredial * nFatorDistrito
        nVVP = nValorVenalPredial
    Else
        nFatorDistrito = 0
        nFatorCategoria = 0
        nVVP = 0
    End If
End With
End Sub

Private Sub LoadMatrix()
Dim nAnoCalculo As Integer
ReDim aFatorD(3)
ReDim aFatorP(6)
ReDim aFatorT(6)
ReDim aFatorS(6)
ReDim aFatorG(23)
ReDim aFatorR(8)
nAnoCalculo = Year(Now)
Sql = "SELECT CODPEDOLOGIA,FATORPEDOLOGIA FROM FATORPEDOLOGIA WHERE ANOPEDOLOGIA=" & nAnoCalculo & " ORDER BY CODPEDOLOGIA; " & _
      "SELECT CODTOPOG,FATORTOPOG FROM FATORTOPOGRAFIA WHERE ANOTOPOG=" & nAnoCalculo & " ORDER BY CODTOPOG; " & _
      "SELECT CODSITUACAO,FATORSITUACAO FROM FATORSITUACAO WHERE ANOSITUACAO=" & nAnoCalculo & " ORDER BY CODSITUACAO; " & _
      "SELECT CODGLEBA,FATORGLEBA FROM FATORGLEBA WHERE ANOGLEBA=" & nAnoCalculo & " ORDER BY CODGLEBA; " & _
      "SELECT CODDISTRITO,FATORDISTRITO FROM FATORDISTRITO WHERE ANODISTRITO=" & nAnoCalculo & " ORDER BY CODDISTRITO; " & _
      "SELECT CODAGRUPAMENTO, VALORTERRENO  FROM TERRENO  WHERE ANOFATOR=" & nAnoCalculo & "  AND  CODMOEDA=1 "
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
Sql = "SELECT CODDISTRITO,CODPROFUN,FATORPROFUN FROM FATORPROFUN WHERE ANOPROFUN=" & nAnoCalculo & " ORDER BY CODDISTRITO,CODPROFUN "

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
Sql = "SELECT CODUSO,CODTIPO,CODCATEG,FATORCATEG FROM FATORCATEG WHERE ANOCATEG=" & nAnoCalculo & " AND CODMOEDA=1 "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
        ReDim Preserve aFatorC(UBound(aFatorC) + 1)
        aFatorC(UBound(aFatorC)).Uso = !coduso
        aFatorC(UBound(aFatorC)).Tipo = !CodTipo
        aFatorC(UBound(aFatorC)).Categoria = !codcateg
        aFatorC(UBound(aFatorC)).Fator = !FATORCATEG
       .MoveNext
     Loop
    .Close
End With

End Sub

Private Sub Importa()
Dim nCodReduz As Long, nAno As Integer, nParc As Integer, dDataVencto As Date, RdoAux2 As rdoResultset, nLanc As Integer, nSeq As Integer, nCompl As Integer, dDataAjuiza As Date
Dim nTot As Long, nPos As Long
'Exit Sub
Pb.value = 0
Sql = "SELECT CODREDUZIDO,ANOEXERCICIO,NUMPARCELA,DATAVENCIMENTO,CODTRIBUTO,DATAAJUIZAMENTO FROM IMPORTACAODEBITOAJUIZADO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        nPos = nPos + 1
        CallPb nPos, CLng(nTot)
        nCodReduz = !CODREDUZIDO
        nAno = !AnoExercicio
        nParc = !NumParcela
        dDataVencto = !DataVencimento
        dDataAjuiza = !DATAAJUIZAMENTO
        Sql = "SELECT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz
        Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND NUMPARCELA=" & nParc & " AND DATAVENCIMENTO='" & Format(dDataVencto, "mm/dd/yyyy") & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                Do Until .EOF
                    If .RowCount > 1 Then
                        If !statuslanc <> 3 Then
                            'MsgBox "STATUSLANC"
                            GoTo proximo
                        End If
                    End If
                    nLanc = !CodLancamento
                    nSeq = !SeqLancamento
                    nCompl = !CODCOMPLEMENTO
                    Sql = "UPDATE DEBITOPARCELA SET DATAAJUIZA='" & Format(dDataAjuiza, "mm/dd/yyyy") & "' WHERE CODREDUZIDO=" & nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                    cn.Execute Sql, rdExecDirect
proximo:
                   .MoveNext
                Loop
            Else
                MsgBox "nao achei"
            End If
           .Close
        End With
       .MoveNext
    Loop
   .Close
End With
MsgBox "IMPORTAÇÃO CONCLUIDA.", vbInformation, "INFORMAÇÃO"

End Sub


Private Sub ImportaConsist2()
Dim strLinha As String, aLinha() As ISSELETRO, x As Integer, nSeq As Integer, sDataVencto As String, nPos As Integer, nPos2 As Integer
Dim sGuia As String, sDoc As String, sRazao As String, nDoc As Long, sCNPJ As String
Dim nLastCod As Long, nCodReduz As Long, aGuia() As GUIAAVULSA, bAchou As Boolean, nCompl As Integer

If MsgBox("Deseja importar o arquivo de guias do ISS Eletrônico ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
ReDim aLinha(0): ReDim aGuia(0)
 
If Dir(sPathBin & "\GUIAS.TXT") = "" Then
    MsgBox "O Arquivo 'GUIAS.TXT' não foi localizado, certifique-se de ter copiado " & vbCrLf & "para o diretório 'BIN' do diretório raiz do sistema.", vbCritical, "Atenção"
    Exit Sub
End If

Open sPathBin & "\GUIAS.TXT" For Input As #1
   Do While Not EOF(1)
        Line Input #1, strLinha
        If Left(strLinha, 1) = 1 Then
            GoTo PROXIMO2
        End If
        If Mid(strLinha, 1, 13) = "0201008021295" Then MsgBox "teste"
        ReDim Preserve aLinha(UBound(aLinha) + 1)
        aLinha(UBound(aLinha)).Identificacao = Mid(strLinha, 1, 1)
        aLinha(UBound(aLinha)).Numero = Mid(strLinha, 2, 12)
        
        If Val(Mid(strLinha, 14, 8)) > 0 Then
            aLinha(UBound(aLinha)).Inscricao = Mid(strLinha, 14, 8)
        Else
            '*****************************************************************************
            '*********** IDENTIFICAÇÃO DO CNPJ, CPF OU CÓDIGO ****************************
            '*****************************************************************************
            sDoc = Mid(strLinha, 91, 14)
            sRazao = UCase(Mid(strLinha, 105, 50))
            sCNPJ = RetornaNumero(sDoc)
            Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB=" & RetornaNumero(sCNPJ)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    nCodReduz = !codigomob
                Else
                    MsgBox "Código não encontrado, empresa: " & sRazao & " doc: " & sCNPJ, vbExclamation, "Atenção"
                    nCodReduz = 0
                    Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & RetornaNumero(sCNPJ)
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        If .RowCount > 0 Then
                            nCodReduz = !CodCidadao
                        Else
                            If ValidaCGC(sCNPJ) Then
                                Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & sCNPJ & "'"
                                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux
                                    If .RowCount > 0 Then
                                        nCodReduz = !CodCidadao
                                    Else
                                        'NAO ACHOU, CADASTRA UM NOVO
                                        Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
                                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                        With RdoAux2
                                            nCodReduz = !maximo + 1
                                           .Close
                                        End With
                                        Sql = "INSERT CIDADAO (CODCIDADAO,NOMECIDADAO,CNPJ) VALUES(" & nCodReduz & ",'"
                                        Sql = Sql & UCase(Left(sRazao, 50)) & "','" & sCNPJ & "')"
                                   '     cn.Execute Sql, rdExecDirect
                                    End If
                                   .Close
                                End With
                            Else
                                'CNPJ INVALIDO
                                'TESTA CPF
                                If ValidaCPF(Right(sCNPJ, 11)) Then
                                    Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CPF='" & sCNPJ & "'"
                                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                    With RdoAux
                                        If .RowCount > 0 Then
                                            nCodReduz = !CodCidadao
                                        Else
                                            'NAO ACHOU, CADASTRA UM NOVO
                                            Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
                                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                            With RdoAux2
                                                nCodReduz = !maximo + 1
                                               .Close
                                            End With
                                            Sql = "INSERT CIDADAO (CODCIDADAO,NOMECIDADAO,CPF) VALUES(" & nCodReduz & ",'"
                                            Sql = Sql & Left(sRazao, 50) & "','" & sCNPJ & "')"
                                        '    cn.Execute Sql, rdExecDirect
                                        End If
                                       .Close
                                    End With
                                Else
                                    'CPF INVALIDO
                                    nCodReduz = 0
                                End If
                            End If
                        End If
                       .Close
                    End With
                End If
               .Close
            End With
            
            If nCodReduz > 0 Then
                aLinha(UBound(aLinha)).Inscricao = nCodReduz
            Else
                GoTo PROXIMO2
            End If
            
            '*****************************************************************************
        End If
        
        aLinha(UBound(aLinha)).Sequencia = Mid(strLinha, 22, 2)
        aLinha(UBound(aLinha)).Ano = Mid(strLinha, 24, 4)
        aLinha(UBound(aLinha)).Mes = Mid(strLinha, 28, 2)
        If aLinha(UBound(aLinha)).Identificacao = 0 Then
            aLinha(UBound(aLinha)).Aliquota = CDbl(Mid(strLinha, 30, 5)) / 100
            aLinha(UBound(aLinha)).Tipo = Mid(strLinha, 35, 2)
            If aLinha(UBound(aLinha)).Mes <> 12 Then
                aLinha(UBound(aLinha)).DataVencto = "15" & "/" & aLinha(UBound(aLinha)).Mes + 1 & "/" & aLinha(UBound(aLinha)).Ano
            Else
                aLinha(UBound(aLinha)).DataVencto = "15" & "/" & "1" & "/" & aLinha(UBound(aLinha)).Ano + 1
            End If
            aLinha(UBound(aLinha)).ValorPrincipal = CDbl(Mid(strLinha, 45, 11)) / 100
            If Mid(strLinha, 55, 11) <> "" Then
                aLinha(UBound(aLinha)).ValorAcrescimo = CDbl(Mid(strLinha, 56, 11)) / 100
                aLinha(UBound(aLinha)).DataExporta = Mid(strLinha, 73, 2) & "/" & Mid(strLinha, 71, 2) & "/" & Mid(strLinha, 67, 4)
            Else
                aLinha(UBound(aLinha)).ValorAcrescimo = "0"
                aLinha(UBound(aLinha)).DataExporta = ""
            End If
            
        Else
            aLinha(UBound(aLinha)).Aliquota = "0"
            aLinha(UBound(aLinha)).Tipo = "0"
            aLinha(UBound(aLinha)).DataVencto = "01/01/1900"
            aLinha(UBound(aLinha)).ValorPrincipal = "0"
            aLinha(UBound(aLinha)).ValorAcrescimo = "0"
            aLinha(UBound(aLinha)).DataExporta = Mid(strLinha, 73, 2) & "/" & Mid(strLinha, 71, 2) & "/" & Mid(strLinha, 67, 4)
        End If
PROXIMO2:
   Loop
Close #1


For x = 1 To UBound(aLinha)
    'If x = 817 Then MsgBox "teste"
    With aLinha(x)
        If Val(.Ano) = 0 Then
            GoTo proximo
        End If
        If Val(.Identificacao) = 0 Then
      '      If Val(.Inscricao) = 700145 Then
      '          s = 1
      '      End If
            If Val(.Inscricao) < 100000 Or Val(.Inscricao) > 800000 Then
                GoTo proximo
            End If
            If Val(.Inscricao) > 800000 Then GoTo proximo
            Sql = "SELECT NUMERO FROM GUIAISSELETRONICO WHERE NUMERO=" & Val(Right(.Numero, 6)) + 3000000
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
            With RdoAux
                If .RowCount > 0 Then
                    .Close
                    GoTo proximo
                End If
               .Close
            End With
            
            Sql = "INSERT guiaisseletronico(numero,inscricao,sequencia,ano,mes,aliquota,tipo,datavencto,valorprincipal,valoracrescimo,"
            Sql = Sql & "dataexportacao,identificacao) values(" & Val(Right(.Numero, 6)) + 3000000 & "," & Val(.Inscricao) & "," & Val(.Sequencia) & "," & Val(.Ano) & ","
            Sql = Sql & Val(.Mes) & "," & Virg2Ponto(CStr(.Aliquota)) & "," & Val(.Tipo) & ",'" & Format(.DataVencto, "mm/dd/yyyy") & "',"
            Sql = Sql & Virg2Ponto(CStr(.ValorPrincipal)) & "," & Virg2Ponto(CStr(.ValorAcrescimo)) & ",'" & Format(.DataExporta, "mm/dd/yyyy") & "',"
            Sql = Sql & Val(.Identificacao) & ")"
            cn.Execute Sql, rdExecDirect
            
            'CRIA DÉBITO DE ISS VARIAVEL NO GTI
            
            'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
            Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & Val(.Inscricao) & ") AND ANOEXERCICIO=" & Val(.Ano) & " And (CodLancamento = 5) "
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If IsNull(!maximo) Then
                    nSeq = 0
                Else
                    nSeq = !maximo + 1
                End If
               .Close
            End With
            
'            If Val(.Inscricao) = 521131 Then MsgBox "TESTE"
            sDataVencto = .DataVencto
            'CRIAR PARCELA DE ISS VARIAVEL NESTE MES E ANO COM O VENCIMENTO QUE VEIO DO BANCO
'            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'            Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & Val(.Inscricao) & "," & Val(.Ano) & "," & 5 & "," & nSeq & ","
'            Sql = Sql & 1 & "," & nCompl & ",3,'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,'GTI')"
            Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
            Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & Val(.Inscricao) & "," & Val(.Ano) & "," & 5 & "," & nSeq & ","
            Sql = Sql & 1 & "," & nCompl & ",3,'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,236)"
            cn.Execute Sql, rdExecDirect
            'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
            Sql = Sql & "VALORTRIBUTO) VALUES(" & Val(.Inscricao) & "," & Val(.Ano) & "," & 5 & "," & nSeq & ","
            Sql = Sql & 1 & "," & nCompl & "," & 13 & "," & Virg2Ponto(CStr(.ValorPrincipal)) & ")"
            'Sql = Sql & 1 & "," & 0 & "," & 13 & "," & Virg2Ponto(.ValorPrincipal + .ValorAcrescimo) & ")"
            cn.Execute Sql, rdExecDirect
            'CRIAR O DOCUMENTO PARA ELA
            On Error Resume Next
           ' Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA) VALUES(" & Val(Right(.Numero, 6)) + 3000000 & ",'"
            'Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & ")"
            'cn.Execute Sql, rdExecDirect
            'On Error GoTo 0
            ''CRIAR A PARCELADOCUMENTO
            'Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & Val(.Inscricao) & "," & Val(.Ano) & "," & 5 & "," & nSeq & ","
            'Sql = Sql & 1 & "," & nCompl & "," & Val(Right(.Numero, 6)) + 3000000 & ")"
            'cn.Execute Sql, rdExecDirect
        
        End If
    End With
    
proximo:
Next

MsgBox "Arquivo importado com sucesso.", vbInformation, "Informação"

End Sub

Private Sub ImportaConsistNovo(sArq As String)
Dim strLinha As String, aLinha() As ISSELETRONOVO, x As Long, nSeq As Integer, sDataVencto As String, nPos As Integer, nPos2 As Integer
Dim sGuia As String, sDoc As String, nDoc As Long, nCodReduzErrado As Long, nLanc As Integer, nParc As Integer, nAno As Integer
Dim nLastCod As Long, nCodReduz As Long, aGuia() As GUIAAVULSA, bAchou As Boolean, nCompl As Integer, bTomador As Boolean
Dim sInscricao As String, sCPF As String, sCNPJ As String, sRazao As String, k As Integer, nSeq2 As Integer, sObs As String
Dim bErro As Boolean

If MsgBox("Deseja importar o arquivo de guias do ISS Eletrônico (novo layout)?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
ReDim aLinha(0): ReDim aGuia(0)
 
Ocupado

Open sArq For Input As #1
   Do While Not EOF(1)
        DoEvents
        Line Input #1, strLinha
        If Left(strLinha, 1) = 1 Then
            ReDim Preserve aLinha(UBound(aLinha) + 1)
            aLinha(UBound(aLinha)).sDataGeracao = Mid(strLinha, 8, 2) & "/" & Mid(strLinha, 6, 2) & "/" & Mid(strLinha, 2, 4)
            aLinha(UBound(aLinha)).sHoraGeracao = Mid(strLinha, 10, 2) & ":" & Mid(strLinha, 12, 2) & ":" & Mid(strLinha, 14, 2)
            GoTo PROXIMO2
        End If
        
        If Left(strLinha, 1) = 3 Then
            ReDim Preserve aLinha(UBound(aLinha) + 1)
            aLinha(UBound(aLinha)).nQtdeLinhas = Val(Mid(strLinha, 2, 6))
            aLinha(UBound(aLinha)).nSomaTotalGuias = CDbl(Mid(strLinha, 8, 13)) / 100
            aLinha(UBound(aLinha)).nQtdeSemMov = Val(Mid(strLinha, 21, 6))
            GoTo PROXIMO2
        End If
        If Val(Mid(strLinha, 7, 7)) = 0 Then
            GoTo PROXIMO2
        End If
        
'        If Mid(strLinha, 325, 6) = "Eliete" Then MsgBox "tyeste"
'If Val(Mid(strLinha, 7, 7)) = 3063081 Then MsgBox "teste"
        If Val(Mid(strLinha, 17, 1)) = 3 Then
            bTomador = True
        Else
            bTomador = False
        End If
        
        
        If Val(Mid(strLinha, 698, 1)) = 1 And Val(Mid(strLinha, 17, 1)) = 4 Then
            bTomador = True 'guias avulsas olhar para coluna 699
        End If
        
        bAchou = False
        If UBound(aLinha) = 156 Then
           ' MsgBox "teste"
        End If
        
        If Left(strLinha, 1) = 2 Then 'inicio detalhe
            On Error Resume Next
            For k = 1 To UBound(aLinha)
                If aLinha(k).nNumeroDaGuia = Val(Mid(strLinha, 7, 7)) Then
                    bAchou = True
                    Exit For
                End If
            Next
            On Error GoTo 0
            If Not bAchou Then
                ReDim Preserve aLinha(UBound(aLinha) + 1)
                aLinha(UBound(aLinha)).nNumeroDaGuia = Mid(strLinha, 7, 7)
'                If aLinha(UBound(aLinha)).nNumeroDaGuia = 3062465 Then
'                    MsgBox "teste"
'                End If
                aLinha(UBound(aLinha)).nSequencia = Val(Mid(strLinha, 14, 3))
                aLinha(UBound(aLinha)).nTipoDeEmissao = Val(Mid(strLinha, 17, 1))
                aLinha(UBound(aLinha)).sSimplesNacional = Mid(strLinha, 18, 1)
                aLinha(UBound(aLinha)).nExercicio = Val(Mid(strLinha, 19, 4))
                aLinha(UBound(aLinha)).nMes = Val(Mid(strLinha, 23, 2))
                aLinha(UBound(aLinha)).nAliquota = CDbl(Mid(strLinha, 25, 5)) / 100
                aLinha(UBound(aLinha)).sInscricaoT = Mid(strLinha, 30, 15)
                aLinha(UBound(aLinha)).sRazaoSocialT = Mid(strLinha, 45, 100)
                aLinha(UBound(aLinha)).sCPFT = Mid(strLinha, 145, 15)
                aLinha(UBound(aLinha)).sCNPJT = Mid(strLinha, 161, 14)
                aLinha(UBound(aLinha)).sEnderecoT = Mid(strLinha, 175, 100)
                aLinha(UBound(aLinha)).sComplEndT = Mid(strLinha, 275, 35)
                aLinha(UBound(aLinha)).sInscricaoP = Mid(strLinha, 310, 15)
                aLinha(UBound(aLinha)).sRazaoSocialP = Mid(strLinha, 325, 100)
                aLinha(UBound(aLinha)).sCPFP = Mid(strLinha, 425, 15)
                aLinha(UBound(aLinha)).sCNPJP = Mid(strLinha, 441, 14)
                aLinha(UBound(aLinha)).sEnderecoP = Mid(strLinha, 455, 100)
                aLinha(UBound(aLinha)).sComplEndP = Mid(strLinha, 555, 35)
                aLinha(UBound(aLinha)).nValorMovimento = CDbl(Mid(strLinha, 590, 13)) / 100
                aLinha(UBound(aLinha)).nValorImposto = CDbl(Mid(strLinha, 603, 13)) / 100
                aLinha(UBound(aLinha)).nValorMulta = CDbl(Mid(strLinha, 616, 13)) / 100
                aLinha(UBound(aLinha)).nValorJuros = CDbl(Val(Mid(strLinha, 629, 13))) / 100
                aLinha(UBound(aLinha)).nValorCorrecao = CDbl(Mid(strLinha, 642, 13)) / 100
                aLinha(UBound(aLinha)).sDataEmissao = Mid(strLinha, 661, 2) & "/" & Mid(strLinha, 659, 2) & "/" & Mid(strLinha, 655, 4)
                aLinha(UBound(aLinha)).sDataVencimento = Mid(strLinha, 669, 2) & "/" & Mid(strLinha, 667, 2) & "/" & Mid(strLinha, 663, 4)
                aLinha(UBound(aLinha)).sUsuario = Mid(strLinha, 671, 16)
                aLinha(UBound(aLinha)).sAtivIss = Mid(strLinha, 687, 8)
                aLinha(UBound(aLinha)).sAtivSeq = Mid(strLinha, 695, 3)
                aLinha(UBound(aLinha)).nStatus = Val(Mid(strLinha, 698, 1))
                aLinha(UBound(aLinha)).nER = Val(Mid(strLinha, 699, 1))
                
                If bTomador Then
                    sInscricao = aLinha(UBound(aLinha)).sInscricaoT
                    sCPF = aLinha(UBound(aLinha)).sCPFT
                    sCNPJ = aLinha(UBound(aLinha)).sCNPJT
                    If Val(sCNPJ) = 0 Then
                        sCNPJ = sInscricao
                    End If
                    If Val(sCPF) = 0 Then
                        sCPF = Right(sInscricao, 11)
                    End If
                    sRazao = aLinha(UBound(aLinha)).sRazaoSocialT
                Else
                    sInscricao = aLinha(UBound(aLinha)).sInscricaoP
                    sCPF = Val(aLinha(UBound(aLinha)).sCPFP)
                    sCNPJ = aLinha(UBound(aLinha)).sCNPJP
                    If Val(sCNPJ) = 0 Then
                        sCNPJ = sInscricao
                    End If
                    If Val(sCPF) = 0 Then
                        sCPF = Right(sInscricao, 11)
                    End If
                    sRazao = aLinha(UBound(aLinha)).sRazaoSocialP
                End If
                
                '*****************************************************************************
                '*********** IDENTIFICAÇÃO DO CNPJ, CPF OU CÓDIGO ****************************
                '*****************************************************************************
            
                If Val(sInscricao) > 0 Then
'                    If Val(sInscricao > 999999) Then
'                        MsgBox "teste"
'                    End If
                
                
                    Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODIGOMOB=" & RetornaNumero(sInscricao)
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        If .RowCount > 0 Then
                            nCodReduz = !codigomob
                        Else
'                            nCodReduz = 0
'                            MsgBox "Código não encontrado, empresa: " & sRazao & " doc: " & sCNPJ, vbExclamation, "Atenção"
'                        Else
                            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & RetornaNumero(sInscricao)
                            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux
                                If .RowCount > 0 Then
                                    nCodReduz = !CodCidadao
                                Else
                                    If ValidaCGC(Right(sCNPJ, 14)) Then
                                        Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & Val(sCNPJ) & "' or CNPJ='" & sCNPJ & "'"
                                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                        With RdoAux
                                            If .RowCount > 0 Then
                                                nCodReduz = !CodCidadao
                                            Else
                                                'NAO ACHOU, CADASTRA UM NOVO
                                                Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
                                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                                With RdoAux2
                                                    nCodReduz = !maximo + 1
                                                   .Close
                                                End With

                                                Sql = "INSERT CIDADAO (CODCIDADAO,NOMECIDADAO,CNPJ) VALUES(" & nCodReduz & ",'"
                                                Sql = Sql & UCase(Left(Trim(sRazao), 50)) & "','" & sCNPJ & "')"
                                           '     cn.Execute Sql, rdExecDirect
                                            End If
                                           .Close
                                        End With
                                    Else
                                       'CNPJ INVALIDO
                                       'TESTA CPF
                                       If Val(Right(sCPF, 11)) = 0 Then
                                            nCodReduz = 0
                                            GoTo FIMCPF
                                       End If
                                        If ValidaCPF(Right(sCPF, 11)) Then
                                            Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CPF='" & sCPF & "'"
                                            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                            With RdoAux
                                                If .RowCount > 0 Then
                                                    nCodReduz = !CodCidadao
                                                Else
                                                    'NAO ACHOU, CADASTRA UM NOVO
                                                    Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO WHERE CODCIDADAO<700000"
                                                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                                    With RdoAux2
                                                        nCodReduz = !maximo + 1
                                                       .Close
                                                    End With
                                                    Sql = "INSERT CIDADAO (CODCIDADAO,NOMECIDADAO,CPF) VALUES(" & nCodReduz & ",'"
                                                    Sql = Sql & Left(sRazao, 50) & "','" & Val(sCPF) & "')"
                                        '            cn.Execute Sql, rdExecDirect
                                                End If
                                               .Close
                                            End With
                                        Else
                                            'CPF INVALIDO
                                            nCodReduz = 0
                                        End If
FIMCPF:
                                    End If
                                End If
                               .Close
                            End With
                        End If
                       .Close
                    End With
                    
                    If nCodReduz > 0 Then
                        aLinha(UBound(aLinha)).nCodReduz = nCodReduz
                    Else
                        GoTo PROXIMO2
                    End If
                End If 'fim da identificação do prestador
        Else
            aLinha(k).nValorMovimento = aLinha(k).nValorMovimento + CDbl(Mid(strLinha, 590, 13)) / 100
            aLinha(k).nValorImposto = aLinha(k).nValorImposto + CDbl(Mid(strLinha, 603, 13)) / 100
            aLinha(k).nValorMulta = aLinha(k).nValorMulta + CDbl(Mid(strLinha, 616, 13)) / 100
            aLinha(k).nValorJuros = aLinha(k).nValorJuros + CDbl(Mid(strLinha, 629, 13)) / 100
            aLinha(k).nValorCorrecao = aLinha(k).nValorCorrecao + CDbl(Mid(strLinha, 642, 13)) / 100
            End If 'Fim do detalhe
        End If

        DoEvents
        
PROXIMO2:
   Loop
Close #1
'Exit Sub
'Open sPathBin & "\ERROGUIA.TXT" For Output As #8
Pb.value = 0
For x = 1 To UBound(aLinha)
    CallPb CLng(x), CLng(UBound(aLinha))
    'If x = 3463 Then MsgBox "teste"
    DoEvents
    With aLinha(x)
'        If .nNumeroDaGuia = 3060691 Then
'            MsgBox "teste"
'        End If
        
        If Val(.nExercicio) = 0 Then
            GoTo proximo
        End If
        
        
        If .nCodReduz = 0 Then
            GoTo proximo
        End If
        nCodReduz = .nCodReduz
        nDoc = .nNumeroDaGuia
        
        
        
        bErro = False
'        If nDoc = 3063140 Then MsgBox "teste"
        GoTo continua
        Sql = "SELECT NUMERO,INSCRICAO FROM GUIAISSELETRONICO WHERE NUMERO=" & nDoc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        With RdoAux
            If .RowCount > 0 Then
 '               If nCodReduz <> !Inscricao Then
 '                   Print #8, nDoc & "," & nCodReduz & "," & !Inscricao

'                    MsgBox "TESTE"
  '              End If
                .Close
                GoTo proximo
                'CORREÇÃO DO LAYOUT ERRADO DA CONSIST -- ROTINA TEMPORÁRIA
                Sql = "SELECT NUMERO,INSCRICAO FROM GUIAISSELETRONICO WHERE NUMERO=" & nDoc & " AND TIPO=3 AND (MONTH(datavencto) = 5 or MONTH(datavencto) = 6 or MONTH(datavencto) = 7) AND (YEAR(datavencto) = 2010)"
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                If RdoAux.RowCount = 0 Then
                    GoTo proximo
                Else
                    
                    bErro = True
                End If
                nCodReduzErrado = RdoAux!Inscricao
                
                '1 - ATUALIZAMOS A TABELA GUIAISSELETRONICO COM O NOVO CODIGO REDUZIDO E VALOR
                Sql = "UPDATE GUIAISSELETRONICO SET INSCRICAO=" & nCodReduz & ",VALORPRINCIPAL=" & Virg2Ponto(CStr(aLinha(x).nValorImposto)) & ","
                Sql = Sql & "VALORACRESCIMO=" & Virg2Ponto(CStr(aLinha(x).nValorJuros + aLinha(x).nValorMulta + aLinha(x).nValorCorrecao)) & " WHERE NUMERO=" & nDoc
                cn.Execute Sql, rdExecDirect
                
                '2 - PEGAMOS OS DADOS DA PARCELA ERRADA
                Sql = "SELECT * FROM PARCELADOCUMENTO WHERE NUMDOCUMENTO=" & nDoc
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        nAno = !AnoExercicio
                        nLanc = !CodLancamento
                        nSeq = !SeqLancamento
                        nParc = !NumParcela
                        nCompl = !CODCOMPLEMENTO
                       .Close
                        
                        '3 - APAGAMOS O LANCAMENTO EM PARCELADOCUMENTO
                        Sql = "DELETE FROM PARCELADOCUMENTO WHERE NUMDOCUMENTO=" & nDoc
                        cn.Execute Sql, rdExecDirect
                    
                        '4 - CANCELAMOS O LANCAMENTO EM DEBITOPARCELA
                        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=8 WHERE CODREDUZIDO=" & nCodReduzErrado & " AND ANOEXERCICIO=" & nAno & " AND "
                        Sql = Sql & "CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND CODCOMPLEMENTO=" & nCompl
                        cn.Execute Sql, rdExecDirect
                        
                        '5 - EXPLICAMOS PQ EM OBS DA PARCELA
                        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduzErrado & " AND ANOEXERCICIO=" & nAno
                        Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
                        Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            If IsNull(!maximo) Then
                                nSeq2 = 1
                            Else
                                nSeq2 = !maximo + 1
                            End If
                           .Close
                        End With
                        sObs = "Lancamento cancelado automaticamente pelo sistema de correção do GTI pelo motivo de ter sido criado errôneamente durante a integração com o novo sistema de ISS eletrônico."
'                        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduzErrado & "," & nAno & ","
'                        Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nSeq2 & ",'" & sObs & "','GTI','" & Format(Now, "mm/dd/yyyy") & "')"
                        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduzErrado & "," & nAno & ","
                        Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nSeq2 & ",'" & sObs & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
                        cn.Execute Sql, rdExecDirect
                        
                        GoTo continua
                    Else
                       .Close
                    End If
                End With
            End If
'           .Close
        End With
continua:
        sDataVencto = .sDataVencimento
            
        If .nMes <> 12 Then
            sDataVencto = "15" & "/" & CStr(.nMes + 1) & "/" & .nExercicio
        Else
            sDataVencto = "15" & "/" & "1" & "/" & .nExercicio + 1
        End If

'        If Not bErro Then
            Sql = "SELECT NUMERO,INSCRICAO FROM GUIAISSELETRONICO WHERE NUMERO=" & nDoc
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
            If RdoAux.RowCount > 0 Then GoTo proximo
'            Sql = "DELETE FROM guiaisseletronico where numero=" & nDoc
'            cn.Execute Sql, rdExecDirect
            
            Sql = "INSERT guiaisseletronico(numero,inscricao,sequencia,ano,mes,aliquota,tipo,datavencto,valorprincipal,valoracrescimo,"
            Sql = Sql & "dataexportacao,identificacao) values(" & .nNumeroDaGuia & "," & nCodReduz & "," & .nSequencia & "," & .nExercicio & ","
            Sql = Sql & .nMes & "," & Virg2Ponto(CStr(.nAliquota)) & "," & Val(.nTipoDeEmissao) & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "',"
            Sql = Sql & Virg2Ponto(CStr(.nValorImposto)) & "," & Virg2Ponto(CStr(.nValorJuros + .nValorMulta + .nValorCorrecao)) & ",'" & Format(.sDataEmissao, "mm/dd/yyyy") & "',"
            Sql = Sql & 0 & ")"
            cn.Execute Sql, rdExecDirect
 '       End If
        
       
        'CRIA DÉBITO DE ISS VARIAVEL NO GTI

        'BUSCAR A ÚLTIMA SEQUENCIA DE LANCAMENTO PARA EVITAR DUPLICIDADE
        Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE (codreduzido = " & nCodReduz & ") AND ANOEXERCICIO=" & .nExercicio & " And (CodLancamento = 5) "
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If IsNull(!maximo) Then
                nSeq = 0
            Else
                nSeq = !maximo + 1
            End If
           .Close
        End With

        'CRIAR PARCELA DE ISS VARIAVEL NESTE MES E ANO COM O VENCIMENTO QUE VEIO DO BANCO
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodReduz & "," & .nExercicio & "," & 5 & "," & nSeq & ","
'        Sql = Sql & 1 & "," & nCompl & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,'GTI')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodReduz & "," & .nExercicio & "," & 5 & "," & nSeq & ","
        Sql = Sql & 1 & "," & nCompl & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy") & "',0,236)"
        cn.Execute Sql, rdExecDirect
        'CRIAR O TRIBUTO PARA ELA (13 - iss variavel)
        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
        Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodReduz & "," & .nExercicio & "," & 5 & "," & nSeq & ","
        Sql = Sql & 1 & "," & nCompl & "," & 13 & "," & Virg2Ponto(CStr(.nValorImposto)) & ")"
        cn.Execute Sql, rdExecDirect
        'CRIAR O DOCUMENTO PARA ELA
        On Error Resume Next
     '   If Not bErro Then
     '       Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA) VALUES(" & .nNumeroDaGuia & ",'"
     '       Sql = Sql & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & ")"
     '       cn.Execute Sql, rdExecDirect
     '   End If
        On Error GoTo 0
        'CRIAR A PARCELADOCUMENTO
      '  Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & .nExercicio & "," & 5 & "," & nSeq & ","
      '  Sql = Sql & 1 & "," & nCompl & "," & .nNumeroDaGuia & ")"
      '  cn.Execute Sql, rdExecDirect

    End With

proximo:
Next

'VERIFICA GUIAS CANCELADAS
Pb.value = 0
For x = 1 To UBound(aLinha)
    CallPb CLng(x), CLng(UBound(aLinha))
    DoEvents
    With aLinha(x)
 '       nDoc = .nNumeroDaGuia
        If Val(.nStatus) = 9 Then
            nDoc = .nNumeroDaGuia
'            If nDoc = 3062465 Then MsgBox "TESTE"
            Sql = "SELECT parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, parceladocumento.seqlancamento, "
            Sql = Sql & "parceladocumento.NumParcela , parceladocumento.CODCOMPLEMENTO, parceladocumento.NumDocumento, debitoparcela.statuslanc FROM parceladocumento INNER JOIN "
            Sql = Sql & "debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND parceladocumento.anoexercicio = debitoparcela.anoexercicio AND "
            Sql = Sql & "parceladocumento.codlancamento = debitoparcela.codlancamento AND parceladocumento.seqlancamento = debitoparcela.seqlancamento AND "
            Sql = Sql & "parceladocumento.NumParcela = debitoparcela.NumParcela And parceladocumento.CODCOMPLEMENTO = debitoparcela.CODCOMPLEMENTO "
            Sql = Sql & "Where parceladocumento.NumDocumento = " & nDoc
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    If !statuslanc = 3 Or !statuslanc = 25 Or !statuslanc = 20 Or !statuslanc = 19 Then
                        Sql = "update debitoparcela set statuslanc=5 where codreduzido=" & !CODREDUZIDO & " and anoexercicio=" & !AnoExercicio & " and codlancamento=" & !CodLancamento & " and "
                        Sql = Sql & "seqlancamento=" & !SeqLancamento & " and numparcela=" & !NumParcela & " and codcomplemento=" & !CODCOMPLEMENTO
                        cn.Execute Sql, rdExecDirect
                    
                        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio
                        Sql = Sql & " AND CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela
                        Sql = Sql & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO
                        Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux3
                            If IsNull(!maximo) Then
                                nSeq2 = 1
                            Else
                                nSeq2 = !maximo + 1
                            End If
                           .Close
                        End With
                        sObs = "Lancamento cancelado automaticamente pelo sistema GTI de acordo com a importação de Guias de ISS eletrônico."
'                        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
'                        Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeq2 & ",'" & sObs & "','GTI','" & Format(Now, "mm/dd/yyyy") & "')"
                        Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & !CODREDUZIDO & "," & !AnoExercicio & ","
                        Sql = Sql & !CodLancamento & "," & !SeqLancamento & "," & !NumParcela & "," & !CODCOMPLEMENTO & "," & nSeq2 & ",'" & sObs & "',236,'" & Format(Now, "mm/dd/yyyy") & "')"
                        cn.Execute Sql, rdExecDirect
                    
                    
                    End If
                End If
            End With
        End If
    End With
Next

'Close #8
Liberado
Pb.value = 0
MsgBox "Arquivo importado com sucesso.", vbInformation, "Informação"

End Sub

Private Function SNCheck(nCodigo As Long) As Integer
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNASN(" & Format(nCodigo, "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
     If RdoAux!RETORNO = 1 Then
        SNCheck = 1
     Else
        SNCheck = 0
     End If
    .Close
End With

End Function

Private Sub ExportaConsistOld()
Dim ax As String, aCodigos() As Long, x As Integer, nPos As Long, nCPF As Byte, sDoc As String, sDataEncerra As String, RdoAux2 As rdoResultset
Dim sEnd As String, sCep As String, aAtiv(24) As ATIVIDADES, t As Integer, sAtiv As String, aAliq(24) As Double, bAchou As Boolean, nSimples As Single, nRegEspecial As Single

Open sPathBin & "\PREFJABC.TXT" For Output As #1

Print #1, "[INICIO-ANO]"
Print #1, Format(Year(Now), "0000")
Print #1, "[FIM-ANO]"

Print #1, ""
Print #1, "[INICIO-SELIC]"
Print #1, "[FIM-SELIC]"

Print #1, ""
Print #1, "[INICIO-ATIVIDADE]"
Sql = "SELECT DISTINCT ATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE, TABELAISS.TIPOISS,TABELAISS.ALIQUOTA * 100 as aliquota,TABELAISS.DATA "
'Sql = Sql & "FROM ATIVIDADEISS INNER JOIN TABELAISS ON ATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV where tipoiss<>11 ORDER BY codatividade,DATA"
Sql = Sql & "FROM ATIVIDADEISS INNER JOIN TABELAISS ON ATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV ORDER BY codatividade,DATA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        If !TipoISS <> 11 Then
            ax = Format(!codatividade, "00000000") & FillSpace(!descatividade, 300) & Format(!Aliquota * 100, "0000") & Format(Day(!Data), "00") & Format(Month(!Data), "00") & Year(!Data)
        Else
            ax = Format(!codatividade, "00000000") & FillSpace(!descatividade, 300) & Format(0, "0000") & Format(Day(!Data), "00") & Format(Month(!Data), "00") & Year(!Data)
        End If
        Print #1, ax
       .MoveNext
    Loop
   .Close
End With
Print #1, "[FIM-ATIVIDADE]"
 
Print #1, ""
Print #1, "[INICIO-EMPRESA]"
ReDim aCodigos(0)
'CARREGA APENAS AS EMPRESAS VARIAVEL E ESTIMADO
'Sql = "SELECT DISTINCT codmobiliario From mobiliarioatividadeiss Where (CodTributo <> 11) And (codmobiliario > 100000)  ORDER BY CODMOBILIARIO"
Sql = "SELECT DISTINCT CODIGOMOB FROM VWCNSMOBILIARIO WHERE (DATAENCERRAMENTO > '10/1/2007'or dataencerramento is null) ORDER BY CODIGOMOB"
'Sql = "SELECT DISTINCT CODIGOMOB FROM VWCNSMOBILIARIO WHERE (DATAENCERRAMENTO > '10/1/2007'or dataencerramento is null) AND CODIGOMOB=116534 ORDER BY CODIGOMOB"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCodigos(UBound(aCodigos) + 1)
        aCodigos(UBound(aCodigos)) = !codigomob
       .MoveNext
    Loop
   .Close
End With

'CARREGA DADOS DAS EMPRESAS
nPos = 0: Pb.value = 0
nTot = UBound(aCodigos): lblRegTot.Caption = nTot
For x = 1 To UBound(aCodigos)
    nPos = x
    CallPb nPos, CLng(nTot)
    lblRegPerc.Caption = nPos
'   Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & aCodigos(x) & " and  DATAENCERRAMENTO > '10/1/2007' "
'If aCodigos(x) = 116528 Then MsgBox "teste"

    Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & aCodigos(x) & " and (DATAENCERRAMENTO > '10/1/2007'or dataencerramento is null)"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    With RdoAux
        If .RowCount = 0 Then GoTo proximo
        
        nSimples = SNCheck(aCodigos(x))
        nRegEspecial = Val(SubNull(!REGESPECIAL))
        nCPF = IIf(IsNull(!Cnpj), 1, 0)
        sDoc = IIf(IsNull(!CPF), SubNull(!Cnpj), SubNull(!CPF))
        sDataEncerra = "        "
        If Not IsNull(!dataencerramento) Then
            sDataEncerra = Format(!dataencerramento, "yyyymmdd")
        End If
       'SUSPENÇÃO
        Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & !codigomob & " ORDER BY DATAEVENTO DESC"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If !CODTIPOEVENTO = 2 Then
                    sDataEncerra = Format(!DATAPROCEVENTO, "yyyymmdd")
                End If
            End If
           .Close
        End With
        
        If !CodCidade = 413 Then
            sCep = RetornaCEP(!CodLogradouro, !Numero)
        Else
            sCep = Left$(!Cep, 5) & "-" & Right$(!Cep, 3)
        End If
        
       'ENDERECO
        sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & ", " & !Numero & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
        sEnd = sEnd & !DescBairro & " - " & !descCidade & " - " & !SiglaUF & " - " & sCep
        sEnd = Left$(sEnd, 200)
        ax = Format(!codigomob, "00000000") & CStr(nCPF) & Format(Val(sDoc), "00000000000000") & FillSpace(!RazaoSocial, 100) & FillSpace(sDataEncerra, 8) & FillSpace(sEnd, 200)
       .Close
    End With
   'ATIVIDADES
    For t = 1 To 24
        aAtiv(t).nCodigo = 0
        aAtiv(t).nSeq = 0
        aAliq(t) = 0
    Next
   'CARREGA ATIVIDADE DISTINTAS
    'Sql = "SELECT DISTINCT VALORISS FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & aCodigos(x) & " AND CODTRIBUTO<>11"
    Sql = "SELECT DISTINCT CODTRIBUTO,VALORISS FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & aCodigos(x)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
    With RdoAux
        Do Until .EOF
            If !CodTributo = 11 Then
                aAliq(.AbsolutePosition) = 0
            Else
                aAliq(.AbsolutePosition) = !valoriss
            End If
           .MoveNext
        Loop
       .Close
    End With
    'CARREGA AS ATIVIDADES E SUA SEQUENCIA
    'Sql = "SELECT CODATIVIDADE,VALORISS,CODTRIBUTO FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & aCodigos(x) & " AND CODTRIBUTO<>11"
    Sql = "SELECT CODATIVIDADE,VALORISS,CODTRIBUTO FROM MOBILIARIOATIVIDADEISS WHERE CODMOBILIARIO=" & aCodigos(x)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
    With RdoAux
        Do Until .EOF
            For t = 1 To UBound(aAliq)
                If aAliq(t) = !valoriss Then
                    Exit For
                End If
            Next
            aAtiv(.AbsolutePosition).nCodigo = !codatividade
            aAtiv(.AbsolutePosition).nSeq = t
            aAtiv(.AbsolutePosition).sEstimado = IIf(!CodTributo = 12, "X", " ")
           .MoveNext
        Loop
       .Close
    End With
    
    sAtiv = ""
    For t = 1 To 24
        sAtiv = sAtiv & Format(aAtiv(t).nCodigo, "00000000")
    Next
    For t = 1 To 24
        sAtiv = sAtiv & Format(aAtiv(t).nSeq, "00")
    Next
    For t = 1 To 24
        sAtiv = sAtiv & IIf(aAtiv(t).sEstimado = "", " ", aAtiv(t).sEstimado)
    Next
    ax = ax & sAtiv
    ax = ax & IIf(nSimples = 1, "X", "")
    ax = ax & IIf(nRegEspecial = 1, "X", "")
    Print #1, ax
proximo:
Next

'******* empresa retensao iss

'CARREGA DADOS DAS EMPRESAS COM RETENÇÃO NA FONTE
Sql = "SELECT CODREDUZIDO FROM ISSRETIDO ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 0: Pb.value = 0
    nTot = .RowCount: lblRegTot.Caption = nTot
    Do Until .EOF
        nPos = .AbsolutePosition
        CallPb nPos, CLng(nTot)
        lblRegPerc.Caption = nPos
        Sql = "SELECT * FROM VWCIDADAO WHERE CODCIDADAO=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        With RdoAux2
            If .RowCount > 0 Then
                nCPF = IIf(IsNull(!Cnpj), 1, 0)
                sDoc = IIf(IsNull(!CPF) Or Trim(!CPF) = "", Trim(SubNull(!Cnpj)), Trim(SubNull(!CPF)))
                sCep = ""
                If !CodCidade = 413 Then
                    If Not IsNull(!NUMIMOVEL) And !CodLogradouro > 0 Then
                        sCep = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
                    ElseIf Not IsNull(!NUMIMOVEL) And Not IsNull(!Cep) Then
                        sCep = !Cep
                    End If
                Else
                    If Not IsNull(!Cep) Then
                        sCep = Left$(!Cep, 5) & "-" & Right$(!Cep, 3)
                    End If
                End If
                If Not IsNull(!NomeLogradouro) Then
                   If !NomeLogradouro <> "" Then
                       sEnd = Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL)) & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
                   Else
                       sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
                   End If
                Else
                   sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & IIf(!Complemento = "", "", " " & !Complemento) & " - bairro: "
                End If
'                If Not IsNull(!NOMEBairro) Then
'                   sEnd = sEnd & SubNull(!NOMEBairro) & " - " & SubNull(!NomeCidade) & " - " & !NOMEUF & " - " & sCEP
'                Else
                   sEnd = sEnd & SubNull(!DescBairro) & " - " & SubNull(!descCidade) & " - " & !SiglaUF & " - " & sCep
 '               End If
                sEnd = Left$(sEnd, 200)
                
                ax = Format(!CodCidadao, "00000000") & CStr(nCPF) & Format(Val(sDoc), "00000000000000") & FillSpace(!nomecidadao, 100) & FillSpace("", 8) & FillSpace(sEnd, 200)
                Print #1, ax
            End If
           .Close
        End With
PROXIMO2:
       .MoveNext
    Loop
End With

Print #1, "[FIM-EMPRESA]"
Print #1, ""
Print #1, "[INICIO-VENCIMENTOS]"
Print #1, "200701102007021220070312200704102007051020070611200707102007081020070910200710102007111620071217"
Print #1, "200801152008021520080317200804152008051520080616200807152008081520080915200810152008111720081215"
Print #1, "200901152009021620090316200904152009051520090615200907152009081720090915200910152009111620091215"
Print #1, "201001152010021720100315201004152010051720100615201007152010081620100915201010152010111520101215"
Print #1, "[FIM-VENCIMENTOS]"

Print #1, ""
Print #1, "[INICIO-GUIAS-PAGAS]"
Sql = "SELECT guiaisseletronico.numero, guiaisseletronico.datavencto, guiaisseletronico.valorprincipal, guiaisseletronico.valoracrescimo, "
Sql = Sql & "debitopago.DataPagamento , NumDocumento.ValorPago FROM numdocumento INNER JOIN guiaisseletronico ON numdocumento.numdocumento = guiaisseletronico.numero LEFT OUTER JOIN "
Sql = Sql & "debitopago ON guiaisseletronico.numero = debitopago.numdocumento Where (NumDocumento.ValorPago > 0) ORDER BY guiaisseletronico.numero"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ax = Format(!Numero, "000000000000")
        If Not IsNull(!DataPagamento) Then
            ax = ax & Format(Year(!DataPagamento), "0000") & Format(Month(!DataPagamento), "00") & Format(Day(!DataPagamento), "00")
        Else
            ax = ax & Format(Year(!DataVencto), "0000") & Format(Month(!DataVencto), "00") & Format(Day(!DataVencto), "00")
        End If
        ax = ax & Format(!ValorPrincipal, "000000000.00")
        ax = ax & Format(!ValorAcrescimo, "000000000.00")
        Print #1, ax
       .MoveNext
    Loop
   .Close
End With

Print #1, "[FIM-GUIAS-PAGAS]"

Print #1, ""
Print #1, "[INICIO-FERIADOS]"
Print #1, "[FIM-FERIADOS]"

Close #1
Liberado
MsgBox "Exportação finalizada com sucesso.", vbInformation, "Operação finalizada"
End Sub

