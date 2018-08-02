VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmGeracaoDebito 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geração de débitos ao contribuinte"
   ClientHeight    =   6300
   ClientLeft      =   2145
   ClientTop       =   3015
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6300
   ScaleWidth      =   11505
   Begin VB.TextBox txtLanc 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2820
      MaxLength       =   4
      TabIndex        =   48
      Top             =   780
      Width           =   3675
   End
   Begin VB.Frame pnlEnd 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Endereços do contribuinte"
      Height          =   2025
      Left            =   3000
      TabIndex        =   19
      Top             =   1350
      Visible         =   0   'False
      Width           =   8325
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço.......:"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   43
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblRuaEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1290
         TabIndex        =   42
         Top             =   1080
         Width           =   4860
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Index           =   3
         Left            =   6345
         TabIndex        =   41
         Top             =   1080
         Width           =   405
      End
      Begin VB.Label lblNumEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6765
         TabIndex        =   40
         Top             =   1065
         Width           =   705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.:"
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   39
         Top             =   1365
         Width           =   1155
      End
      Begin VB.Label lblComplentrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1290
         TabIndex        =   38
         Top             =   1365
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...:"
         Height          =   225
         Index           =   5
         Left            =   4305
         TabIndex        =   37
         Top             =   1380
         Width           =   690
      End
      Begin VB.Label lblBairroEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5010
         TabIndex        =   36
         Top             =   1365
         Width           =   2460
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade...........:"
         Height          =   225
         Index           =   7
         Left            =   60
         TabIndex        =   35
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label lblCidadeEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1290
         TabIndex        =   34
         Top             =   1650
         Width           =   2730
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep......:"
         Height          =   225
         Index           =   8
         Left            =   5340
         TabIndex        =   33
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label lblCepEntrega 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6030
         TabIndex        =   32
         Top             =   1665
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UF...:"
         Height          =   225
         Index           =   12
         Left            =   4290
         TabIndex        =   31
         Top             =   1680
         Width           =   390
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...................:"
         Height          =   225
         Index           =   6
         Left            =   105
         TabIndex        =   30
         Top             =   270
         Width           =   1695
      End
      Begin VB.Label lblRua 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1755
         TabIndex        =   29
         Top             =   300
         Width           =   3690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Index           =   1
         Left            =   6105
         TabIndex        =   28
         Top             =   315
         Width           =   405
      End
      Begin VB.Label lblNumImovel 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6525
         TabIndex        =   27
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label lblCep 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6525
         TabIndex        =   26
         Top             =   585
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep.:"
         Height          =   225
         Index           =   9
         Left            =   6105
         TabIndex        =   25
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lblCompl 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1755
         TabIndex        =   24
         Top             =   585
         Width           =   1680
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento.............:"
         Height          =   225
         Index           =   10
         Left            =   90
         TabIndex        =   23
         Top             =   570
         Width           =   1740
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro..:"
         Height          =   225
         Index           =   11
         Left            =   3495
         TabIndex        =   22
         Top             =   585
         Width           =   570
      End
      Begin VB.Label lblBairro 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4110
         TabIndex        =   21
         Top             =   600
         Width           =   1845
      End
      Begin VB.Label lbluf 
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   4710
         TabIndex        =   20
         Top             =   1680
         Width           =   555
      End
   End
   Begin VB.TextBox txtProp 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   5970
      Locked          =   -1  'True
      MaxLength       =   200
      TabIndex        =   18
      Top             =   60
      Width           =   5445
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5565
      MaxLength       =   4
      TabIndex        =   14
      Top             =   450
      Width           =   915
   End
   Begin VB.CheckBox chkUnica 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Emitir Parcela Ùnica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1980
      TabIndex        =   13
      Top             =   480
      Width           =   2070
   End
   Begin VB.TextBox txtNumParc 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   990
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   12
      Top             =   420
      Width           =   945
   End
   Begin VB.Frame fr4 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Atividade ISS"
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   0
      TabIndex        =   10
      Top             =   4320
      Width           =   3810
      Begin MSComctlLib.ListView lvISS 
         Height          =   1230
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1236
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   10233
         EndProperty
      End
   End
   Begin VB.Frame fr5 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Atividade Tx.Licença"
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   3795
      TabIndex        =   8
      Top             =   4260
      Width           =   3810
      Begin MSComctlLib.ListView lvTL 
         Height          =   1230
         Left            =   60
         TabIndex        =   9
         Top             =   225
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   1412
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrição"
            Object.Width           =   9878
         EndProperty
      End
   End
   Begin VB.Frame fr6 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Atividade Vig.Sanitária"
      ForeColor       =   &H00800000&
      Height          =   1530
      Left            =   7620
      TabIndex        =   6
      Top             =   4260
      Width           =   3810
      Begin MSComctlLib.ListView lvVS 
         Height          =   1230
         Left            =   60
         TabIndex        =   7
         Top             =   225
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   2170
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cód1"
            Object.Width           =   1236
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cód2"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Descrição"
            Object.Width           =   9878
         EndProperty
      End
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   990
      MaxLength       =   6
      TabIndex        =   0
      Top             =   90
      Width           =   945
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   1980
      TabIndex        =   1
      ToolTipText     =   "Consulta Imóvel"
      Top             =   60
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
      MICON           =   "frmGeracaoDebito.frx":0000
      PICN            =   "frmGeracaoDebito.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEnd 
      Height          =   345
      Left            =   150
      TabIndex        =   44
      ToolTipText     =   "Exibe os endereços do contribuinte"
      Top             =   5880
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Endereços"
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
      MICON           =   "frmGeracaoDebito.frx":0176
      PICN            =   "frmGeracaoDebito.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataInicio 
      Height          =   285
      Left            =   10305
      TabIndex        =   45
      Top             =   450
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BackColor       =   15658734
      MouseIcon       =   "frmGeracaoDebito.frx":040B
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
      Locked          =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição do lançamento no boleto..:"
      Height          =   225
      Index           =   13
      Left            =   90
      TabIndex        =   47
      Top             =   825
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Efetuar o cálculo proporcional a partir da data.....:"
      Height          =   225
      Index           =   16
      Left            =   6720
      TabIndex        =   46
      Top             =   495
      Width           =   3555
   End
   Begin VB.Label lbltipoend 
      Height          =   225
      Left            =   4320
      TabIndex        =   17
      Top             =   5940
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano do Exercício:"
      Height          =   225
      Index           =   17
      Left            =   4185
      TabIndex        =   16
      Top             =   510
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parcelas....:"
      Height          =   225
      Index           =   14
      Left            =   90
      TabIndex        =   15
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label lblNumInsc 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2970
      TabIndex        =   5
      Top             =   90
      Width           =   2370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código/I.M.:"
      Height          =   225
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   105
      Width           =   975
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome.:"
      Height          =   225
      Left            =   5400
      TabIndex        =   3
      Top             =   105
      Width           =   615
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      Caption         =   "Inscr:"
      Height          =   225
      Left            =   2550
      TabIndex        =   2
      Top             =   90
      Width           =   480
   End
   Begin VB.Menu mnuTipo 
      Caption         =   "Tipo"
      Visible         =   0   'False
      Begin VB.Menu mnuMob 
         Caption         =   "Mobiliário"
      End
      Begin VB.Menu mnuImob 
         Caption         =   "Imobiliário"
      End
      Begin VB.Menu mnuOutros 
         Caption         =   "Cidadão"
      End
   End
End
Attribute VB_Name = "frmGeracaoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim sRet As String, xImovel As clsImovel

Private Sub cmdCnsImovel_Click()
PopupMenu mnuTipo
End Sub

Private Sub cmdEnd_Click()
pnlEnd.Visible = cmdEnd.Value
If pnlEnd.Visible Then pnlEnd.ZOrder 0
End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
     txtCod.text = Val(Left$(CodImovel, 7))
     CodImovel = 0
     txtCod_LostFocus
Else
    If Val(CodEmpresa) > 0 Then
         txtCod.text = Val(Left$(CodEmpresa, 7))
         CodEmpresa = 0
         txtCod_LostFocus
    Else
        If Val(CodCidadao) > 0 Then
             Unload frmCnsCidadao
             DoEvents
             txtCod.text = Val(CodCidadao)
             CodCidadao = 0
             txtCod_LostFocus
        End If
    End If
End If
End Sub

Private Sub Form_Load()

Ocupado
frmMdi.AddWindow Me.Name, Me.Caption
Set xImovel = New clsImovel
Centraliza Me
Liberado
sRet = RetEventUserForm(Me.Name)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub CarregaImovel(nCodigoImovel As Long)
Dim Sql As String, RdoAux As rdoResultset

Ocupado
With xImovel
    .CarregaImovel nCodigoImovel
    If .CodigoImovel > 0 Then
          lblNumInsc.Caption = .Inscricao
          txtProp.text = .NomePropPrincipal
          lblRua.Caption = Trim$(.AbrevTipoLog) & " " & Trim$(.AbrevTitLog) & " " & .NomeLogradouro
          lblNumImovel.Caption = .Li_Num
          lblCep.Caption = RetornaCEP(.CodLogr, .Li_Num)
          lblCompl.Caption = .Li_Compl
          lblBairro.Caption = .DescBairro
          Select Case .Ee_TipoEnd
                Case 0
                    lbltipoend.Caption = "(Endereço do Imóvel)"
                    lblRuaEntrega.Caption = lblRua.Caption
                    lblNumEntrega.Caption = lblNumImovel.Caption
                    lblComplentrega.Caption = lblCompl.Caption
                    lblBairroEntrega.Caption = lblBairro.Caption
                    lblCidadeEntrega.Caption = "JABOTICABAL"
                    lblCepEntrega.Caption = lblCep.Caption
                    lbluf.Caption = lbluf.Caption
                Case 1
                    lbltipoend.Caption = "(Endereço do Proprietário)"
                    CarregaEndCidadao .CodPropPrincipal
                Case 2
                    lbltipoend.Caption = "(Endereço de Entrega Específico)"
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
                    lbluf.Caption = .Ee_Uf
          End Select
    End If
End With

fim:
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
    lblComplentrega.Caption = SubNull(!COMPLEMENTO)
    lblBairroEntrega.Caption = SubNull(!DescBairro)
    lblCidadeEntrega.Caption = SubNull(!desccidade)
    lblCepEntrega.Caption = SubNull(!cep)
    lbluf.Caption = SubNull(!SiglaUF)
End With

End Sub

Private Sub CarregaLista()
Dim itmX As ListItem, z As Long

z = SendMessage(lvISS.hwnd, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvTL.hwnd, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvVS.hwnd, LVM_DELETEALLITEMS, 0, 0)


'VS
Sql = "SELECT MOBILIARIOATIVIDADEVS.CODVIGSANIT,MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT,"
Sql = Sql & "MOBILIARIOATIVIDADEVS.SEQ,MOBILIARIOATIVIDADEVS.QTDE,VIGSANITARIA.DESCVIGSANITARIA,VIGSANITARIA.VALORALIQ "
Sql = Sql & "FROM MOBILIARIOATIVIDADEVS INNER JOIN VIGSANITARIA ON MOBILIARIOATIVIDADEVS.CODVIGSANIT = VIGSANITARIA.CODVIGSANIT "
Sql = Sql & "AND MOBILIARIOATIVIDADEVS.SUBCODVIGSANIT = VIGSANITARIA.SUBCODVIGSANIT "
Sql = Sql & "Where MOBILIARIOATIVIDADEVS.CODMOBILIARIO = " & Val(txtCod.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If !SUBCODVIGSANIT > 0 Then
           Sql = "SELECT DESCVIGSANITARIA FROM VIGSANITARIA WHERE CODVIGSANIT=" & !CODVIGSANIT & " AND SUBCODVIGSANIT=0"
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           Set itmX = lvVS.ListItems.Add(, "VS" & Format(!CODVIGSANIT, "000") & Format(!SUBCODVIGSANIT, "00"), Format(!CODVIGSANIT, "000"))
           itmX.SubItems(1) = Format(!SUBCODVIGSANIT, "00")
           itmX.SubItems(2) = RdoAux2!DESCVIGSANITARIA & " - " & !DESCVIGSANITARIA
           RdoAux2.Close
        Else
           On Error Resume Next
           Set itmX = lvVS.ListItems.Add(, "VS" & Format(!CODVIGSANIT, "000") & Format(!SUBCODVIGSANIT, "00"), Format(!CODVIGSANIT, "000"))
           itmX.SubItems(1) = Format(!SUBCODVIGSANIT, "00")
           itmX.SubItems(2) = !DESCVIGSANITARIA
        End If
       .MoveNext
    Loop
End With

'TX.LIC.
Sql = "SELECT MOBILIARIO.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE FROM MOBILIARIO INNER JOIN "
Sql = Sql & "ATIVIDADE ON MOBILIARIO.CODATIVIDADE = ATIVIDADE.CODATIVIDADE "
Sql = Sql & "Where MOBILIARIO.CODIGOMOB = " & Val(txtCod.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       Set itmX = lvTL.ListItems.Add(, "IS" & Format(!CODATIVIDADE, "000"), Format(!CODATIVIDADE, "000"))
       itmX.SubItems(1) = !DESCATIVIDADE
      .MoveNext
    Loop
End With
Sql = "SELECT MOBILIARIOATIVIDADETL.CODATIVIDADE,ATIVIDADE.DESCATIVIDADE FROM ATIVIDADE INNER JOIN "
Sql = Sql & "MOBILIARIOATIVIDADETL ON ATIVIDADE.CODATIVIDADE = MOBILIARIOATIVIDADETL.CODATIVIDADE "
Sql = Sql & "WHERE MOBILIARIOATIVIDADETL.CODIGOMOB =" & Val(txtCod.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
       Set itmX = lvTL.ListItems.Add(, "IS" & Format(!CODATIVIDADE, "000"), Format(!CODATIVIDADE, "000"))
       itmX.SubItems(1) = !DESCATIVIDADE
      .MoveNext
    Loop
End With

End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.text)
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
Dim nCodImovel As Long

If Val(txtCod.text) = 0 Then Exit Sub
nCodImovel = Val(txtCod.text)
txtAno.text = Year(Now)
Limpa
Sql = "SELECT CODREDUZIDO,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & txtCod.text
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If !Inativo = 1 Then
           MsgBox "Este imóvel encontra-se inativo.", vbExclamation, "Atenção"
           Exit Sub
        End If
        CarregaImovel nCodImovel
    Else
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR,CODCIDADE "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & txtCod.text
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               lblNumInsc.Caption = SubNull(!INSCESTADUAL)
               txtProp.text = !RAZAOSOCIAL
               If Not IsNull(!DATAENCERRAMENTO) Or !DATAENCERRAMENTO <> CDate("01/01/1900") Then
                  MsgBox "Esta empresa foi encerrada em " & Format(!DATAENCERRAMENTO, "dd/mm/yyyy"), vbExclamation, "Atenção"
                  Exit Sub
               End If
              'suspenção
               Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.text
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
               
               lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
               If Trim(lblRua.Caption) = "" Then
                    lblRua.Caption = SubNull(!NOMELOGR)
               End If
               lblNumImovel.Caption = Val(SubNull(!Numero))
               lblCep.Caption = IIf(IsNull(!cep), "", Left$(!cep, 5) & "-" & Right$(!cep, 3))
               lblCompl.Caption = SubNull(!COMPLEMENTO)
               Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=" & !CODCIDADE & " AND CODBAIRRO=" & !CODBAIRRO
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
               Sql = Sql & "BAIRRO.CODBAIRRO = MOBILIARIOENDENTREGA.CODBAIRRO WHERE MOBILIARIOENDENTREGA.CODMOBILIARIO=" & Val(txtCod.text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                    If .RowCount > 0 Then
                        lbltipoend.Caption = "(Endereço de Entrega Específico)"
                        lblRuaEntrega.Caption = SubNull(!NomeLogradouro)
                        lblNumEntrega.Caption = SubNull(!NUMIMOVEL)
                        lblComplentrega.Caption = SubNull(!COMPLEMENTO)
                        lblBairroEntrega.Caption = IIf(IsNull(!DescBairro), SubNull(!DescBairro1), SubNull(!DescBairro))
                        lblCidadeEntrega.Caption = IIf(IsNull(!desccidade), SubNull(!DESCCIDADE1), SubNull(!desccidade))
                        lblCepEntrega.Caption = SubNull(!cep)
                        lbluf.Caption = SubNull(!UF)
                    Else
                        lbltipoend.Caption = "(Endereço da Empresa)"
                        lblRuaEntrega.Caption = lblRua.Caption
                        lblNumEntrega.Caption = lblNumImovel.Caption
                        lblComplentrega.Caption = lblCompl.Caption
                        lblBairroEntrega.Caption = lblBairro.Caption
                        lblCidadeEntrega.Caption = "JABOTICABAL"
                        lblCepEntrega.Caption = lblCep.Caption
                        lbluf.Caption = "SP"
                    End If
                   .Close
               End With
               CarregaLista
            Else
               Sql = "SELECT CIDADAO.CODCIDADAO,CIDADAO.NOMECIDADAO,CIDADAO.CPF, CIDADAO.CNPJ, CIDADAO.CODLOGRADOURO,vwLOGRADOURO.ABREVTIPOLOG,"
               Sql = Sql & "vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,CIDADAO.NUMIMOVEL, CIDADAO.COMPLEMENTO,CIDADAO.CODBAIRRO, BAIRRO.DESCBAIRRO,"
               Sql = Sql & "CIDADAO.CODCIDADE, CIDADE.DESCCIDADE,CIDADAO.SIGLAUF, UF.DESCUF, CIDADAO.CEP,CIDADAO.NOMELOGRADOURO AS RUA2,CIDADAO.NOMECIDADE,"
               Sql = Sql & "CIDADAO.NOMEBAIRRO,CIDADAO.NOMEUF FROM vwLOGRADOURO RIGHT OUTER JOIN CIDADAO ON vwLOGRADOURO.CODLOGRADOURO = CIDADAO.CODLOGRADOURO "
               Sql = Sql & "LEFT OUTER JOIN CIDADE INNER JOIN BAIRRO ON CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE INNER JOIN "
               Sql = Sql & "UF ON CIDADE.SIGLAUF = UF.SIGLAUF ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO WHERE CODCIDADAO=" & Val(txtCod.text)
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                   If .RowCount > 0 Then
                       txtProp.text = !NOMECIDADAO
                       If !CODLOGRADOURO > 0 Then
                          lblRua.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                       Else
                          lblRua.Caption = SubNull(!RUA2)
                       End If
                       lblNumImovel.Caption = Val(SubNull(!NUMIMOVEL))
                       If IsNull(!cep) Then
                          lblCep.Caption = ""
                       Else
                          lblCep.Caption = IIf(IsNull(!cep), "", Left$(!cep, 5) & "-" & Right$(!cep, 3))
                       End If
                       lblCompl.Caption = SubNull(!COMPLEMENTO)
                       lblBairro.Caption = IIf(IsNull(!DescBairro), SubNull(!NOMEBAIRRO), !DescBairro)
                       lblNumImovel.Caption = Val(SubNull(!NUMIMOVEL))
                       If IsNull(!cep) Then
                          lblCep.Caption = ""
                       Else
                          lblCep.Caption = IIf(IsNull(!cep), "", Left$(!cep, 5) & "-" & Right$(!cep, 3))
                       End If
                       lblCompl.Caption = SubNull(!COMPLEMENTO)
                       lblBairro.Caption = IIf(IsNull(!DescBairro), SubNull(!NOMEBAIRRO), !DescBairro)
                   
                       lblRuaEntrega.Caption = lblRua.Caption
                       lblNumEntrega.Caption = lblNumImovel.Caption
                       lblComplentrega.Caption = lblCompl.Caption
                       lblBairroEntrega.Caption = lblBairro.Caption
                       lblCidadeEntrega.Caption = IIf(IsNull(!desccidade), SubNull(!NomeCidade), !desccidade)
                       lblCepEntrega.Caption = lblCep.Caption
                       lbluf.Caption = IIf(IsNull(!SiglaUF), SubNull(!NOMEUF), !SiglaUF)
                   Else
                       MsgBox "Código não cadastrado.", vbCritical, "Atenção"
                   End If
                  .Close
               End With
            End If
           .Close
        End With
    End If
End With
End Sub

Private Sub Limpa()

'z = SendMessage(lvTrib.hwnd, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvISS.hwnd, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvTL.hwnd, LVM_DELETEALLITEMS, 0, 0)
z = SendMessage(lvVS.hwnd, LVM_DELETEALLITEMS, 0, 0)
txtProp.text = ""
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
lbluf.Caption = ""
lblNumInsc.Caption = ""
'lbllanc.Caption = ""
lbltipoend.Caption = ""
chkUnica.Value = 0
'pnlObs.Visible = False
'txtObs.text = ""
'lblDataVencto.Caption = "  /  /    "
'txtNumParc.text = ""
'LimpaMascara mskDataInicio
'For x = 1 To 12
'    LimpaMascara mskVenc(x)
'Next
'bExec = False: cmbLanc.ListIndex = -1: bExec = True
End Sub

