VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCnsCidadao 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Cidadão"
   ClientHeight    =   5310
   ClientLeft      =   7935
   ClientTop       =   3945
   ClientWidth     =   6705
   Icon            =   "frmCnsCidadao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optOrdem 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Código"
      Height          =   210
      Index           =   4
      Left            =   690
      TabIndex        =   8
      Top             =   510
      Width           =   945
   End
   Begin VB.OptionButton optOrdem 
      BackColor       =   &H00EEEEEE&
      Caption         =   "RG"
      Height          =   210
      Index           =   3
      Left            =   4665
      TabIndex        =   7
      Top             =   495
      Width           =   945
   End
   Begin VB.OptionButton optOrdem 
      BackColor       =   &H00EEEEEE&
      Caption         =   "CNPJ"
      Height          =   210
      Index           =   2
      Left            =   3660
      TabIndex        =   6
      Top             =   510
      Width           =   945
   End
   Begin VB.OptionButton optOrdem 
      BackColor       =   &H00EEEEEE&
      Caption         =   "CPF"
      Height          =   210
      Index           =   1
      Left            =   2670
      TabIndex        =   5
      Top             =   495
      Width           =   945
   End
   Begin VB.OptionButton optOrdem 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Nome"
      Height          =   210
      Index           =   0
      Left            =   1665
      TabIndex        =   4
      Top             =   495
      Value           =   -1  'True
      Width           =   945
   End
   Begin VB.TextBox txtPesq 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   990
      TabIndex        =   0
      Top             =   90
      Width           =   4305
   End
   Begin MSComctlLib.ListView lvCid 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nome do Cidadão"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Endereço"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CPF"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "CNPJ"
         Object.Width           =   2542
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CODIGO"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Imóvel"
         Object.Width           =   2540
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdPesq 
      Height          =   345
      Left            =   5430
      TabIndex        =   3
      ToolTipText     =   "Pesquisar"
      Top             =   75
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "C&onsultar"
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
      MICON           =   "frmCnsCidadao.frx":014A
      PICN            =   "frmCnsCidadao.frx":0166
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
      Left            =   4290
      TabIndex        =   9
      ToolTipText     =   "Cancelar Edição"
      Top             =   4800
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
      MICON           =   "frmCnsCidadao.frx":02C0
      PICN            =   "frmCnsCidadao.frx":02DC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Retorna Cidadão Selecionado"
      Top             =   4800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Selecionar"
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
      MICON           =   "frmCnsCidadao.frx":0436
      PICN            =   "frmCnsCidadao.frx":0452
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
      Caption         =   "Pesquisa:"
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   150
      Width           =   795
   End
End
Attribute VB_Name = "frmCnsCidadao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim RdoAux As rdoResultset
Dim Sql As String, NodX As Object
Dim NomeForm As String, sTipoCid As String

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property
Public Property Let sTipoCidadao(sValue As String)
    sTipoCid = sValue
End Property

Private Sub cmdCancel_Click()
CodCidadao = 0
CodEmpresa = 0

Unload Me
End Sub

Private Sub cmdConsultar_Click()
Dim x As Integer
If lvCid.ListItems.Count = 0 Then
   MsgBox "Selecione um Cidadão.", vbExclamation, "Atenção"
   Exit Sub
End If
CodCidadao = Val(Mid$(lvCid.SelectedItem.Key, 2, Len(lvCid.SelectedItem.Key) - 7))
modLg "Consulta Cidadão: " & CodCidadao & " - " & lvCid.SelectedItem.Text
If CodCidadao > 0 Then
   If NomeForm = "frmCadImob" Then
       For x = 1 To frmCadImob.tvProp.Nodes.Count
           If Right$(frmCadImob.tvProp.Nodes(x).Key, 6) = Format(CodCidadao, "000000") Then
              MsgBox "Ja existe um Proprietário ou Proprietário Solidário com este nome.", vbCritical, "Atenção"
              GoTo fim
           End If
       Next
       If sTipoCid = "P" Then
          If frmCadImob.tvProp.Nodes("PROP").Children > 0 Then
             Set NodX = frmCadImob.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text, 1)
          Else
             Set NodX = frmCadImob.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text & " - Principal", 1)
          End If
          frmCadImob.tvProp.Nodes("PROP" & Format(CodCidadao, "000000")).ForeColor = vbBlue
       Else
          Set NodX = frmCadImob.tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text, 2)
          frmCadImob.tvProp.Nodes("COMP" & Format(CodCidadao, "000000")).ForeColor = vbBlue
       End If
    
       For x = 1 To frmCadImob.tvProp.Nodes.Count
          frmCadImob.tvProp.Nodes(x).EnsureVisible
       Next
       
       
'       Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & CodCidadao & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'       Sql = Sql & "O Cidadão foi incluido como proprietário/proprietário solidário do imóvel de inscrição:" & frmCadImob.lblIC.Caption & "." & frmCadImob.lblUnid.Caption & "." & frmCadImob.lblSubUnid.Caption & "')"
       Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & CodCidadao & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
       Sql = Sql & "O Cidadão foi incluido como proprietário/proprietário solidário do imóvel de inscrição:" & frmCadImob.lblIC.Caption & "." & frmCadImob.lblUnid.Caption & "." & frmCadImob.lblSubUnid.Caption & "')"
       cn.Execute Sql, rdExecDirect
       
fim:
       CodCidadao = 0
       sTipoCid = ""
   ElseIf NomeForm = "frmDesmembramento" Then
       For x = 1 To frmDesmembramento.tvProp.Nodes.Count
           If Right$(frmDesmembramento.tvProp.Nodes(x).Key, 6) = Format(CodCidadao, "000000") Then
              MsgBox "Ja existe um Proprietário ou Proprietário Solidário com este nome.", vbCritical, "Atenção"
              GoTo Fim2
           End If
       Next
       If sTipoCid = "P" Then
          If frmDesmembramento.tvProp.Nodes("PROP").Children > 0 Then
             Set NodX = frmDesmembramento.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text, 1)
          Else
             Set NodX = frmDesmembramento.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text & " - Principal", 1)
          End If
          frmDesmembramento.tvProp.Nodes("PROP" & Format(CodCidadao, "000000")).ForeColor = vbBlue
       Else
          Set NodX = frmDesmembramento.tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text, 2)
          frmDesmembramento.tvProp.Nodes("COMP" & Format(CodCidadao, "000000")).ForeColor = vbBlue
       End If
    
       For x = 1 To frmDesmembramento.tvProp.Nodes.Count
          frmDesmembramento.tvProp.Nodes(x).EnsureVisible
       Next
Fim2:
       CodCidadao = 0
       sTipoCid = ""
   ElseIf NomeForm = "frmUnifica" Then
       For x = 1 To frmUnifica.tvProp.Nodes.Count
           If Right$(frmUnifica.tvProp.Nodes(x).Key, 6) = Format(CodCidadao, "000000") Then
              MsgBox "Ja existe um Proprietário ou Proprietário Solidário com este nome.", vbCritical, "Atenção"
              GoTo fim3
           End If
       Next
       If sTipoCid = "P" Then
          If frmUnifica.tvProp.Nodes("PROP").Children > 0 Then
             Set NodX = frmUnifica.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text, 1)
          Else
             Set NodX = frmUnifica.tvProp.Nodes.Add("PROP", tvwChild, "PROP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text & " - Principal", 1)
          End If
          frmUnifica.tvProp.Nodes("PROP" & Format(CodCidadao, "000000")).ForeColor = vbBlue
       Else
          Set NodX = frmUnifica.tvProp.Nodes.Add("COMP", tvwChild, "COMP" & Format(CodCidadao, "000000"), lvCid.SelectedItem.Text, 2)
          frmUnifica.tvProp.Nodes("COMP" & Format(CodCidadao, "000000")).ForeColor = vbBlue
       End If
    
       For x = 1 To frmUnifica.tvProp.Nodes.Count
          frmUnifica.tvProp.Nodes(x).EnsureVisible
       Next
fim3:
       CodCidadao = 0
       sTipoCid = ""
   ElseIf NomeForm = "frmCadMob" Then
       Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF FROM CIDADAO WHERE CODCIDADAO=" & CodCidadao
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       If sTipoCid = "R" Then
            frmCadMob.txtNomeProf.Text = Format(CodCidadao, "000000") & " - " & RdoAux!nomecidadao
       Else
            If frmCadMob.grdProp.Rows > 1 Then
                 For x = 1 To frmCadMob.grdProp.Rows - 1
                     If frmCadMob.grdProp.TextMatrix(x, 0) = Format(CodCidadao, "000000") Then
                        MsgBox "Ja existe um Proprietário/Sócio com este nome.", vbCritical, "Atenção"
                        Exit Sub
                     End If
                 Next
            End If
            frmCadMob.grdProp.AddItem Format(CodCidadao, "000000") & Chr(9) & RdoAux!nomecidadao & Chr(9) & SubNull(RdoAux!CPF)
       End If
       RdoAux.Close
Fim4:
       CodCidadao = 0
       CodEmpresa = 0
       sTipoCid = ""
   ElseIf NomeForm = "frmCnsImovel" Then
       frmCnsImovel.txtProp.Text = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "frmCnsAvancadaImob" Then
       frmCnsAvancadaImob.txtProp.Text = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "frmCadastroRural" Then
       frmCadastroRural.lblProp.Caption = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "frmCnsMob" Then
       frmCnsMob.txtProp.Text = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "frmCnsRegAtend" Then
       frmCnsRegAtend.txtProp.Text = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "frmDebitoImob" Then
       frmDebitoImob.txtCod.Text = Format(CodCidadao, "000000")
       frmDebitoImob.lblProp.Caption = lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "frm2ViaLaser" Then
       frm2ViaLaser.txtCod.Text = Format(CodCidadao, "000000")
       frm2ViaLaser.lblProp.Caption = lvCid.SelectedItem.Text
       sTipoCid = ""
   ElseIf NomeForm = "ALUGUEL" Then
       frmManAluguel.txtCod.Text = Format(CodCidadao, "000000")
   ElseIf NomeForm = "frmResumoProtocolo" Then
       frmResumoProtocolo.lblReq.Caption = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
   ElseIf NomeForm = "2VIA" Then
       frmEmissao2Via.txtCod.Text = Format(CodCidadao, "000000")
   ElseIf NomeForm = "2VIAe" Then
       frmEmissao2ViaEspecial.txtCod.Text = Format(CodCidadao, "000000")
   ElseIf NomeForm = "frmCertidao" Then
       frmCertidao.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmRequerimento" Then
       frmRequerimento.lblCodRequerente.Caption = Format(CodCidadao, "000000")
       frmRequerimento.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmRequerIPTU" Then
       frmRequerIPTU.lblRequerente.Tag = Format(CodCidadao, "000000")
       frmRequerIPTU.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmGare" Then
       frmGare.lblCodRequerente.Caption = Format(CodCidadao, "000000")
       frmGare.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmGuiaPratico3" Then
       frmGuiaPratico3.lblCodRequerente.Caption = Format(CodCidadao, "000000")
       frmGuiaPratico3.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmGuiaPratico4" Then
    '   frmGuiaPratico4.lblCodRequerente.Caption = Format(CodCidadao, "000000")
       frmGuiaPratico4.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmFundoDespesa" Then
       frmFundoDespesa.lblCodRequerente.Caption = Format(CodCidadao, "000000")
       frmFundoDespesa.lblRequerente.Caption = lvCid.SelectedItem.Text
   ElseIf NomeForm = "frmConfissaoDivida" Then
       frmConfissaoDivida.lblRequerente.Caption = Format(CodCidadao, "000000") & " - " & lvCid.SelectedItem.Text
       Sql = "SELECT * FROM vwCIDADAO WHERE CODCIDADAO=" & CodCidadao
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux
            If .RowCount > 0 Then
                frmConfissaoDivida.lblEndCor.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2
                If Trim$(frmConfissaoDivida.lblEndCor.Caption) = "" Then
                    frmConfissaoDivida.lblEndCor.Caption = SubNull(!NomeLogradouro)
                End If
                If Trim$(frmConfissaoDivida.lblEndCor.Caption) <> "" Then
                   frmConfissaoDivida.lblEndCor.Caption = frmConfissaoDivida.lblEndCor.Caption & " Nº " & !NUMIMOVEL
                End If
                If Trim(!CPF) <> "" Then
                    frmConfissaoDivida.lblCPF.Caption = !CPF
                Else
                    If Trim(!Cnpj) <> "" Then
                        frmConfissaoDivida.lblCPF.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
                    End If
                End If
            End If
           .Close
       End With
   End If
End If


Unload Me
End Sub

Private Sub cmdPesq_Click()
Ocupado
If txtPesq.Text = "" Then
   MsgBox "Digite o início da pesquisa.", vbExclamation, "Atenção"
   txtPesq.SetFocus
   Liberado
   Exit Sub
End If

Screen.MousePointer = vbHourglass
CarregaLista txtPesq.Text
Screen.MousePointer = vbDefault

If lvCid.ListItems.Count = 0 Then
   MsgBox "Nenhum registro coincidente.", vbInformation, "Atenção"
End If
Liberado
End Sub

Private Sub Form_Deactivate()
'Me.ZOrder 0
End Sub

Private Sub Form_Load()
Ocupado
'Add3DBorder lvCid
Centraliza Me
Liberado
End Sub

Private Sub CarregaLista(Letra As String)

Dim itmX As ListItem
Dim z As Long, sNomeLogr As String
z = SendMessage(lvCid.hwnd, LVM_DELETEALLITEMS, 0, 0)
Ocupado
Sql = "SELECT     cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,"
Sql = Sql & "   vwLOGRADOURO.NomeLogradouro , Proprietario.CODREDUZIDO FROM  cidadao LEFT OUTER JOIN proprietario ON cidadao.codcidadao = proprietario.codcidadao LEFT OUTER JOIN "
Sql = Sql & "  vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO WHERE "
If NomeForm = "frmDebitoImob" Or NomeForm = "frm2ViaLaser" Or NomeForm = "2VIA" Then
    Sql = Sql & "cidadao.CODCIDADAO > 500000  AND   "
End If

If optOrdem(0).value = True Then
   Sql = Sql & "NOMECIDADAO LIKE '" & Mask(Letra) & "%' ORDER BY NOMECIDADAO"
ElseIf optOrdem(1).value = True Then
   Sql = Sql & "CPF LIKE '%" & Letra & "%' ORDER BY CPF"
ElseIf optOrdem(2).value = True Then
   Sql = Sql & "CNPJ LIKE '%" & Letra & "%' ORDER BY CNPJ"
ElseIf optOrdem(3).value = True Then
   Sql = Sql & "RG LIKE '" & Letra & "%' ORDER BY RG"
ElseIf optOrdem(4).value = True Then
   Sql = Sql & "cidadao.CODCIDADAO LIKE '" & Letra & "%' ORDER BY cidadao.CODCIDADAO"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    
    Do Until .EOF
       Set itmX = lvCid.ListItems.Add(, "C" & Format(!CodCidadao, "000000") & Format(Val(SubNull(!CODREDUZIDO)), "000000"), !nomecidadao)
       itmX.SubItems(1) = sNomeLogr
       itmX.SubItems(2) = SubNull(!CPF)
       itmX.SubItems(3) = SubNull(!Cnpj)
       itmX.SubItems(4) = SubNull(!rg)
       itmX.SubItems(5) = SubNull(!CodCidadao)
       itmX.SubItems(6) = SubNull(!CODREDUZIDO)
      .MoveNext
    Loop
   .Close
End With
Liberado
End Sub

Private Sub Form_LostFocus()
'Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

For x = 0 To Forms.Count - 1
    If Forms(x).Name = "frmCidadao" Then
        frmCidadao.bZOrder = True
        frmCidadao.ZOrder 0
        Exit For
    End If
Next

End Sub

Private Sub lvCid_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvCid.SortKey = ColumnHeader.Position - 1
lvCid.Sorted = True
lvCid.SortOrder = lvwAscending
End Sub

Private Sub lvCid_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim sNomeLogr As String, nCodCidadao As Long

nCodCidadao = Val(Item.SubItems(5))

'If nCodCidadao < 100000 Then
'    Sql = "SELECT  dbo.vwCnsImovel.ABREVTIPOLOG, dbo.vwCnsImovel.ABREVTITLOG, dbo.vwCnsImovel.NOMELOGRADOURO, dbo.vwCnsImovel.LI_NUM "
'    Sql = Sql & "FROM dbo.CIDADAO INNER JOIN dbo.PROPRIETARIO ON dbo.CIDADAO.CODCIDADAO = dbo.PROPRIETARIO.CODCIDADAO LEFT OUTER JOIN "
'    Sql = Sql & "dbo.vwCnsImovel ON dbo.PROPRIETARIO.CODREDUZIDO = dbo.vwCnsImovel.CODREDUZIDO "
'    Sql = Sql & "Where dbo.CIDADAO.CodCidadao =" & nCodCidadao
    Sql = "select * from vwfullcidadao where codcidadao=" & nCodCidadao
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
'         If IsNull(RdoAux2!NomeLogradouro) Then
'            sNomeLogradouro = ""
'         Else
'            sNomeLogr = Trim$(SubNull(RdoAux2!AbrevTipoLog)) & " " & Trim$(SubNull(RdoAux2!AbrevTitLog)) & " " & RdoAux2!NomeLogradouro & " Nº " & RdoAux2!Li_Num
 '        End If
          sNomeLogr = SubNull(RdoAux2!Endereco) & ", " & RdoAux2!NUMIMOVEL
    Else
         sNomeLogr = ""
    End If
    Item.SubItems(1) = sNomeLogr
'End If

End Sub

Private Sub txtPesq_KeyPress(KeyAscii As Integer)
Ocupado
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If txtPesq.Text = "" Then
      MsgBox "Digite o início da pesquisa.", vbExclamation, "Atenção"
      txtPesq.SetFocus
      Exit Sub
   Else
      CarregaLista txtPesq.Text
   End If
End If
Liberado
End Sub

