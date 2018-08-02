VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAvisoDebito 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aviso de Débito"
   ClientHeight    =   4875
   ClientLeft      =   2520
   ClientTop       =   4680
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   8415
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   6990
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   4365
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
      MICON           =   "frmAvisoDebito.frx":0000
      PICN            =   "frmAvisoDebito.frx":001C
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
      Left            =   5580
      TabIndex        =   11
      ToolTipText     =   "Impressão do Aviso de Débito"
      Top             =   4365
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
      MICON           =   "frmAvisoDebito.frx":008A
      PICN            =   "frmAvisoDebito.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtNumProcessoE 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6885
      MaxLength       =   15
      TabIndex        =   10
      Top             =   3015
      Width           =   1110
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Height          =   2445
      Left            =   90
      TabIndex        =   28
      Top             =   2385
      Width           =   4290
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Despesas de Cert.CRI......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   135
         TabIndex        =   9
         Top             =   2115
         Width           =   3300
      End
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Taxa Judiciária...........:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   135
         TabIndex        =   8
         Top             =   1800
         Width           =   3300
      End
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Despesas Postais..........:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   135
         TabIndex        =   7
         Top             =   1485
         Width           =   3300
      End
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Precatória................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   135
         TabIndex        =   6
         Top             =   1170
         Width           =   3300
      End
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Restit. de Cert.CRI.......:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   3300
      End
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Restit.de Diligências.....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   540
         Width           =   3300
      End
      Begin VB.CheckBox chkT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Honorários................:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   3300
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   6
         Left            =   3150
         TabIndex        =   36
         Top             =   2115
         Width           =   960
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   5
         Left            =   3150
         TabIndex        =   35
         Top             =   1800
         Width           =   960
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   4
         Left            =   3150
         TabIndex        =   33
         Top             =   1485
         Width           =   960
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   3
         Left            =   3150
         TabIndex        =   32
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   2
         Left            =   3150
         TabIndex        =   31
         Top             =   855
         Width           =   960
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   1
         Left            =   3150
         TabIndex        =   30
         Top             =   540
         Width           =   960
      End
      Begin VB.Label lblT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00800000&
         Height          =   240
         Index           =   0
         Left            =   3150
         TabIndex        =   29
         Top             =   225
         Width           =   960
      End
   End
   Begin VB.TextBox txtExtenso 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2655
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2070
      Width           =   5700
   End
   Begin VB.TextBox txtTaxa 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1395
      MaxLength       =   15
      TabIndex        =   2
      Top             =   2025
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1995
      Left            =   45
      TabIndex        =   13
      Top             =   -45
      Width           =   8340
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   1
         Top             =   900
         Width           =   1110
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1275
         MaxLength       =   6
         TabIndex        =   0
         Top             =   180
         Width           =   1110
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2445
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   5835
      End
      Begin VB.TextBox txtRequerente 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   945
         Width           =   5835
      End
      Begin VB.TextBox txtEndereco 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   585
         Width           =   7005
      End
      Begin VB.TextBox txtEnd2 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1305
         Width           =   7005
      End
      Begin VB.TextBox txtBairroCidade 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1665
         Width           =   4845
      End
      Begin VB.TextBox txtCEP 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   6900
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1380
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inscrição.........:"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   210
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Processo nº....:"
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   24
         Top             =   945
         Width           =   1245
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço........:"
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   23
         Top             =   585
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço........:"
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   22
         Top             =   1305
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro\Cidade..:"
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   21
         Top             =   1665
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP..:"
         Height          =   225
         Index           =   5
         Left            =   6315
         TabIndex        =   20
         Top             =   1665
         Width           =   480
      End
   End
   Begin VB.Label lblSoma 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   6885
      TabIndex        =   38
      Top             =   2610
      Width           =   1140
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Soma dos Tributos...............:"
      Height          =   225
      Index           =   8
      Left            =   4680
      TabIndex        =   37
      Top             =   2610
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Processo de execução nº....:"
      Height          =   225
      Index           =   7
      Left            =   4680
      TabIndex        =   34
      Top             =   3060
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total......:"
      Height          =   225
      Index           =   6
      Left            =   135
      TabIndex        =   26
      Top             =   2070
      Width           =   1245
   End
End
Attribute VB_Name = "frmAvisoDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkT_Click(Index As Integer)
Dim z As Variant, nTotal As Double

If chkT(Index).value = 0 Then
    lblT(Index).Caption = "0,00"
Else
    z = InputBox("Digite o valor para " & chkT(Index).Caption, "Valor requerido")
    If z <> "" Then
        If IsNumeric(z) Then
            lblT(Index).Caption = FormatNumber(z, 2)
        Else
            MsgBox "Valor inválido.", vbCritical, "Atenção"
        End If
    End If
    If lblT(Index).Caption = "0,00" Then
        chkT(Index).value = 0
    End If
End If

nTotal = 0
For X = 0 To 6
    nTotal = nTotal + CDbl(lblT(X).Caption)
Next
lblSoma.Caption = FormatNumber(nTotal, 2)

End Sub

Private Sub cmdPrint_Click()
Dim ValidaProcesso As String, sNumProcesso As String, nTotal As Double, X As Integer, nTaxa As Double

If txtNome.Text = "" Then
    MsgBox "Selecione a inscrição", vbExclamation, "Atenção"
    Exit Sub
End If

If txtRequerente.Text = "" Then
    MsgBox "Selecione o requerente", vbExclamation, "Atenção"
    Exit Sub
End If

If txtTaxa.Text = "" Then
    MsgBox "Digite o valor da taxa", vbExclamation, "Atenção"
    Exit Sub
End If

If CDbl(txtTaxa.Text) = 0 Then
    MsgBox "Digite o valor da taxa", vbExclamation, "Atenção"
    Exit Sub
End If

ValidaProcesso = "": sNumProcesso = txtNumProcessoE.Text
If InStr(1, sNumProcesso, "/", vbBinaryCompare) = 0 Then
    ValidaProcesso = "Nº do processo de execução inválido. Formato deve ser: Nº do Processo/Ano."
    ElseIf Not IsNumeric(Right$(sNumProcesso, 4)) Then
        ValidaProcesso = "Nº do processo de execução inválido. O ano deve ter 4 digitos."
    ElseIf IsNumeric(Right$(sNumProcesso, 5)) Then
        ValidaProcesso = "Nº do processo de execução inválido. O ano deve ter 4 digitos."
    ElseIf Not IsNumeric(Left$(sNumProcesso, 1)) Then
        ValidaProcesso = "Nº do processo de execução inválido."
End If

If ValidaProcesso <> "" Then
    MsgBox ValidaProcesso, vbExclamation, "Atenção"
    Exit Sub
End If

nTotal = 0
For X = 0 To 6
    nTotal = nTotal + CDbl(lblT(X).Caption)
Next

nTaxa = CDbl(txtTaxa.Text)
If Round(nTotal, 2) <> Round(nTaxa, 2) Then
    MsgBox "A soma dos tributos " & FormatNumber(nTotal, 2) & " não corresponde ao valor da taxa de " & txtTaxa.Text, vbExclamation, "Atenção"
    Exit Sub
End If

frmReport.ShowReport2 "AVISODEBITO", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub txtCod_Change()
txtNome.Text = "": txtEndereco.Text = ""
End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtCod_LostFocus()
Dim xImovel As clsImovel, nCodReduz As Long, RdoAux As rdoResultset, Sql As String

If Val(txtCod.Text) = 0 Then Exit Sub
Set xImovel = New clsImovel
nCodReduz = Val(txtCod.Text)

If nCodReduz < 100000 Then
   Sql = "SELECT CODREDUZIDO,SETOR,INATIVO FROM CADIMOB WHERE CODREDUZIDO=" & nCodReduz
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
        If .RowCount > 0 Then
            With xImovel
                .CarregaImovel nCodReduz
                txtNome.Text = .NomePropPrincipal
                txtEndereco.Text = Trim$(SubNull(.AbrevTipoLog)) & " " & Trim$(SubNull(.AbrevTitLog)) & " " & .NomeLogradouro & ", " & .Li_Num & " " & .Li_Compl
            End With
        Else
            MsgBox "Imóvel não Cadastrado.", vbExclamation, "Atenção"
            Exit Sub
        End If
       .Close
   End With
ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
   Sql = "SELECT CODIGOMOB,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,NOMELOGR FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
        If .RowCount > 0 Then
            txtNome.Text = !RazaoSocial
            If SubNull(!NomeLogradouro) = "" Then
                txtEndereco.Text = Trim$(SubNull(!NomeLogr)) & ", " & !Numero
            Else
                txtEndereco.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & ", " & !Numero
            End If
        Else
            MsgBox "Empresa não Cadastrada.", vbExclamation, "Atenção"
            Exit Sub
        End If
       .Close
   End With
ElseIf nCodReduz >= 500000 And nCodReduz < 700000 Then
    Sql = "SELECT NOMECIDADAO, ABREVTIPOLOG, ABREVTITLOG, NOMELOGRADOURO, NOMELOGRADOURO2, NUMIMOVEL, COMPLEMENTO "
    Sql = Sql & "From vwCIDADAO Where CodCidadao =" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
         If .RowCount > 0 Then
             txtNome.Text = !nomecidadao
             If Not IsNull(!NomeLogradouro) Then
                If !NomeLogradouro <> "" Then
                    txtEndereco.Text = Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL))
                Else
                    txtEndereco.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL
                End If
             Else
                txtEndereco.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL
             End If
         Else
             MsgBox "Cidadão não Cadastrado.", vbExclamation, "Atenção"
             Exit Sub
         End If
        .Close
    End With
Else
    MsgBox "Código não Cadastrado.", vbExclamation, "Atenção"
    Exit Sub
End If


End Sub

Private Sub txtNumProc_Change()
txtRequerente.Text = ""
End Sub

Private Sub txtNumProc_GotFocus()
txtRequerente.SelStart = 0
txtRequerente.SelLength = Len(txtRequerente.Text)
End Sub

Private Sub txtNumProc_LostFocus()
Dim nCodCidadao As Long, nNumProc As Long, nAnoProc As Integer, ValidaProcesso As String, sNumProcesso As String
Dim RdoAux As rdoResultset, Sql As String

If Trim(txtNumProc.Text) = "" Then Exit Sub

ValidaProcesso = "": sNumProcesso = txtNumProc.Text
If InStr(1, sNumProcesso, "/", vbBinaryCompare) = 0 Then
    ValidaProcesso = "Nº do processo inválido. Formato deve ser: Nº do Processo/Ano."
    ElseIf Not IsNumeric(Right$(sNumProcesso, 4)) Then
        ValidaProcesso = "Nº do processo inválido. O ano deve ter 4 digitos."
    ElseIf IsNumeric(Right$(sNumProcesso, 5)) Then
        ValidaProcesso = "Nº do processo inválido. O ano deve ter 4 digitos."
    ElseIf Not IsNumeric(Left$(sNumProcesso, 1)) Then
        ValidaProcesso = "Nº do processo inválido."
End If

If ValidaProcesso <> "" Then
    MsgBox ValidaProcesso, vbExclamation, "Atenção"
    Exit Sub
End If

If NovoProtocolo = 0 Then
    Sql = "SELECT CODCIDAPRO FROM PROCESSO WHERE ANOPROCESS=" & ExtraiAnoProcesso(txtNumProc.Text) & " AND NUMEROPROC=" & ExtraiNumeroProcesso(txtNumProc.Text)
    'Sql = "SELECT CODCIDAPRO FROM PROCESSO WHERE ANOPROCESS=" & Val(Right$(txtNumProc.Text, 4)) & " AND NUMEROPROC=" & Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5))
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
'    nNumProc = Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 1))
'    nAnoProc = Val(Right$(txtNumProc.Text, 4))
    nNumProc = ExtraiNumeroProcesso(txtNumProc.Text)
    nAnoProc = ExtraiAnoProcesso(txtNumProc.Text)
'    If Right$(nNumProc, 1) <> RetornaDVProcesso(Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))) Then
'        MsgBox "Número de Processo inválido", vbExclamation, "Atenção"
'        Exit Sub
'    Else
        'Sql = "SELECT CODCIDADAO FROM PROCESSOGTI WHERE ANO=" & nAnoProc & " AND NUMERO=" & Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))
        Sql = "SELECT CODCIDADAO FROM PROCESSOGTI WHERE ANO=" & nAnoProc & " AND NUMERO=" & nNumProc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Requerente não localizado.", vbExclamation, "Atenção"
                Exit Sub
            Else
                nCodCidadao = !CodCidadao
            End If
           .Close
        End With
 '   End If
End If

If nCodCidadao > 0 Then
    Sql = "SELECT *  FROM vwFULLCIDADAO WHERE CODCIDADAO=" & nCodCidadao
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            txtRequerente.Text = !nomecidadao
            txtEnd2.Text = Trim$(SubNull(!Endereco)) & ", " & Val(SubNull(!NUMIMOVEL))
            txtBairroCidade.Text = Trim$(SubNull(!DescBairro)) & " - " & Trim$(SubNull(!descCidade)) & " - " & Trim$(SubNull(!SiglaUF))
            txtCEP.Text = Format(Val(SubNull(!Cep)), "00000-000")
        End If
       .Close
    End With
Else
    lblRequerente.Caption = ""
End If

End Sub

Private Sub txtTaxa_Change()
txtExtenso.Text = ""
End Sub

Private Sub txtTaxa_KeyPress(KeyAscii As Integer)
Tweak txtTaxa, KeyAscii, DecimalPositive
End Sub

Private Sub txtTaxa_LostFocus()
If Trim(txtTaxa.Text) = "" Then Exit Sub
If CDbl(txtTaxa.Text) = 0 Then Exit Sub
txtExtenso.Text = "(" & Extenso(txtTaxa.Text) & ")"
End Sub
