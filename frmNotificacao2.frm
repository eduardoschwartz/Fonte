VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmNotificacao2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notificação ISS Construção Civil"
   ClientHeight    =   3030
   ClientLeft      =   6045
   ClientTop       =   3465
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   5835
   Begin VB.CheckBox chkHabitese 
      Alignment       =   1  'Right Justify
      Caption         =   "Pedido de HABITE-SE"
      Height          =   195
      Left            =   2940
      TabIndex        =   26
      Top             =   1770
      Width           =   2355
   End
   Begin VB.TextBox txtISSPago 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4380
      MaxLength       =   10
      TabIndex        =   8
      Top             =   1350
      Width           =   945
   End
   Begin MSComCtl2.DTPicker mskDataVencto 
      Height          =   330
      Left            =   1440
      TabIndex        =   9
      Top             =   2535
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Format          =   104726529
      CurrentDate     =   40750
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmNotificacao2.frx":0000
      Left            =   1455
      List            =   "frmNotificacao2.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   960
   End
   Begin VB.ComboBox cmbCateg 
      Height          =   315
      ItemData        =   "frmNotificacao2.frx":0004
      Left            =   1440
      List            =   "frmNotificacao2.frx":0023
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1710
      Width           =   1365
   End
   Begin VB.ComboBox cmbUso 
      Height          =   315
      ItemData        =   "frmNotificacao2.frx":0069
      Left            =   1440
      List            =   "frmNotificacao2.frx":0076
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1305
      Width           =   1365
   End
   Begin VB.TextBox txtArea 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4380
      MaxLength       =   10
      TabIndex        =   7
      Top             =   945
      Width           =   945
   End
   Begin VB.TextBox txtNot 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4380
      MaxLength       =   6
      TabIndex        =   5
      Top             =   135
      Width           =   945
   End
   Begin VB.TextBox txtCodCidadao 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4380
      MaxLength       =   6
      TabIndex        =   6
      Top             =   540
      Width           =   945
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1455
      MaxLength       =   15
      TabIndex        =   2
      Top             =   945
      Width           =   945
   End
   Begin VB.TextBox txtCodImovel 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1455
      MaxLength       =   6
      TabIndex        =   1
      Top             =   540
      Width           =   945
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   4455
      TabIndex        =   10
      ToolTipText     =   "Gravar os Dados"
      Top             =   2550
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmNotificacao2.frx":009E
      PICN            =   "frmNotificacao2.frx":00BA
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
      Height          =   270
      Left            =   5355
      TabIndex        =   24
      ToolTipText     =   "Novo cidadão"
      Top             =   540
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   476
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmNotificacao2.frx":045F
      PICN            =   "frmNotificacao2.frx":047B
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
      Caption         =   "ISS Pago...........:"
      Height          =   225
      Index           =   9
      Left            =   2970
      TabIndex        =   25
      Top             =   1395
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Vencto......:"
      Height          =   225
      Index           =   8
      Left            =   180
      TabIndex        =   23
      Top             =   2610
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total R$.....:"
      Height          =   195
      Index           =   1
      Left            =   2745
      TabIndex        =   22
      Top             =   2205
      Width           =   1320
   End
   Begin VB.Label lblValorTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   4140
      TabIndex        =   21
      Top             =   2205
      Width           =   960
   End
   Begin VB.Label lblValorM2 
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000040&
      Height          =   195
      Left            =   1485
      TabIndex        =   20
      Top             =   2205
      Width           =   1140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do m² R$..:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   19
      Top             =   2205
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Categ. constr....:"
      Height          =   225
      Index           =   7
      Left            =   180
      TabIndex        =   18
      Top             =   1755
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Uso constr........:"
      Height          =   225
      Index           =   6
      Left            =   180
      TabIndex        =   17
      Top             =   1365
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Área notificada..:"
      Height          =   225
      Index           =   5
      Left            =   2970
      TabIndex        =   16
      Top             =   990
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Notificação...:"
      Height          =   225
      Index           =   4
      Left            =   2970
      TabIndex        =   15
      Top             =   180
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Cidadão.:"
      Height          =   225
      Index           =   3
      Left            =   2970
      TabIndex        =   14
      Top             =   585
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº do Processo.:"
      Height          =   225
      Index           =   2
      Left            =   180
      TabIndex        =   13
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Imóvel...:"
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   12
      Top             =   585
      Width           =   1305
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exercício...........:"
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   11
      Top             =   180
      Width           =   1305
   End
End
Attribute VB_Name = "frmNotificacao2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum UsoConstrucao
    residencial = 1
    Industrial = 2
    Comercial = 3
End Enum

Enum Categoria
    Baixo = 0
    Médio = 1
    Alto = 2
    Único = 3
    Fino = 4
    Luxuoso = 5
    Barracao = 6
    Popular = 7
    Bom = 8
End Enum

Dim nCodTributo As Integer, nValor As Double, nValorTotal As Double

Private Sub cmbAno_Change()
ComboValor
End Sub

Private Sub cmbAno_Click()
ComboValor
End Sub

Private Sub cmbCateg_Click()
ComboValor
End Sub

Private Sub cmbUso_Click()
ComboValor
End Sub

Private Sub ComboValor()
Dim nUso As Integer, nCateg As Integer, nArea As Double, nValorPago As Double
Dim RdoAux As rdoResultset, Sql As String

'If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
lblValorM2.Caption = "R$ 0,00"
nCodTributo = 0
If cmbUso.ListIndex = -1 Then cmbUso.ListIndex = 0
If cmbCateg.ListIndex = -1 Then cmbCateg.ListIndex = 0
nUso = Val(Left(cmbUso.ItemData(cmbUso.ListIndex), 1))
nCateg = Val(cmbCateg.ItemData(cmbCateg.ListIndex))
nValorPago = 0

If nUso = UsoConstrucao.Industrial Then
    If nCateg = Categoria.Único Then
        nCodTributo = 185
    ElseIf nCateg = Categoria.Barracao Then
        nCodTributo = 671
    ElseIf nCateg = Categoria.Popular Then
        nCodTributo = 672
    ElseIf nCateg = Categoria.Médio Then
        nCodTributo = 673
    ElseIf nCateg = Categoria.Bom Then
        nCodTributo = 674
    End If
Else
    cmbCateg.Enabled = True
    cmbCateg.BackColor = Branco
    If nUso = UsoConstrucao.residencial Then
        If nCateg = Categoria.Baixo Then
            nCodTributo = 179
        ElseIf nCateg = Categoria.Médio Then
            nCodTributo = 180
        ElseIf nCateg = Categoria.Alto Then
            nCodTributo = 181
        ElseIf nCateg = Categoria.Fino Then
            nCodTributo = 676
        ElseIf nCateg = Categoria.Luxuoso Then
            nCodTributo = 670
        Else
            nCodTributo = 0
        End If
    ElseIf nUso = UsoConstrucao.Comercial Then
        If nCateg = Categoria.Baixo Then
            nCodTributo = 182
        ElseIf nCateg = Categoria.Médio Then
            nCodTributo = 183
        ElseIf nCateg = Categoria.Alto Then
            nCodTributo = 184
        ElseIf nCateg = Categoria.Fino Then
            nCodTributo = 675
        Else
            nCodTributo = 0
        End If
    End If
End If

If nCodTributo = 0 Then
    nValor = 0
Else
    Sql = "select valoraliq from tributoaliquota where ano=" & Val(cmbAno.Text) & " and codtributo=" & nCodTributo
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount = 0 Then
        nValor = 0
    Else
        nValor = RdoAux!valoraliq
    End If
    RdoAux.Close
End If


lblValorM2.Caption = "R$ " & FormatNumber(nValor, 4)
If txtArea.Text = "" Then txtArea.Text = 0
nArea = CDbl(txtArea.Text)
nValorTotal = nValor * nArea
If txtISSPago.Text <> "" Then
    nValorPago = CDbl(txtISSPago.Text)
Else
    nValorPago = 0
End If
nValorTotal = nValorTotal - nValorPago
lblValorTotal.Caption = FormatNumber(nValorTotal, 2)

End Sub

Private Sub cmdNovo_Click()
Dim z As Variant, nCod As Long, Sql As String, RdoAux As rdoResultset
z = InputBox("Digite o nome do novo cidadão.", "Inclusão de Cidadão")
If z <> "" Then
    If MsgBox("Deseja incluir o nome " & z & " no cadastro de cidadões?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "SELECT MAX(CODCIDADAO) AS MAXIMO FROM CIDADAO"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nCod = !maximo + 1
            .Close
        End With
        Sql = "INSERT CIDADAO(CODCIDADAO,NOMECIDADAO) VALUES(" & nCod & ",'" & Mask(CStr(z)) & "')"
        cn.Execute Sql, rdExecDirect
        txtCodCidadao.Text = nCod
        
'        Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'        Sql = Sql & "Cidadão criado através da tela de Notificação de ISS Construção Civil')"
        Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
        Sql = Sql & "Cidadão criado através da tela de Notificação de ISS Construção Civil')"
        cn.Execute Sql, rdExecDirect
    End If
End If

End Sub

Private Sub Form_Load()
Dim x As Integer

For x = 2000 To Year(Now) + 1
    cmbAno.AddItem x
Next
cmbAno.Text = Year(Now)
cmbUso.ListIndex = 0
cmbCateg.ListIndex = 0
mskDataVencto.value = Now
Centraliza Me
End Sub

Private Sub txtArea_Change()
ComboValor
End Sub

Private Sub txtArea_GotFocus()
txtArea.SelStart = 0
txtArea.SelLength = Len(txtArea.Text)
End Sub

Private Sub txtArea_KeyPress(KeyAscii As Integer)
Tweak txtArea, KeyAscii, DecimalPositive
End Sub

Private Sub txtCodCidadao_KeyPress(KeyAscii As Integer)
Tweak txtCodCidadao, KeyAscii, IntegerPositive
End Sub

Private Sub txtCodImovel_KeyPress(KeyAscii As Integer)
Tweak txtCodImovel, KeyAscii, IntegerPositive
End Sub

Private Sub txtISSPago_Change()
ComboValor
End Sub

Private Sub txtISSPago_GotFocus()
txtISSPago.SelStart = 0
txtISSPago.SelLength = Len(txtISSPago.Text)

End Sub

Private Sub txtISSPago_KeyPress(KeyAscii As Integer)
Tweak txtISSPago, KeyAscii, DecimalPositive
End Sub

Private Sub txtISSPago_LostFocus()
ComboValor
End Sub

Private Sub txtNot_KeyPress(KeyAscii As Integer)
Tweak txtNot, KeyAscii, IntegerPositive
End Sub

Private Sub cmdGravar_Click()
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, nAreaOld As Double, nCodReduz As Long, nDif As Integer, nArea As Double
Dim sNome As String, sInscricao As String, sEndereco As String, nNumero As Integer, sComplemento As String, sExtenso As String, nCodCidadao As Long
Dim sBairro As String, sCidade As String, sEndereco2 As String, sCep As String, nSeq As Integer, sProcesso As String, sInsc As String, nAno As Integer
Dim nTipoEnd As Integer, sLogradouro As String, sBAIRRO2 As String, sCEP2 As String, nNumero2 As Integer, nValorIss As Double, nSeq2 As Integer
Dim sTipo As String, nUso As Integer, nCateg As Integer, sDataBase As String, sDataVencto As String, sHist As String, nValorPago As Double

sDataBase = Format(Now, "dd/mm/yyyy")
sDataVencto = mskDataVencto.value
nAno = Val(cmbAno.Text)

If Trim(txtISSPago.Text) = "" Then txtISSPago.Text = "0"
nValorPago = CDbl(txtISSPago.Text)

If Not (Valida) Then Exit Sub
If MsgBox("Emitir notificação para o cidadão selecionado?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "DELETE FROM NOTIFICACAOISS WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "select * from vwfullimovel2 where codreduzido=" & Val(txtCodImovel.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    sNome = !nomecidadao
    nCodReduz = Val(txtCodImovel.Text)
    nCodCidadao = Val(txtCodCidadao.Text)
    sProcesso = txtNumProc.Text
    sEndereco = !Logradouro
    sEndereco2 = sEndereco
    nNumero = !Li_Num
    sBairro = !DescBairro
    sCep = RetornaCEP(!CodLogr, !Li_Num)
    sEndereco = sEndereco & " Nº " & nNumero & ", " & sBairro & " " & sCep
    nSeq = Val(txtNot.Text)
   ' nDif = Format(123.67, "#0.00")
    'sExtenso = Extenso(nDif)
    nArea = txtArea.Text
    nUso = cmbUso.ItemData(cmbUso.ListIndex)
    nCateg = cmbCateg.ItemData(cmbCateg.ListIndex)
    sTipo = cmbUso.Text
    If nUso <> UsoConstrucao.Industrial Then
        If nCateg = Categoria.Médio Then
            sTipo = sTipo & " Médio"
        ElseIf nCateg = Categoria.Único Then
            sTipo = sTipo & " Baixo"
        End If
    End If
            
    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodCidadao
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    sNome = RdoAux2!nomecidadao
    RdoAux2.Close
                
    Sql = "delete from notificacaoiss where seq=" & Val(txtNot.Text) & " and ano=" & nAno
    cn.Execute Sql, rdExecDirect
                
    Sql = "insert notificacaoiss(usuario,codigo,razao,processo,seq,ano,endereco,tipo,area,valoriss,valorpago,codcidadao,isspago) "
    Sql = Sql & "values('"
    Sql = Sql & IIf(NomeDeLogin = "GLEISE", "EDUARDO", NomeDeLogin) & "'," & nCodReduz & ",'" & Mask(sNome) & "','" & sProcesso & "'," & Val(txtNot.Text) & "," & nAno & ",'" & Left(Mask(sEndereco), 70) & "','" & sTipo & "',"
    Sql = Sql & Virg2Ponto(CStr(nArea)) & "," & Virg2Ponto(CStr(nValor)) & "," & Virg2Ponto(CStr(nValorTotal)) & "," & nCodCidadao & "," & Virg2Ponto(Format(nValorPago, "#0.00")) & ")"
    cn.Execute Sql, rdExecDirect
    
    'insere parcelas de calculo
    Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodCidadao & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=65"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeq2 = 1
        Else
            nSeq2 = !maximo + 1
        End If
       .Close
    End With
    
'    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USUARIO) VALUES("
'    Sql = Sql & nCodCidadao & "," & nAno & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "','" & Left$("GTI/" & NomeDeLogin, 25) & "')"
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,DATAVENCIMENTO,DATADEBASE,USERID) VALUES("
    Sql = Sql & nCodCidadao & "," & nAno & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & Format(sDataBase, "mm/dd/yyyy") & "'," & 236 & ")"
    cn.Execute Sql, rdExecDirect
    
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
    Sql = Sql & nCodCidadao & "," & nAno & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & nCodTributo & "," & Virg2Ponto(Format(nValorTotal, "#0.00")) & ")"
    cn.Execute Sql, rdExecDirect
    
    sHist = "Iss construção civil processo nº " & sProcesso & " notificação nº " & nSeq & "/" & cmbAno.Text & " codigo imóvel: " & nCodReduz
    
'    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES("
'    Sql = Sql & nCodCidadao & "," & nAno & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 0 & ",'" & Mask(sHist) & "','"
'    Sql = Sql & Left$("GTI/" & NomeDeLogin, 25) & "','" & Format(Now, "mm/dd/yyyy") & "')"
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES("
    Sql = Sql & nCodCidadao & "," & nAno & "," & 65 & "," & nSeq2 & "," & 1 & "," & 0 & "," & 0 & ",'" & Mask(sHist) & "',"
    Sql = Sql & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
    
    sHist = "Iss construção civil lançado no código " & nCodCidadao & " processo nº " & sProcesso & " notificação nº " & nSeq & "/" & cmbAno.Text & " Área notificada: " & txtArea.Text & " m²"
    
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!maximo) Then
            nSeq2 = 1
        Else
            nSeq2 = !maximo + 1
        End If
       .Close
    End With

'    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USUARIO,DATAHIST2) VALUES("
'    Sql = Sql & nCodReduz & "," & nSeq2 & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "','" & "GTI/Iss.C.Civil" & "','" & Format(Now, "mm/dd/yyyy") & "')"
    Sql = "INSERT HISTORICO(CODREDUZIDO,SEQ,DATAHIST,DESCHIST,USERID,DATAHIST2) VALUES("
    Sql = Sql & nCodReduz & "," & nSeq2 & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "'," & 236 & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect
   .Close
    
    Sql = "delete from notificacao_iss_ccivil where codigo_cidadao=" & nCodCidadao & " and codigo_imovel=" & nCodReduz & " and numero_notificacao=" & Val(txtNot.Text) & " and ano_notificacao=" & Val(cmbAno.Text)
    cn.Execute Sql, rdExecDirect
    
    Sql = "insert notificacao_iss_ccivil(codigo_cidadao,codigo_imovel,numero_notificacao,ano_notificacao,valor,data_gravacao,processo) values(" & nCodCidadao & "," & nCodReduz & ","
    Sql = Sql & Val(txtNot.Text) & "," & Val(cmbAno.Text) & "," & Virg2Ponto(Format(nValorTotal, "#0.00")) & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(txtNumProc.Text) & "')"
    cn.Execute Sql, rdExecDirect
    
End With

Liberado
If chkHabitese.value = 0 Then
    frmReport.ShowReport2 "NOTIFICACAO3", frmMdi.hwnd, Me.hwnd
Else
    frmReport.ShowReport2 "NOTIFICACAO4", frmMdi.hwnd, Me.hwnd
End If

txtArea.Text = ""
txtCodImovel.Text = ""
txtCodCidadao.Text = ""
txtNumProc.Text = ""
txtISSPago.Text = ""

Sql = "DELETE FROM NOTIFICACAOISS WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
Sql = "DELETE FROM NOTIFICACAOISS WHERE USUARIO='EDUARDO'"
cn.Execute Sql, rdExecDirect


End Sub

Private Function Valida() As Boolean

Dim Sql As String, RdoAux As rdoResultset

If Val(txtNot.Text) = 0 Then
    MsgBox "Digite o nº da notificação", vbExclamation, "Atenção"
    Valida = False
    Exit Function
End If

If Val(txtCodImovel.Text) = 0 Then
    MsgBox "Digite o código do imóvel", vbExclamation, "Atenção"
    Valida = False
    Exit Function
End If

If Val(txtCodCidadao.Text) = 0 Then
    MsgBox "Digite o código do cidadão.", vbExclamation, "Atenção"
    Valida = False
    Exit Function
End If

If txtNumProc.Text = "" Then
    MsgBox "Digite o número do processo", vbExclamation, "Atenção"
    Valida = False
    Exit Function
End If

If Val(txtArea.Text) = 0 Then
    MsgBox "Digite o valor da área.", vbExclamation, "Atenção"
    Valida = False
    Exit Function
End If

If Val(lblValorTotal.Caption) = 0 Then
    MsgBox "Valor total não calculado.", vbExclamation, "Atenção"
    Valida = False
    Exit Function
End If

Sql = "select nomecidadao from cidadao where codcidadao>=500000 and codcidadao=" & Val(txtCodCidadao.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        .Close
        MsgBox "Código cidadão não cadastrado.", vbExclamation, "Atenção"
        Valida = False
        Exit Function
    End If
   .Close
End With

Sql = "select distrito from cadimob where codreduzido=" & Val(txtCodImovel.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        .Close
        MsgBox "Imóvel não cadastrado.", vbExclamation, "Atenção"
        Valida = False
        Exit Function
    End If
   .Close
End With

Valida = True
End Function

