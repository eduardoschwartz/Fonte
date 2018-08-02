VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGuiaAmbulante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de guias para ambulantes"
   ClientHeight    =   3495
   ClientLeft      =   11145
   ClientTop       =   6210
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   5445
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   135
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1260
      Width           =   5190
   End
   Begin VB.ComboBox cmbValor 
      Height          =   315
      ItemData        =   "frmGuiaAmbulante.frx":0000
      Left            =   2580
      List            =   "frmGuiaAmbulante.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   690
      Width           =   1335
   End
   Begin VB.TextBox txtQtde 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2565
      TabIndex        =   0
      Top             =   225
      Width           =   1335
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   4110
      TabIndex        =   1
      ToolTipText     =   "Cancelar Edição"
      Top             =   450
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   609
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGuiaAmbulante.frx":001B
      PICN            =   "frmGuiaAmbulante.frx":0037
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
      Caption         =   "Selecione o valor da diária.....:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   750
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Quantidade de guias a gerar..:"
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   270
      Width           =   2265
   End
End
Attribute VB_Name = "frmGuiaAmbulante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbValor_Click()
Dim sObs As String

Select Case cmbValor.ListIndex
    Case 0
        sObs = "1.Alimentos preparados, refrigerantes não engarrafados e produtos hortifrutigranjeiros." & vbCrLf
        sObs = sObs & "2.Aparelhos de uso doméstico, armarinhos, artefatos de couro, artigos de papelaria, artigos de toucador, brinquedos e presentes, artefatos de ferragens, plásticos, "
        sObs = sObs & "borrachas, vassouras e semelhantes doces, frutas, estatuetas, sorvetes e quadros." & vbCrLf
        sObs = sObs & "5.Artigos não especificados na Tabela."
    Case 1
        sObs = "3.Tecidos e roupas, refrigerantes engarrafados." & vbCrLf
        sObs = sObs & "4.Artigos para fumantes, artigos de jogos de azar, fogos de artifício, jóias, pedras preciosas, peles, relógios e confecções de luxo e bebidas alcoólicas." & vbCrLf
End Select

txtObs.Text = sObs
End Sub

Private Sub Form_Load()
cmbValor.ListIndex = 0
Centraliza Me
End Sub

Private Sub txtQtde_KeyPress(KeyAscii As Integer)

Tweak txtQtde, KeyAscii, IntegerPositive

End Sub

Private Sub cmdGerar_Click()
If Val(txtQtde.Text) = 0 Or Val(txtQtde.Text) > 30 Then
    MsgBox "Quantidade de guias deve estar entre 1 e 30", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Confirma criação da(s) Guia(s) ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   Exit Sub
End If

GerarGuia

End Sub

Private Sub GerarGuia()

Dim nCodReduz As Long, Sql As String, RdoAux As rdoResultset, nQtde As Integer, dDataVencto As Date
Dim sNome As String, nValorExp As Double, sCNPJ As String, nValorParcela As Double, nCodLanc As Integer
Dim nCodTrib As Integer, sTipoImposto As String, sEndereco As String, sBairro As String, sCidade As String
Dim sCep As String, sUF As String, nLastCod As Long, sNumProc As String, dDataProc As Date, nCount As Integer
Dim NumBarra1 As String, NumBarra2 As String, NumBarra2a As String, NumBarra2b As String, nNumImovel As Integer
Dim NumBarra2c As String, NumBarra2d As String, StrBarra1 As String, StrBarra2 As String, sDescImposto As String
Dim nLastDoc As Long, nSeqLanc As Integer, sObs As String, nSeq As Integer, nExercicio As Integer, nFirstDoc As Long

MsgBox "Bloqueado"
Exit Sub


nCodReduz = 577831
nQtde = Val(txtQtde.Text)
dDataVencto = DateAdd("d", 30, Now)
dDataVencto = RetornaDiaUtil(dDataVencto)
nValorParcela = CDbl(cmbValor.Text)
nCodLanc = 11
nCodTrib = 505
nExercicio = Year(Now)
sTipoImposto = "TAXAS DIVERSAS"
sDescImposto = "TX.FUNC.AMBULANTE"
sObs = "Taxa de funcionamento para comércio ambulante."
sEndereco = "ESPLANADA DO LAGO"
nNumImovel = 160
sBairro = "VILA SERRA"
sCidade = "JABOTICABAL"
sCep = "14870-200"
sUF = "SP"
sNumProc = Format(nCodReduz, "000000") & "/" & CStr(Year(Now))
dDataProc = Format(Now, "dd/mm/yyyy")
NumBarra1 = Format(ExtraiNumero(sNumProc), "0000000000")
StrBarra1 = Gera2of5Str(NumBarra1)

Sql = "select nomecidadao,cnpj from cidadao where codcidadao=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNome = Mask(!nomecidadao)
        sCNPJ = Format(!Cnpj, "0#\.###\.###/####-##")
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT VALORPARCELA FROM EXPEDIENTE WHERE ANOEXPED=" & nExercicio & " AND CODLANCAMENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount > 0 Then
        nValorExp = FormatNumber(!VALORPARCELA, 2)
     Else
        MsgBox "Taxa de Expediente não cadastrada.", vbCritical, "Atenção"
        Exit Sub
     End If
    .Close
End With

Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

For nCount = 1 To nQtde
    Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    nLastDoc = RdoAux!maximo + 1
    RdoAux.Close
    
    If nCount = 1 Then
        nFirstDoc = nLastDoc
    End If

    Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nExercicio & " AND "
    Sql = Sql & "CODLANCAMENTO=" & nCodLanc & " AND NUMPARCELA=1 AND CODCOMPLEMENTO=0"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeqLanc = 0
    Else
        nSeqLanc = RdoAux!maximo + 1
    End If
    RdoAux.Close

   'DEBITOPARCELA
'    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USUARIO) VALUES(" & nCodReduz & "," & nExercicio & "," & nCodLanc & ","
'    Sql = Sql & nSeqLanc & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
'    Sql = Sql & 1 & ",'" & Left$(NomeDeLogin, 25) & "')"
    Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
    Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,USERID) VALUES(" & nCodReduz & "," & nExercicio & "," & nCodLanc & ","
    Sql = Sql & nSeqLanc & "," & 1 & "," & 0 & "," & 3 & ",'" & Format(dDataVencto, "mm/dd/yyyy") & "','" & Format(Right$(frmMdi.Sbar.Panels(6).Text, 10), "mm/dd/yyyy") & "',"
    Sql = Sql & 1 & "," & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect

   'OBS PARCELA
'    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduz & ","
'    Sql = Sql & nExercicio & "," & nCodLanc & "," & nSeqLanc & "," & 1 & "," & 0 & "," & 0 & ",'" & sObs & "','" & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
    Sql = "INSERT OBSPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USERID,DATA) VALUES(" & nCodReduz & ","
    Sql = Sql & nExercicio & "," & nCodLanc & "," & nSeqLanc & "," & 1 & "," & 0 & "," & 0 & ",'" & sObs & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Format(Now, "mm/dd/yyyy") & "')"
    cn.Execute Sql, rdExecDirect

   'GRAVA DEBITOTRIBUTO
    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
    Sql = Sql & "CODTRIBUTO,VALORTRIBUTO) VALUES(" & nCodReduz & "," & nExercicio & "," & nCodLanc & ","
    Sql = Sql & nSeqLanc & "," & 1 & "," & 0 & "," & nCodTrib & "," & Virg2Ponto(CStr(nValorParcela)) & ")"
    cn.Execute Sql, rdExecDirect

   'GRAVA DOCUMENTO
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,emissor) VALUES("
    Sql = Sql & nLastDoc & ",'" & Format(dDataProc, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(nValorExp)) & ",'" & NomeDeLogin & " (GUIA AMBULANTE)" & "')"
    cn.Execute Sql, rdExecDirect
         
   'GRAVA PARCELADOCUMENTO
    Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,"
    Sql = Sql & "CODCOMPLEMENTO,NUMDOCUMENTO) VALUES(" & nCodReduz & "," & nExercicio & "," & nCodLanc & ","
    Sql = Sql & nSeqLanc & "," & 1 & "," & 0 & "," & nLastDoc & ")"
    cn.Execute Sql, rdExecDirect

   'CÓDIGO DE BARRAS
    'NumBarra2 = Gera2of5Cod(nValorParcela + nValorExp, dDataVencto, nLastDoc, 1, nCodLanc, nSeqLanc, 0)
        
    NumBarra2a = Left$(NumBarra2, 13)
    NumBarra2b = Mid$(NumBarra2, 14, 13)
    NumBarra2c = Mid$(NumBarra2, 27, 13)
    NumBarra2d = Right$(NumBarra2, 13)
    StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))

   'GRAVA CARNETMP
    Sql = "INSERT CARNETMP(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,ENDIMOVEL,NUMIMOVEL,COMPLIMOVEL,"
    Sql = Sql & "BAIRROIMOVEL,ENDENTREGA,NUMENTREGA,COMPLENTREGA,BAIRROENTREGA,CEPENTREGA,CIDADEENTREGA,UFENTREGA,"
    Sql = Sql & "DESCIMPOSTO,EXERCICIO,NUMPROCESSO,DATAPROCESSO,NUMDOCUMENTO,DV,QUADRA,LOTE,DATAVENCTO,NUMPARCELA,"
    Sql = Sql & "NUMTOTPARCELA,VALORPARCELA,STRBARRA1,STRBARRA2,NUMBARRA1,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
    Sql = Sql & "DADOSLANCAMENTO,TAXAEXP,SAIR,OBS) VALUES('" & NomeDoUsuario & "'," & nCount & ",'" & "" & "','" & CStr(nCodReduz) & "','"
    Sql = Sql & Left$(sTipoImposto, 15) & "','" & Mask(Left$(sNome, 40)) & "','" & Left$(sEndereco, 40) & "'," & nNumImovel & ",'" & "" & "','"
    Sql = Sql & Left(sBairro, 25) & "','" & Left(sEndereco, 40) & "'," & nNumImovel & ",'" & "" & "','" & Left(sBairro, 25) & "','"
    Sql = Sql & sCep & "','" & sCidade & "','" & sUF & "','" & sDescImposto & "'," & nExercicio & ",'" & sNumProc & "','"
    Sql = Sql & Format(dDataProc, "mm/dd/yyyy") & "','" & CStr(nLastDoc) & "','" & CStr(RetornaDVNumDoc(nLastDoc)) & "','" & "" & "','"
    Sql = Sql & "" & "','" & Format(dDataVencto, "mm/dd/yyyy") & "'," & 1 & "," & 1 & ","
    Sql = Sql & Virg2Ponto(RemovePonto(CStr(nValorParcela))) & ",'" & Mask(StrBarra1) & "','" & Mask(StrBarra2) & "'," & NumBarra1 & ",'" & NumBarra2a & "','"
    Sql = Sql & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "','" & sDescImposto & "'," & Virg2Ponto(CStr(nValorExp)) & "," & "0" & ",'" & Mask(Trim(sObs)) & "')"
    cn.Execute Sql, rdExecDirect

Next

frmReport.ShowReport2 "CARNEINDIVIDUAL", frmMdi.hwnd, Me.hwnd, nFirstDoc
Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect


End Sub
