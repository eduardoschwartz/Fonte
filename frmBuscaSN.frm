VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmBuscaSN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscas nos Arq. do Simples Nacional"
   ClientHeight    =   6165
   ClientLeft      =   3435
   ClientTop       =   2565
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   11640
   Begin VB.TextBox txtAno2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      MaxLength       =   4
      TabIndex        =   2
      Top             =   90
      Width           =   795
   End
   Begin VB.TextBox txtAno1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3840
      MaxLength       =   4
      TabIndex        =   1
      Top             =   90
      Width           =   795
   End
   Begin VB.ListBox lstArq 
      Height          =   450
      Left            =   12750
      TabIndex        =   11
      Top             =   435
      Width           =   1050
   End
   Begin VB.ListBox lstPath 
      Height          =   450
      Left            =   12810
      TabIndex        =   10
      Top             =   960
      Width           =   1050
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   12660
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "\\192.168.200.130\atualizagti\SimplesNacional\"
      Top             =   1650
      Width           =   2205
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   60
      TabIndex        =   6
      Top             =   480
      Width           =   5415
      _ExtentX        =   9551
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
      Color           =   8421376
   End
   Begin prjChameleon.chameleonButton cmdRebuild 
      Height          =   315
      Left            =   7740
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Reconstruir a lista de arquivos bancários"
      Top             =   60
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Rebuild"
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
      MICON           =   "frmBuscaSN.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdFind 
      Height          =   315
      Left            =   6630
      TabIndex        =   3
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   60
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "B&usca"
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
      MICON           =   "frmBuscaSN.frx":001C
      PICN            =   "frmBuscaSN.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskCNPJ 
      Height          =   285
      Left            =   795
      TabIndex        =   0
      Top             =   90
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   503
      MouseIcon       =   "frmBuscaSN.frx":0192
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
      MaxLength       =   18
      Mask            =   "99.999.999/9999-99"
      SelText         =   ""
      Text            =   "__.___.___/____-__"
      HideSelection   =   -1  'True
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   5325
      Left            =   90
      TabIndex        =   4
      Top             =   780
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9393
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   5186
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "CNPJ"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt.Arr."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Dt.Venc."
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Anocp"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Mescp"
         Object.Width           =   1306
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Valor"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Fim.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   4770
      TabIndex        =   13
      Top             =   135
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano Início:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   2820
      TabIndex        =   12
      Top             =   135
      Width           =   1065
   End
   Begin VB.Label lblMsg 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5610
      TabIndex        =   9
      Top             =   450
      Width           =   4860
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   75
      TabIndex        =   8
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmBuscaSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type SimplesNacionalHeader
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    CodigoConvenio As String * 20
    DataGeracao As String * 8
    NumeroRemessa As String * 6
    NumeroVersao As String * 2
    Filler1 As String * 22
    Filler2 As String * 8
    CodigoBanco As String * 3
    Filler3 As String * 422
End Type
Private Type SimplesNacionalDetalhe
    CodigoRegistro As String * 1
    SeqRegistro As String * 8
    DataArrecada As String * 8
    DataVencimento As String * 8
    Filler1 As String * 12
    Filler2 As String * 37
    Cnpj As String * 14
    Filler3 As String * 11
    Esfera As String * 1
    Competencia As String * 6
    ValorPrincipal As String * 17
    ValorMulta As String * 17
    ValorJuros As String * 17
    Filler4 As String * 47
    ValorAutentica As String * 17
    NumeroAutentica As String * 23
    CodigoBanco As String * 3
    CodigoAgencia As String * 4
    Filler6 As String * 249
End Type

Private Sub chkAno_Click()
If chkAno.value = vbChecked Then
    cmdAno.Enabled = True
Else
    cmdAno.value = False
    cmdAno.Enabled = False
    frAno.Height = 375
End If
End Sub

Private Sub cmdAno_Click()
Dim nAno As Integer, x As Integer, y As Integer

If cmdAno.value = True Then
    frAno.Height = 1410
Else
    frAno.Height = 375
End If

End Sub

Private Sub cmdFind_Click()

Dim sPathFile As String, sNomeArq As String, strLinha As String, x As Long, Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim strBuffer As String, lngResult As Long, dDataVencto As Date, z As Long, FF1 As Integer, sCNPJ As String
Dim nPosicao As Integer, nCount As Integer, sRetorno As String, sDataVencto As String, sReg As String
Dim nCodBanco As Integer, sNumRemessa As String, sDataGeracao As String, sDataArrecada As String, sAgencia As String
Dim nAno As Integer, nMes As Integer, nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double
Dim nCodReduz As Long, nExercicio As Integer, nParc As Integer, nSeq As Integer, sRazao As String
Dim bFilterCNPJ As Boolean, bFilterAno As Boolean, aAno() As Long, nCodCidadao As Long, sNomeCidadao As String
Dim Itmx As ListItem
ReDim aAno(0)

If mskCNPJ.ClipText <> "" Then
    If Not ValidaCGC(mskCNPJ.ClipText) Then
        MsgBox "Número de CNPJ inválido.", vbCritical, "Verifique"
        Exit Sub
    End If
End If

If Val(txtAno1.Text) < 2008 Then
    MsgBox "Ano inicial tem que ser a partir de 2008.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno2.Text) < 2008 Or Val(txtAno2.Text) > Year(Now) Then
    MsgBox "Ano final inválido.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno2.Text) < Val(txtAno1.Text) Then
    MsgBox "Ano final inválido.", vbExclamation, "Atenção"
    Exit Sub
End If

For x = Val(txtAno1.Text) To Val(txtAno2.Text)
    ReDim Preserve aAno(UBound(aAno) + 1)
    aAno(UBound(aAno)) = x
    bFilterAno = True
Next
If lstPath.ListCount = 0 Then
    LoadFileBanco
End If

If lstPath.ListCount = 0 Then
    MsgBox "Lista " & sFile & " não localizada.", vbCritical, "Verifique"
    Exit Sub
End If

If MsgBox("Confirme a operação?", vbYesNo + vbQuestion, "Atenção") = vbNo Then Exit Sub
lvMain.ListItems.Clear
bFilterCNPJ = False: bFilterAno = False

If mskCNPJ.ClipText <> "" Then bFilterCNPJ = True
nCount = 0

If Not bFilterCNPJ And NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "NOP"
    Exit Sub
End If



'Sql = "DELETE FROM RESUMOARQSN WHERE USUARIO='" & NomeDeLogin & "'"
'cn.Execute Sql, rdExecDirect

Ocupado
Me.MousePointer = vbHourglass
If bFilterCNPJ Then
    lblMsg.Caption = "Procurando o CNPJ..."
Else
    lblMsg.Caption = "Listando os arquivos..."
End If
lblMsg.Refresh
DoEvents

lstArq.Clear
PBar.Color = vbRed
For x = 0 To lstPath.ListCount - 1
    If lstPath.ItemData(x) = 1 Then GoTo PROXIMO
    If x Mod 20 = 0 Then
        CallPb x, CLng(lstPath.ListCount - 1)
    End If
    sReg = ""
    sPathFile = lstPath.List(x)
    sNomeArq = ParsePath(sPathFile, vbNormal)
    
    Set cCRCSearch = New cFileSearchCRC
    cCRCSearch.SearchAlgorithm = Asm_BMHA
    strBuffer = sTr$(cCRCSearch.FileMapSearch(sPathFile, "SIMPLES"))
    If Val(strBuffer) = 0 Then GoTo PROXIMO
    
    If Not bFilterCNPJ Then GoTo POINT1
    Set cCRCSearch = New cFileSearchCRC
    cCRCSearch.SearchAlgorithm = Asm_BMHA
    strBuffer = sTr$(cCRCSearch.FileMapSearch(sPathFile, mskCNPJ.ClipText))
    Set cCRCSearch = Nothing
    If Val(strBuffer) > 0 Then
POINT1:
        lstArq.AddItem sPathFile
        'LE ARQUIVO ENCONTRADO
        FF1 = FreeFile()
        Open sPathFile For Binary Access Read Write As FF1
           ' sReg = ""
            While Not EOF(FF1)
                
                On Error GoTo CloseFile2
                If Left(sReg, 1) = "9" Or Left(sReg, 1) = "Z" Then GoTo CloseFile2
                Input #FF1, sReg
                If sReg = "" Then GoTo CloseFile2
                If Left(sReg, 1) = "1" Then
                    sSeqArq = Mid(sReg, 2, 8)
                    
                    sNumRemessa = Mid(sReg, 38, 6)
                    sDataGeracao = ConvDataSerial(Mid(sReg, 30, 8))
                ElseIf Left(sReg, 1) = "2" Then
                   'LE OS REGISTROS
                    With grdReg
                        nValorPrincipal = CDbl(Mid(sReg, 107, 17)) / 100
                        nValorJuros = CDbl(Mid(sReg, 124, 17)) / 100
                        nValorMulta = CDbl(Mid(sReg, 141, 17)) / 100
                        nValorGuia = nValorPrincipal + nValorJuros + nValorMulta
                        nAno = Val(Mid(sReg, 101, 4))
                        nMes = Val(Mid(sReg, 105, 2))
                        nCodBanco = Val(Mid(sReg, 245, 3))
                        sAgencia = Mid(sReg, 248, 4)
                        sCNPJ = Mid(sReg, 75, 14)
                        If Val(sCNPJ) <> Val(mskCNPJ.ClipText) And bFilterCNPJ Then GoTo PROXIMOREG
                        
                        sDataVencto = ConvDataSerial(Mid(sReg, 18, 8))
                        sDataArrecada = ConvDataSerial(Mid(sReg, 10, 8))
                        If bFilterAno Then
                            z = BinarySearchLong(aAno, CLng(Right(sDataArrecada, 4)))
                            If z = -1 Then GoTo PROXIMOREG
                        End If
                        sRazao = "": sNomeCidadao = "": nCodCidadao = 0: nCodReduz = 0
                        Sql = "SELECT CODIGOMOB,RAZAOSOCIAL FROM MOBILIARIO WHERE CNPJ='" & sCNPJ & "'"
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux
                            If .RowCount > 0 Then
                                nCodReduz = !codigomob
                                sRazao = !RazaoSocial
                            Else
                                Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & sCNPJ & "'"
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    If .RowCount > 0 Then
                                        nCodCidadao = !CodCidadao
                                        sNomeCidadao = !nomecidadao
                                    Else
                                        sRazao = "CNPJ NÃO LOCALIZADO"
                                    End If
                                   .Close
                                End With
                            End If
                           .Close
                        End With
                        
                        'PROCURA SE O DEBITO JA FOI BAIXADO
                        Sql = "SELECT * FROM COMPLEMENTOSIMPLES WHERE ARQUIVOBANCO='" & sNomeArq & "' AND DATACREDITO='" & Format(ConvDataSerial(Mid(sReg, 10, 8)), "mm/dd/yyyy") & "' AND "
                        Sql = Sql & "CNPJ='" & Mid(sReg, 75, 14) & "' AND ANO=" & Val(Mid(sReg, 101, 4)) & " AND MES=" & Val(Mid(sReg, 105, 2))
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux
                            If .RowCount > 0 Then
                                'CARREGA PARCELA GRAVADA
                                nCodReduz = !CODREDUZIDO
                                nExercicio = !AnoExercicio
                                nSeq = !SeqLancamento
                                nParc = !NumParcela
                            End If
                           .Close
                        End With
                    End With
                    
                    'GRAVA NA TABELA
                    On Error GoTo 0
                    Set Itmx = lvMain.ListItems.Add(, , IIf(IsNull(nCodReduz), SubNull(nCodCidadao), nCodReduz))
                    Itmx.SubItems(1) = IIf(Not IsNull(sNomeCidadao), sRazao, sNomeCidadao)
                    Itmx.SubItems(2) = Format(sCNPJ, "0#\.###\.###/####-##")
                    Itmx.SubItems(3) = Format(sDataArrecada, "dd/mm/yyyy")
                    Itmx.SubItems(4) = Format(sDataVencto, "dd/mm/yyyy")
                    Itmx.SubItems(5) = nAno
                    Itmx.SubItems(6) = Format(nMes, "00")
                    Itmx.SubItems(7) = FormatNumber(nValorPrincipal + nValorJuros + nValorMulta, 2)
'                    Sql = "SELECT * FROM RESUMOARQSN WHERE NUMREMESSA='" & sNumRemessa & "' AND "
'                    Sql = Sql & "CNPJ='" & Format(sCNPJ, "0#\.###\.###/####-##") & "' AND ANOCOMP=" & nAno & " AND MESCOMP=" & nMes
'                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                    With RdoAux
'                        If .RowCount = 0 Then
'                            DoEvents
'                            If nCodReduz = 0 And nCodCidadao = 0 And sRazao <> "CNPJ NÃO LOCALIZADO" Then
'                                MsgBox "TESTE"
'                            End If
'                            Sql = "INSERT RESUMOARQSN(USUARIO,ARQUIVOLONG,ARQUIVOSHORT,BANCO,DATAGERACAO,NUMREMESSA,CNPJ,NOME,DATAARRECADA,DATAVENCTO,"
'                            Sql = Sql & "ANOCOMP,MESCOMP,PRINCIPAL,JUROS,MULTA,AGENCIA,CODREDUZ,ANO,SEQ,PARC,CODCIDADAO,NOMECIDADAO) VALUES('"
'                            Sql = Sql & NomeDeLogin & "','" & sPathFile & "','" & Left(sNomeArq, 25) & "'," & nCodBanco & ",'" & Format(sDataGeracao, "mm/dd/yyyy") & "','"
'                            Sql = Sql & sNumRemessa & "','" & Format(sCNPJ, "0#\.###\.###/####-##") & "','" & Left(Mask(sRazao), 50) & "','" & Format(sDataArrecada, "mm/dd/yyyy") & "','" & Format(sDataVencto, "mm/dd/yyyy") & "',"
'                            Sql = Sql & nAno & "," & nMes & "," & Virg2Ponto(CStr(nValorPrincipal)) & "," & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & ","
 '                           Sql = Sql & Val(sAgencia) & "," & nCodReduz & "," & nExercicio & "," & nSeq & "," & nParc & "," & nCodCidadao & ",'" & Mask(sNomeCidadao) & "')"
 '                           cn.Execute Sql, rdExecDirect
 '                       End If
 '                   End With
                    nCodReduz = 0: nExercicio = 0: nSeq = 0: nParc = 0
                ElseIf Left(sReg, 1) = "9" Then
                    GoTo CloseFile2
                End If
PROXIMOREG:
                nPos = nPos + 1
            Wend
CloseFile2:
        Close #FF1
    End If
PROXIMO:
Next

If lstArq.ListCount = 0 Then
    lblMsg.Caption = "CNPJ não encontrado !!!"
    lblMsg.Refresh
Else
    lblMsg.Caption = "Busca finalizada!"
    lblMsg.Refresh
  '  Sql = "SELECT * FROM RESUMOARQSN WHERE usuario='" & NomeDeLogin & "' order by datavencto "
  '  Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
  '  With RdoAux
  '      Do Until .EOF
  '          Set Itmx = lvMain.ListItems.Add(, , IIf(IsNull(!CodReduz), SubNull(!CodCidadao), !CodReduz))
  '          Itmx.SubItems(1) = IIf(IsNull(!nome), SubNull(!nomecidadao), !nome)
  '          Itmx.SubItems(2) = !Cnpj
  '          Itmx.SubItems(3) = Format(!DataArrecada, "dd/mm/yyyy")
  '          Itmx.SubItems(4) = Format(!DataVencto, "dd/mm/yyyy")
  '          Itmx.SubItems(5) = !anocomp
  '          Itmx.SubItems(6) = !mescomp
  '          Itmx.SubItems(7) = FormatNumber(!principal + !Juros + !Multa, 2)
  '         .MoveNext
  '      Loop
  '     .Close
  '  End With


End If

'Sql = "DELETE FROM RESUMOARQSN WHERE USUARIO='" & NomeDeLogin & "'"
'cn.Execute Sql, rdExecDirect

PBar.value = 0
PBar.Color = vbWhite
Me.MousePointer = vbDefault
Liberado
    
End Sub

Private Sub Form_Load()
Dim x As Integer

Centraliza Me
If NomeDeLogin <> "SCHWARTZ" Then cmdRebuild.Enabled = False


End Sub

Private Sub cmdRebuild_Click()
    
Dim sDirTemp As String
Dim iFilesFile As Integer
Dim iDirsFile As Integer
Dim iStart As Long
Set gsDirsQueue = Nothing
Set gsDirs = Nothing
Set gsFiles = Nothing

If NomeDeLogin <> "SCHWARTZ" Then
    MsgBox "Sem permissão.", vbCritical, "ERRO"
    Exit Sub
End If

If MsgBox("Deseja criar\atualizar a lista de arquivos?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

If txtPath.Text = "" Then Exit Sub

If Right(txtPath.Text, 1) <> "\" Then
    txtPath.Text = txtPath.Text & "\"
End If
Ocupado
DoEvents

On Error Resume Next
Kill App.Path & "\Bin\SNac.txt"
On Error GoTo 0
iFilesFile = FreeFile
'Open App.Path & "\Bin\SNac.txt" For Output Access Write As iFilesFile
Open txtPath.Text & "Layout\SNac.txt" For Output Access Write As iFilesFile

gsDirsQueue.Add txtPath.Text
gsDirs.Add txtPath.Text

While gsDirsQueue.Count > 0
    DoEvents
    sDirTemp = FixPath(gsDirsQueue(1))
    sTemp = ""
    On Error Resume Next
    sTemp = Dir$(sDirTemp & "\*.*", vbNormal + vbReadOnly + vbHidden + vbDirectory)
    If sTemp = "" Then
        Liberado
        MsgBox "Diretório: " & sDirTemp & " não encontrado.", vbCritical, "ERRO"
        Close iFilesFile
        Exit Sub
    End If
    On Error GoTo 0
    While sTemp <> ""
        If sTemp <> "." And sTemp <> ".." Then
            If (GetAttr(sDirTemp & "\" & sTemp) And &H10) = vbDirectory Then
                gsDirsQueue.Add Item:=sDirTemp & "\" & sTemp, After:=DirSearchB(gsDirsQueue, sDirTemp & "\" & sTemp) ' Comes and goes here.
                gsDirs.Add Item:=sDirTemp & "\" & sTemp, After:=DirSearchB(gsDirs, sDirTemp & "\" & sTemp) ' Comes and goes here.
            Else  ' We have a file.
                If UCase(Left(sTemp, 3)) = "DAF" Then
                    gsFiles.Add sDirTemp & "\" & sTemp
                    Print #iFilesFile, sDirTemp & "\" & sTemp
                End If
            End If
        End If
        sTemp = Dir$()
    Wend

    gsDirsQueue.Remove (1)
Wend

Close
Liberado
MsgBox "Lista dos arquivos do simples nacional foi reconstruida.", vbInformation, "Informação"
    
End Sub

Private Sub mskCNPJ_GotFocus()
mskCNPJ.SelStart = 0
mskCNPJ.SelLength = Len(mskCNPJ.Text)
End Sub

Private Sub LoadFileBanco()

Dim strLinha As String
sFile = txtPath.Text & "Layout\SNac.txt"

If Dir$(sFile) = "" Then
    MsgBox "Lista " & sFile & " não localizada, criando lista.", vbCritical, "Verifique"
    cmdRebuild_Click
    Exit Sub
End If

lstPath.Clear
Ocupado

On Error Resume Next
Close #1
On Error GoTo 0

Open sFile For Input As #1
   Do While Not EOF(1)
        Line Input #1, strLinha
        lstPath.AddItem strLinha
        lstPath.ItemData(lstPath.NewIndex) = 0
   Loop
Close #1
Liberado
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)

If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Function ConvDataSerial(sData As String) As String
If Len(sData) = 8 Then
   ConvDataSerial = Right$(sData, 2) & "/" & Mid$(sData, 5, 2) & "/" & Left$(sData, 4)
Else
   ConvDataSerial = Left$(sData, 2) & "/" & Mid$(sData, 3, 2) & "/20" & Right$(sData, 2)
End If
End Function



Private Sub txtAno1_KeyPress(KeyAscii As Integer)
Tweak txtAno1, KeyAscii, IntegerPositive
End Sub



Private Sub txtAno2_KeyPress(KeyAscii As Integer)
Tweak txtAno2, KeyAscii, IntegerPositive
End Sub
