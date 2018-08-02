VERSION 5.00
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmBuscaSN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscas no Simples Nacional"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2130
   ScaleWidth      =   5835
   Begin VB.CheckBox chkNome 
      Appearance      =   0  'Flat
      Caption         =   "Exibir nome dos arquivos no relatório."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   14
      Top             =   1080
      Value           =   1  'Checked
      Width           =   3570
   End
   Begin VB.Frame frAno 
      BackColor       =   &H00EEEEEE&
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   540
      Width           =   1230
      Begin VB.CheckBox chkAno 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   0
         Width           =   195
      End
      Begin VB.ListBox lstAno 
         Height          =   960
         Left            =   45
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   405
         Width           =   1140
      End
   Begin VB.ListBox lstArq 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   45
      TabIndex        =   8
      Top             =   2610
      Width           =   5685
   End
   Begin VB.ListBox lstPath 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   0
      TabIndex        =   6
      Top             =   3195
      Width           =   4695
   End
   Begin VB.TextBox txtPath 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   135
      TabIndex        =   2
      Text            =   "C:\Trabalho\GTI\Bancos\Arquivos Simples Nacional\"
      Top             =   180
      Width           =   4335
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   1440
      Width           =   4380
      _ExtentX        =   7726
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
   End
   Begin esMaskEdit.esMaskedEdit mskCNPJ 
      Height          =   285
      Left            =   885
      TabIndex        =   0
      Top             =   630
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      MouseIcon       =   "frmBuscaSN.frx":0308
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Anos.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2610
      TabIndex        =   12
      Top             =   675
      Width           =   600
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
      Height          =   315
      Left            =   90
      TabIndex        =   7
      Top             =   1755
      Width           =   5670
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   3
      Left            =   135
      TabIndex        =   5
      Top             =   690
      Width           =   645
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
If chkAno.Value = vbChecked Then
    cmdAno.Enabled = True
Else
    cmdAno.Value = False
    cmdAno.Enabled = False
    frAno.Height = 375
End If
End Sub

Private Sub cmdAno_Click()
Dim nAno As Integer, x As Integer, Y As Integer

If cmdAno.Value = True Then
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

frAno.Height = 375: cmdAno.Value = False

If mskCNPJ.ClipText <> "" Then
    If Not ValidaCGC(mskCNPJ.ClipText) Then
        MsgBox "Número de CNPJ inválido.", vbCritical, "Verifique"
        Exit Sub
    End If
End If

If lstPath.ListCount = 0 Then
    LoadFileBanco
End If

If lstPath.ListCount = 0 Then
    MsgBox "Lista " & sFile & " não localizada.", vbCritical, "Verifique"
    Exit Sub
End If

If MsgBox("Confirme a operação?", vbYesNo + vbQuestion, "Atenção") = vbNo Then Exit Sub

bFilterCNPJ = False: ReDim aAno(0): bFilterAno = False

If mskCNPJ.ClipText <> "" Then bFilterCNPJ = True
nCount = 0

If chkAno.Value = vbChecked Then
    For x = 0 To lstAno.ListCount - 1
        If lstAno.Selected(x) = True Then
            nCount = 1
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = lstAno.List(x)
            bFilterAno = True
        End If
    Next
    If nCount = 0 Then
        MsgBox "Selecione ao menos um ano.", vbCritical, "ERRO"
        Exit Sub
    End If
End If

Sql = "DELETE FROM RESUMOARQSN WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

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
                        Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CNPJ='" & sCNPJ & "'"
                        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux
                            If .RowCount > 0 Then
                                sRazao = !RAZAOSOCIAL
                            Else
                                Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CNPJ='" & sCNPJ & "'"
                                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                                With RdoAux2
                                    If .RowCount > 0 Then
                                        nCodCidadao = !CodCidadao
                                        sNomeCidadao = !NOMECIDADAO
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
                    
                    Sql = "SELECT * FROM RESUMOARQSN WHERE NUMREMESSA='" & sNumRemessa & "' AND "
                    Sql = Sql & "CNPJ='" & Format(sCNPJ, "0#\.###\.###/####-##") & "' AND ANOCOMP=" & nAno & " AND MESCOMP=" & nMes
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                        If .RowCount = 0 Then
                            Sql = "INSERT RESUMOARQSN(USUARIO,ARQUIVOLONG,ARQUIVOSHORT,BANCO,DATAGERACAO,NUMREMESSA,CNPJ,NOME,DATAARRECADA,DATAVENCTO,"
                            Sql = Sql & "ANOCOMP,MESCOMP,PRINCIPAL,JUROS,MULTA,AGENCIA,CODREDUZ,ANO,SEQ,PARC,CODCIDADAO,NOMECIDADAO) VALUES('"
                            Sql = Sql & NomeDeLogin & "','" & sPathFile & "','" & Left(sNomeArq, 25) & "'," & nCodBanco & ",'" & Format(sDataGeracao, "mm/dd/yyyy") & "','"
                            Sql = Sql & sNumRemessa & "','" & Format(sCNPJ, "0#\.###\.###/####-##") & "','" & Left(Mask(sRazao), 50) & "','" & Format(sDataArrecada, "mm/dd/yyyy") & "','" & Format(sDataVencto, "mm/dd/yyyy") & "',"
                            Sql = Sql & nAno & "," & nMes & "," & Virg2Ponto(CStr(nValorPrincipal)) & "," & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & ","
                            Sql = Sql & Val(sAgencia) & "," & nCodReduz & "," & nExercicio & "," & nSeq & "," & nParc & "," & nCodCidadao & ",'" & Mask(sNomeCidadao) & "')"
                            cn.Execute Sql, rdExecDirect
                        End If
                    End With
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
    frmReport.ShowReport2 "RESUMOARQSN", frmMdi.hwnd, Me.hwnd
End If

Sql = "DELETE FROM RESUMOARQSN WHERE USUARIO='" & NomeDeLogin & "'"
'cn.Execute Sql, rdExecDirect

PBar.Value = 0
PBar.Color = vbWhite
Me.MousePointer = vbDefault
Liberado
    
End Sub

Private Sub Form_Load()
Dim x As Integer

Centraliza Me

For x = 2007 To Year(Now)
    lstAno.AddItem x
Next

End Sub

Private Sub cmdRebuild_Click()
    
Dim sDirTemp As String
Dim iFilesFile As Integer
Dim iDirsFile As Integer
Dim iStart As Long
Set gsDirsQueue = Nothing
Set gsDirs = Nothing
Set gsFiles = Nothing

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
Open App.Path & "\Bin\SNac.txt" For Output Access Write As iFilesFile

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
                gsFiles.Add sDirTemp & "\" & sTemp
                Print #iFilesFile, sDirTemp & "\" & sTemp
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

sFile = App.Path & "\bin\SNac.txt"

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
   PBar.Value = (nVal * 100) / nTot
Else
   PBar.Value = 100
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

