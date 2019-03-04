VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmParcelamentoBloqueio 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desbloqueamento de Parcelamentos"
   ClientHeight    =   2550
   ClientLeft      =   6150
   ClientTop       =   4185
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2550
   ScaleWidth      =   4920
   Begin Tributacao.XP_ProgressBar Pbar 
      Height          =   240
      Left            =   3285
      TabIndex        =   18
      Top             =   2070
      Width           =   1320
      _ExtentX        =   2328
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
      Color           =   12632064
   End
   Begin VB.ComboBox txtAno 
      Height          =   315
      ItemData        =   "frmParcelamentoBloqueio.frx":0000
      Left            =   1845
      List            =   "frmParcelamentoBloqueio.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   270
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "atualizar"
      Height          =   345
      Left            =   3300
      TabIndex        =   16
      Top             =   2760
      Width           =   1275
   End
   Begin VB.CommandButton cmdCorrige 
      Caption         =   "Corrigir"
      Height          =   345
      Left            =   210
      TabIndex        =   15
      Top             =   2730
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reverter"
      Height          =   345
      Left            =   1740
      TabIndex        =   14
      Top             =   2730
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   795
      Left            =   1680
      TabIndex        =   11
      Top             =   1080
      Width           =   3045
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1620
         MaxLength       =   6
         TabIndex        =   3
         Top             =   150
         Width           =   1275
      End
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   1620
         TabIndex        =   4
         Top             =   450
         Width           =   1275
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido...:"
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   13
         Top             =   180
         Width           =   1485
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo......:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   480
         Width           =   1485
      End
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Geral"
      Height          =   255
      Index           =   1
      Left            =   300
      TabIndex        =   2
      Top             =   1470
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Individual"
      Height          =   255
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.TextBox txtNum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   2070
      Width           =   1125
   End
   Begin VB.TextBox txtPerc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   1125
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   3390
      TabIndex        =   6
      ToolTipText     =   "Sair da Tela"
      Top             =   630
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
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
      MICON           =   "frmParcelamentoBloqueio.frx":0004
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCalculo 
      Height          =   315
      Left            =   3420
      TabIndex        =   5
      ToolTipText     =   "Cancelar Edição"
      Top             =   225
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Executar"
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmParcelamentoBloqueio.frx":0020
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
      Caption         =   "No de Parcelamen..:"
      Height          =   225
      Index           =   2
      Left            =   270
      TabIndex        =   10
      Top             =   2100
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentual %..........:"
      Height          =   225
      Index           =   1
      Left            =   270
      TabIndex        =   8
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano de Bloqueio.....:"
      Height          =   225
      Index           =   0
      Left            =   270
      TabIndex        =   7
      Top             =   330
      Width           =   1515
   End
End
Attribute VB_Name = "frmParcelamentoBloqueio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Debito
    nCodReduzido As Long
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSituacao As Integer
    sSituacao As String
    sVencto As String
    sDA As String
    sAj As String
    nCodTributo As Double
    nValorTributo As Double
    nValorJuros As Double
    nValorMulta As Double
    nValorCorrecao As Double
    nValorHon As Double
    nSaldo As Double
    sDataPago As String
    nValorPago As Double
    nCodBanco As Integer
    dDataPag As Date
End Type

Private Type Juros
    nCodReduzido As Long
    nAno As Integer
    nLanc As Integer
    sLanc As String
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nSaldo As Double
    nJurosPerc As Double
    nJurosValor As Double
End Type

Dim Sql As String, RdoAux As rdoResultset

Private Sub cmdCalculo_Click()
Dim nPos As Long, sNumProc As String, nAnoproc As Integer, nNumproc As Long, nSeq As Integer, nValorTributo As Double, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, ax As String
Dim nValor587 As Double, nValorTributoOld As Double, RdoAux5 As rdoResultset
'If NomeDeLogin <> "SCHWARTZ" Then Exit Sub

If Val(txtAno.Text) > Year(Now) + 3 Or Val(txtAno.Text) < Year(Now) Then
    MsgBox "Ano inválido.", vbCritical, "Atenção"
    Exit Sub
End If

If Val(txtPerc.Text) < 1 Or Val(txtPerc.Text) > 10 Then
    MsgBox "Percentual inválido", vbCritical, "Atenção"
    Exit Sub
End If

If Opt(0).value And (Val(txtCod.Text) = 0 Or txtNumProc.Text = "") Then
    MsgBox "Digite o Código e o número do processo", vbCritical, "Atenção"
    Exit Sub
End If

If Opt(0).value Then
    If InStr(1, txtNumProc.Text, "/", vbBinaryCompare) > 0 Then
        nNumproc = Val(Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2))
        nAnoproc = Val(Right$(txtNumProc.Text, 4))
        Sql = "SELECT NUMPROC,ANOPROC,DATAREPARC,QTDEPARCELA,NOVO,CANCELADO FROM PROCESSOREPARC  WHERE CODIGORESP=" & Val(txtCod.Text) & " AND NUMPROC=" & nNumproc & " AND ANOPROC=" & nAnoproc
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
                txtNumProc.SetFocus
                Liberado
                Exit Sub
            Else
                If !Cancelado Then
                    Liberado
                    MsgBox "Este processo foi cancelado e não pode ser desbloqueado.", vbExclamation, "Atenção"
                    Exit Sub
                End If
            '    If Year(!datareparc) = Year(Now) Then
            '        Liberado
            '        MsgBox "Não é permitido desbloquear parcelamento para o ano atual, utilize a opção de liberação de carnês.", vbCritical, "ERRO"
            '        Exit Sub
            '    End If
                '*** executa desbloqueio
                If MsgBox("Deseja desbloquear este parcelamento?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
                nNumproc = Left$(txtNumProc.Text, InStr(1, txtNumProc.Text, "/", vbBinaryCompare) - 2)
                nAnoproc = Right$(txtNumProc.Text, 4)
                sNumProc = CStr(nNumproc) & "/" & CStr(nAnoproc)
                Sql = "SELECT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,seqlancamento,NUMPARCELA,CODCOMPLEMENTO From debitoparcela Where CODREDUZIDO=" & Val(txtCod.Text) & " AND ANOEXERCICIO>=" & Val(txtAno.Text) & " AND CODLANCAMENTO=20 AND NUMPROCESSO='" & sNumProc & "'"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    Do Until .EOF
                        If .AbsolutePosition = 1 Then
                            nSeq = !SeqLancamento
                            '**grava na tabela processobloqueio
                            Sql = "INSERT PROCESSOBLOQUEIO(ANO,CODREDUZIDO,NUMPROC,ANOPROC) VALUES(" & Val(txtAno.Text) & "," & Val(txtCod.Text) & "," & nNumproc & "," & nAnoproc & ")"
                            'cn.Execute Sql, rdExecDirect
                            '**atualiza o valor dos débitos
        '                    Sql = "UPDATE debitotributo Set valortributo = valortributo + (valortributo * " & Virg2Ponto(CDbl(txtPerc.text) / 100) & ")  Where (CODREDUZIDO = " & Val(txtCod.text) & ") And (AnoExercicio >= " & Val(txtAno.text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") AND (CodTributo <> 3)"
        '                    cn.Execute Sql, rdExecDirect
                            
                            Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                            Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                            Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                nValorTributo = !soma
                                RdoAux2.Close
                            End With
        
                            Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                            Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                            Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=587"
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                If IsNull(!soma) Then
                                    nValor587 = 0
                                Else
                                    nValor587 = !soma
                                End If
                               .Close
                            End With
        
                            '***atualiza o status***
                            Sql = "UPDATE debitoparcela Set statuslanc=3 Where (CODREDUZIDO = " & Val(txtCod.Text) & ") And (AnoExercicio = " & Val(txtAno.Text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") "
                            cn.Execute Sql, rdExecDirect
'                           .Close
                        End If
                        'On Error Resume Next
                        If nValor587 = 0 Then
                            On Error Resume Next
                            Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                            Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                            Sql = Sql & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & ","
                            Sql = Sql & !NumParcela & "," & !CODCOMPLEMENTO & "," & 587 & "," & Virg2Ponto(CStr(Round((nValorTributo * txtPerc.Text / 100), 2))) & ")"
                            cn.Execute Sql, rdExecDirect
                            On Error GoTo 0
                        Else
                            Sql = "SELECT sum(VALORTRIBUTO) as SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & Val(txtAno.Text) - 1 & " AND "
                            Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & 1 & " AND "
                            Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                If Not IsNull(!soma) Then
                                    nValorTributoOld = !soma
                                Else
                                    nValorTributoOld = 0
                                End If
                                RdoAux2.Close
                            End With
                            If (nValorTributo = nValorTributoOld) Or nValorTributoOld = 0 Then
                                Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=VALORTRIBUTO + " & Virg2Ponto(CStr(Round((nValorTributo * txtPerc.Text / 100), 2))) & " WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                                Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                                Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=587"
                                cn.Execute Sql, rdExecDirect
                            End If
                        End If
                       .MoveNext
                    Loop
                End With
                
            End If
        End With
    Else
        MsgBox "Processo de parcelamento não cadastrado para este código.", vbExclamation, "Atenção"
        txtNumProc.SetFocus
    End If
Else
    Sql = "SELECT DISTINCT codreduzido,numprocesso From dbo.debitoparcela Where (CodLancamento = 20) And  (ANOEXERCICIO=" & Val(txtAno.Text) & ") AND (statuslanc = 18) ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        txtNum.Text = .RowCount
       .Close
    End With
    
    If Val(txtNum.Text) = 0 Then
        MsgBox "Nenhum parcelamento à desbloquear.", vbCritical, "Atenção"
        Exit Sub
    End If
    
    If MsgBox("Deseja desbloquear estes parcelamentos?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
    If cGetInputState() <> 0 Then DoEvents
    PBar.value = 0: nPos = 1
    Sql = "SELECT DISTINCT codreduzido,seqlancamento,numprocesso From dbo.debitoparcela Where (CodLancamento = 20) And  (ANOEXERCICIO=" & Val(txtAno.Text) & ") AND (statuslanc = 18) ORDER BY CODREDUZIDO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            If IsNull(!numprocesso) Then GoTo proximo
            If nPos Mod 10 = 0 Then
                CallPb nPos, CLng(Val(txtNum.Text))
'                MsgBox !CODREDUZIDO
            End If
            nPos = nPos + 1
            '**grava na tabela processobloqueio
            sNumProc = !numprocesso
            If sNumProc = "" Then GoTo proximo
            nNumproc = Val(Left$(sNumProc, Len(sNumProc) - 5))
            nAnoproc = Val(Right$(sNumProc, 4))
            
            Sql = "SELECT CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,seqlancamento,NUMPARCELA,CODCOMPLEMENTO From debitoparcela Where CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO>=" & Val(txtAno.Text) & " AND CODLANCAMENTO=20 AND NUMPROCESSO='" & sNumProc & "'"
            Set RdoAux5 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux5
                Do Until .EOF
                    If .AbsolutePosition = 1 Then
                        nSeq = !SeqLancamento
                        '**grava na tabela processobloqueio
                        On Error Resume Next
                        Sql = "INSERT PROCESSOBLOQUEIO(ANO,CODREDUZIDO,NUMPROC,ANOPROC) VALUES(" & Val(txtAno.Text) & "," & !CODREDUZIDO & "," & nNumproc & "," & nAnoproc & ")"
                        cn.Execute Sql, rdExecDirect
                        On Error GoTo 0
                        '**atualiza o valor dos débitos
    '                    Sql = "UPDATE debitotributo Set valortributo = valortributo + (valortributo * " & Virg2Ponto(CDbl(txtPerc.text) / 100) & ")  Where (CODREDUZIDO = " & Val(txtCod.text) & ") And (AnoExercicio >= " & Val(txtAno.text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") AND (CodTributo <> 3)"
    '                    cn.Execute Sql, rdExecDirect
                        
                        Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                        Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                        Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        With RdoAux2
                            nValorTributo = !soma
                            RdoAux2.Close
                        End With
    
                        '***atualiza o status***
                        Sql = "UPDATE debitoparcela Set statuslanc=3 Where (CODREDUZIDO = " & !CODREDUZIDO & ") And (AnoExercicio = " & Val(txtAno.Text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") "
                        cn.Execute Sql, rdExecDirect
                    End If
                    Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND "
                    Sql = Sql & "NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND  CODTRIBUTO=587"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
                    If RdoAux2.RowCount = 0 Then
                        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                        Sql = Sql & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & ","
                        Sql = Sql & !NumParcela & "," & !CODCOMPLEMENTO & "," & 587 & "," & Virg2Ponto(CStr(Round((nValorTributo * CDbl(txtPerc.Text) / 100), 2))) & ")"
                    Else
                        Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=VALORTRIBUTO + " & Virg2Ponto(CStr((Round(nValorTributo * CDbl(txtPerc.Text) / 100, 2)))) & " WHERE CODREDUZIDO=" & RdoAux2!CODREDUZIDO & " AND "
                        Sql = Sql & "ANOEXERCICIO=" & RdoAux2!AnoExercicio & " AND CODLANCAMENTO=" & RdoAux2!CodLancamento & " AND SEQLANCAMENTO=" & RdoAux2!SeqLancamento & " AND NUMPARCELA=" & RdoAux2!NumParcela & " AND "
                        Sql = Sql & "CODCOMPLEMENTO=" & RdoAux2!CODCOMPLEMENTO & " AND CODTRIBUTO=587"
                    End If
                    cn.Execute Sql, rdExecDirect
                    DoEvents
                   .MoveNext
                Loop
            End With
            
            
'            Sql = "INSERT PROCESSOBLOQUEIO(ANO,CODREDUZIDO,NUMPROC,ANOPROC) VALUES(" & Val(txtAno.text) & "," & !CODREDUZIDO & "," & nNumProc & "," & nAnoProc & ")"
'            cn.Execute Sql, rdExecDirect
            '**atualiza o valor dos débitos
'            Sql = "UPDATE debitotributo Set valortributo = valortributo + (valortributo * " & Virg2Ponto(CDbl(txtPerc.text) / 100) & ")  Where (CODREDUZIDO = " & !CODREDUZIDO & ") And (AnoExercicio >= " & Val(txtAno.text) & ") And (CodLancamento = 20) And (SeqLancamento = " & !SeqLancamento & ") AND (CodTributo <> 3)"
'            cn.Execute Sql, rdExecDirect
            
            '***atualiza o status***
'            Sql = "UPDATE debitoparcela Set statuslanc=3 Where (CODREDUZIDO = " & !CODREDUZIDO & ") And (AnoExercicio = " & Val(txtAno.text) & ") And (CodLancamento = 20) And (SeqLancamento = " & !SeqLancamento & ") "
'            cn.Execute Sql, rdExecDirect
            
proximo:
           .MoveNext
        Loop
       .Close
    End With
    PBar.value = 100
End If

'modLg "Acesso ao Sistema"
modLg "Desbloqueou parcelamento Código: " & Val(txtCod.Text) & " Ano: " & Val(txtAno.Text) & " Processo: " & txtNumProc.Text
MsgBox "Parcelamento(s) desbloqueado(s)"
End Sub

Private Sub cmdCorrige_Click()
Dim nCodReduz As Long, nSeq As Integer, sNumProc As String, nNumproc As Long, nAnoproc As Integer
Dim RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim nValorLanc As Double, nSomaPrincipal As Double
Dim nValorJuros As Double, nSomaJuros As Double
Dim nValorMulta As Double, nSomaMulta As Double
Dim nValorCorrecao As Double, nSomaCorrecao As Double
Dim nValorHon As Double, nSomaHonorario As Double
Dim nValorTotal As Double, nEval As Integer, nValorParcela As Double
Dim dDataVencto As Date, nJurosApl As Double, nSomaJurosApl As Double, nJurosValor As Double, nSomaJurosValor As Double
Dim dDataPag As Date, nQtde As Integer, nQtdeJuros As Integer, nJurosPerc As Double
Dim x As Integer
Dim dDataPagto As Date, aJuros() As Juros, bJurosMulta As Boolean
Dim qd As New rdoQuery, aDebito() As Debito, Achou As Boolean

nJurosPerc = 1

Sql = "SELECT DISTINCT debitoparcela.codreduzido, debitoparcela.seqlancamento, destinoreparc.numprocesso FROM  debitoparcela INNER JOIN "
Sql = Sql & "destinoreparc ON debitoparcela.codreduzido = destinoreparc.codreduzido AND debitoparcela.anoexercicio = destinoreparc.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = destinoreparc.codlancamento AND debitoparcela.seqlancamento = destinoreparc.numsequencia AND "
Sql = Sql & "debitoparcela.NumParcela = destinoreparc.NumParcela And debitoparcela.CODCOMPLEMENTO = destinoreparc.CODCOMPLEMENTO "
'Sql = Sql & "Where (debitoparcela.AnoExercicio = 2007) And (debitoparcela.CodLancamento = 20) And (debitoparcela.statuslanc = 18)"
Sql = Sql & "Where (debitoparcela.AnoExercicio > 2007) And (debitoparcela.CodLancamento = 20) And (debitoparcela.CODREDUZIDO=110432)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    bJurosMulta = False
    Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=2007 AND CODLANCAMENTO=20"
    Sql = Sql & " AND SEQLANCAMENTO=" & !SeqLancamento & " And CODCOMPLEMENTO = 0 AND CODTRIBUTO=113"
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
        bJurosMulta = True
    End If
    RdoAux2.Close
    
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        nSeq = !SeqLancamento
        sNumProc = !numprocesso
        txtCod.Text = !CODREDUZIDO
        txtNumProc.Text = sNumProc
'        Me.Refresh
        If cGetInputState() <> 0 Then DoEvents
        'VERIFICA SE JA FOI CORRIGIDO
        Sql = "SELECT CODREDUZIDO,VALORPRINCIPAL FROM DESTINOREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2!ValorPrincipal > 0 Then
'            GoTo proximo
        End If
        RdoAux2.Close
        
       'CARREGA DADOS DO PARCELAMENTO
        Sql = "SELECT  NUMPROCESSO,DATAREPARC,QTDEPARCELA FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            dDataPag = Format(!datareparc, "dd/mm/yyyy")
            nQtde = !qtdeparcela
           .Close
        End With
       'CARREGA ORIGEM
        ReDim aDebito(0)
        Set qd = Nothing
'        Sql = "SELECT DISTINCT numprocesso, dataprocesso, datareparc, qtdeparcela, valorentrada, percentrada, calculamulta, calculajuros, codigoresp, funcionario, codreduzido,"
'        Sql = Sql & "AnoExercicio , CodLancamento, numsequencia, NumParcela, CODCOMPLEMENTO, datavencimento, datadebase, numproc, anoproc "
'        Sql = Sql & "FROM vwCNSREPARCELAMENTOO WHERE CODREDUZIDO=" & nCodReduz & "AND NUMPROCESSO='" & sNumProc & "'"
        Sql = "SELECT DISTINCT numprocesso, dataprocesso, datareparc, qtdeparcela, valorentrada, percentrada, calculamulta, calculajuros, codigoresp, nomelogin, codreduzido,"
        Sql = Sql & "AnoExercicio , CodLancamento, numsequencia, NumParcela, CODCOMPLEMENTO, datavencimento, datadebase, numproc, anoproc "
        Sql = Sql & "FROM vwCNSREPARCELAMENTOO WHERE CODREDUZIDO=" & nCodReduz & "AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                'CARREGA OS TRIBUTOS DE CADA UM DOS LANCAMENTOS
                Set qd.ActiveConnection = cn
                On Error Resume Next
                RdoAux3.Close
                On Error GoTo 0
                qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
                qd(0) = !CODREDUZIDO
                qd(1) = !CODREDUZIDO 'codigo
                qd(2) = !AnoExercicio
                qd(3) = !AnoExercicio  'ano
                qd(4) = !CodLancamento
                qd(5) = !CodLancamento 'lancamento
                qd(6) = !numsequencia
                qd(7) = !numsequencia 'sequencia
                qd(8) = !NumParcela
                qd(9) = !NumParcela 'numparcela
                qd(10) = !CODCOMPLEMENTO
                qd(11) = !CODCOMPLEMENTO 'complemento
                qd(12) = 1
                qd(13) = 99 'statuslanc
                qd(14) = Format(dDataPag, "mm/dd/yyyy")
                qd(15) = NomeDoUsuario
                Set RdoAux3 = qd.OpenResultset(rdOpenKeyset)
                With RdoAux3
                    Do Until .EOF
                        'CARREGA MATRIZ DE DÉBITO
                        nEval = UBound(aDebito)
                        Achou = False
                        For x = 1 To nEval
                            If aDebito(x).nCodReduzido = !CODREDUZIDO And aDebito(x).nAno = !AnoExercicio And aDebito(x).nLanc = !CodLancamento And _
                               aDebito(x).nSeq = !SeqLancamento And _
                               aDebito(x).nParc = !NumParcela And aDebito(x).nCompl = !CODCOMPLEMENTO Then
                               Achou = True
                               Exit For
                            End If
                        Next
                        'SE NÃO ENCONTRAR O LANCAMENTO NA MATRIZ, ADICIONAR ELE
                        If Not Achou Then
                           ReDim Preserve aDebito(UBound(aDebito) + 1)
                           nEval = UBound(aDebito)
                           aDebito(nEval).nCodReduzido = !CODREDUZIDO
                           aDebito(nEval).nAno = !AnoExercicio
                           aDebito(nEval).nLanc = !CodLancamento
                           aDebito(nEval).nSeq = !SeqLancamento
                           aDebito(nEval).nParc = !NumParcela
                           aDebito(nEval).nCompl = !CODCOMPLEMENTO
                           aDebito(nEval).nSituacao = !statuslanc
                           aDebito(nEval).sSituacao = !Situacao
                           aDebito(nEval).sVencto = Format(!DataVencimento, "dd/mm/yyyy")
                           aDebito(nEval).nValorTributo = !ValorTributo
                           If bJurosMulta Then
                               aDebito(nEval).nValorJuros = !ValorJuros
                               aDebito(nEval).nValorMulta = !ValorMulta
                           End If
                           aDebito(nEval).nValorCorrecao = !ValorCorrecao
                        Else
                            'SE ENCONTRAR ADICIONAR O VALOR AO JA EXISTENTE
                            aDebito(x).nValorTributo = aDebito(x).nValorTributo + !ValorTributo
                            If bJurosMulta Then
                                aDebito(x).nValorJuros = aDebito(x).nValorJuros + !ValorJuros
                                aDebito(x).nValorMulta = aDebito(x).nValorMulta + !ValorMulta
                            End If
                            aDebito(x).nValorCorrecao = aDebito(x).nValorCorrecao + !ValorCorrecao
                        End If
                       .MoveNext
                    Loop
                   .Close
                End With
               .MoveNext
            Loop
           .Close
        End With
       'ATUALIZA TABELA ORIGEM REPARC
        nSomaPrincipal = 0: nSomaMulta = 0: nSomaJuros = 0: nSomaCorrecao = 0
        For x = 1 To UBound(aDebito)
            With aDebito(x)
                Sql = "UPDATE ORIGEMREPARC SET PRINCIPAL=" & Virg2Ponto(CStr(Round(.nValorTributo, 2))) & ",JUROS=" & Virg2Ponto(CStr(Round(.nValorJuros, 2))) & ",MULTA=" & Virg2Ponto(CStr(Round(.nValorMulta, 2))) & ","
                Sql = Sql & "CORRECAO=" & Virg2Ponto(CStr(Round(.nValorCorrecao, 2))) & " WHERE CODREDUZIDO=" & .nCodReduzido & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO="
                Sql = Sql & .nLanc & " AND NUMSEQUENCIA=" & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl
                cn.Execute Sql, rdExecDirect
                nSomaPrincipal = nSomaPrincipal + .nValorTributo
                nSomaJuros = nSomaJuros + .nValorJuros
                nSomaMulta = nSomaMulta + .nValorMulta
                nSomaCorrecao = nSomaCorrecao + .nValorCorrecao
            End With
        Next
       
       'VALORES DAS PARCELAS DE DESTINO
        nValorLanc = nSomaPrincipal / nQtde
        If bJurosMulta Then
            nValorJuros = nSomaJuros / nQtde
            nValorMulta = nSomaMulta / nQtde
        End If
        nValorCorrecao = nSomaCorrecao / nQtde
        nValorTotal = nSomaPrincipal + nSomaJuros + nSomaMulta + nSomaCorrecao
        nValorParcela = nValorLanc + nValorJuros + nValorMulta + nValorCorrecao
        nSaldo = nValorTotal
               
               
               
DEST:
        ReDim aJuros(0)
       'CARREGA PARCELAS DE DESTINO E CALCULA O JUROS A SER APLICADO/QTDE DE PARCELAS > 2006
        Sql = "SELECT * FROM vwCNSREPARCELAMENTOD WHERE CODREDUZIDO=" & nCodReduz & " AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            'APROVEITAMOS O SELECT E VERIFICAMOS SE EXISTE HONORARIOS
            Sql = "SELECT VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=" & !CodLancamento & " AND "
            Sql = Sql & "SEQLANCAMENTO=" & !numsequencia & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO=90"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If RdoAux2.RowCount > 0 Then
                nValorHon = RdoAux2!ValorTributo
            Else
                nValorHon = 0
            End If
            
            nSomaJurosValor = 0: nQtdeJuros = 0: nJurosApl = 0
            Do Until .EOF
                ReDim Preserve aJuros(UBound(aJuros) + 1)
                If Year(!DataVencimento) > 2006 Then
                    nQtdeJuros = nQtdeJuros + 1
                    nJurosValor = nJurosPerc * nSaldo / 100
                    nSomaJurosValor = nSomaJurosValor + nJurosValor
                End If
                nSaldo = nSaldo - nValorParcela
                aJuros(UBound(aJuros)).nCodReduzido = !CODREDUZIDO
                aJuros(UBound(aJuros)).nAno = Year(!DataVencimento)
                aJuros(UBound(aJuros)).nLanc = !CodLancamento
                aJuros(UBound(aJuros)).nSeq = !numsequencia
                aJuros(UBound(aJuros)).nParc = !NumParcela
                aJuros(UBound(aJuros)).nCompl = !CODCOMPLEMENTO
                aJuros(UBound(aJuros)).nSaldo = nSaldo
                If Year(!DataVencimento) > 2006 Then
                    aJuros(UBound(aJuros)).nJurosPerc = nJurosPerc
                    aJuros(UBound(aJuros)).nJurosValor = nJurosValor
                End If
                
               .MoveNext
            Loop
            
           .Close
        End With
        nJurosApl = nSomaJurosValor / nQtdeJuros
        
        
        For x = 1 To UBound(aJuros)
            With aJuros(x)
                
               'ATUALIZA TABELA DESTINOREPARC
                Sql = "UPDATE DESTINOREPARC SET VALORLIQUIDO=" & Virg2Ponto(CStr(Round(nValorLanc, 2))) & ",JUROS=" & Virg2Ponto(CStr(Round(nValorJuros, 2))) & ","
                Sql = Sql & "MULTA=" & Virg2Ponto(CStr(Round(nValorMulta, 2))) & ",CORRECAO=" & Virg2Ponto(CStr(Round(nValorCorrecao, 2))) & ",VALORPRINCIPAL=" & Virg2Ponto(CStr(Round(nValorParcela, 2))) & ","
                Sql = Sql & "SALDO=" & Virg2Ponto(CStr(Round(.nSaldo, 2))) & ",JUROSPERC=" & Virg2Ponto(CStr(Round(.nJurosPerc, 2))) & ",JUROSVALOR=" & Virg2Ponto(CStr(Round(.nJurosValor, 2))) & ","
                Sql = Sql & "JUROSAPL=" & IIf(.nAno > 2006, Virg2Ponto(CStr(Round(nJurosApl, 2))), 0) & ",HONORARIO=" & Virg2Ponto(CStr(Round(nValorHon, 2))) & ",TOTAL=" & Virg2Ponto(CStr(Round(nValorLanc + nValorJuros + nValorMulta + nValorCorrecao + nValorHon + IIf(.nAno > 2006, nJurosApl, 0), 2))) & " WHERE "
                Sql = Sql & "CODREDUZIDO=" & .nCodReduzido & " AND ANOEXERCICIO=" & .nAno & " AND CODLANCAMENTO=" & .nLanc & " AND NUMSEQUENCIA="
                Sql = Sql & .nSeq & " AND NUMPARCELA=" & .nParc & " AND CODCOMPLEMENTO=" & .nCompl
                cn.Execute Sql, rdExecDirect
                
               'INSERE O TRIBUTO JUROS APLICADO EM CADA PARCELA (TRIBUTO 585)
                If .nAno > 2006 Then
                    Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                    Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                    Sql = Sql & .nCodReduzido & "," & .nAno & "," & .nLanc & "," & .nSeq & ","
                    Sql = Sql & .nParc & "," & .nCompl & "," & 585 & "," & Virg2Ponto(CStr(nJurosApl)) & ")"
                    cn.Execute Sql, rdExecDirect
                End If
            End With
        Next
proximo:
       .MoveNext
    Loop
   .Close
End With
MsgBox "FIM"

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim sNumProc As String, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

Sql = "SELECT ANO,CODREDUZIDO,NUMPROC,ANOPROC FROM PROCESSOBLOQUEIO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNumProc = CStr(!NumProc) & "/" & CStr(!AnoProc)
        Sql = "SELECT codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento From debitoparcela Where CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO>2006  AND CODLANCAMENTO=20 AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                nSeq = !SeqLancamento
                Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & !NumParcela & " AND CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    Do Until .EOF
                        '**atualiza o valor dos débitos
                        Sql = "UPDATE debitotributo Set valortributo = " & Virg2Ponto(Round(!ValorTributo * 100 / 103.7, 2)) & " WHERE CODREDUZIDO=" & RdoAux2!CODREDUZIDO & " AND ANOEXERCICIO=" & RdoAux2!AnoExercicio & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & RdoAux2!NumParcela & " AND CODCOMPLEMENTO=" & RdoAux2!CODCOMPLEMENTO & " AND CodTributo = " & !CodTributo
                        cn.Execute Sql, rdExecDirect
                       .MoveNext
                    Loop
                   .Close
                End With
                '***atualiza o status***
    '            Sql = "UPDATE debitoparcela Set statuslanc=18 Where (CODREDUZIDO = " & Rdoaux!CODREDUZIDO & ") And (AnoExercicio = " & Val(txtAno.text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") "
    '            cn.Execute Sql, rdExecDirect
               .MoveNext
           Loop
          .Close
        End With
        
       .MoveNext
    Loop
   .Close
End With
MsgBox "FIM"
End Sub



Private Sub Command1old_Click()
Dim sNumProc As String, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

Sql = "SELECT ANO,CODREDUZIDO,NUMPROC,ANOPROC FROM PROCESSOBLOQUEIO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNumProc = CStr(!NumProc) & "/" & CStr(!AnoProc)
        Sql = "SELECT seqlancamento From debitoparcela Where CODREDUZIDO=" & !CODREDUZIDO & " AND CODLANCAMENTO=20 AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nSeq = !SeqLancamento
            Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & RdoAux!CODREDUZIDO & " AND ANOEXERCICIO=2006 AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=2 AND CODTRIBUTO<>3"
            Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux3
                Do Until .EOF
                    '**atualiza o valor dos débitos
                    Sql = "UPDATE debitotributo Set valortributo = " & Virg2Ponto(!ValorTributo) & "  Where (CODREDUZIDO = " & RdoAux!CODREDUZIDO & ") And (AnoExercicio = " & Val(txtAno.Text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") AND (CodTributo = " & !CodTributo & " )"
                    cn.Execute Sql, rdExecDirect
                   .MoveNext
                Loop
               .Close
            End With
            '***atualiza o status***
            Sql = "UPDATE debitoparcela Set statuslanc=18 Where (CODREDUZIDO = " & RdoAux!CODREDUZIDO & ") And (AnoExercicio = " & Val(txtAno.Text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") "
            cn.Execute Sql, rdExecDirect
           .Close
        End With
        
       .MoveNext
    Loop
   .Close
End With
MsgBox "FIM"
End Sub

Private Sub Command2_Click()
Dim sNumProc As String, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

Sql = "SELECT ANO,CODREDUZIDO,NUMPROC,ANOPROC FROM PROCESSOBLOQUEIO WHERE CODREDUZIDO=110432"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNumProc = CStr(!NumProc) & "/" & CStr(!AnoProc)
        Sql = "SELECT codreduzido,anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento From debitoparcela Where CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO>2006  AND CODLANCAMENTO=20 AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND CODTRIBUTO=587"
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux3
                    If .RowCount > 0 Then
'                        .Close
'                        GoTo proximo
                    End If
                   .Close
                End With
            
                If .AbsolutePosition = 1 Then
                    Sql = "SELECT SUM(VALORTRIBUTO) AS SOMA FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & !CODREDUZIDO & " AND ANOEXERCICIO=" & !AnoExercicio & " AND "
                    Sql = Sql & "CODLANCAMENTO=" & !CodLancamento & " AND SEQLANCAMENTO=" & !SeqLancamento & " AND NUMPARCELA=" & !NumParcela & " AND "
                    Sql = Sql & "CODCOMPLEMENTO=" & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                        nValorTributo = (!soma * 0.037)
                        RdoAux2.Close
                    End With

                    '***atualiza o status***
 '                   Sql = "UPDATE debitoparcela Set statuslanc=3 Where (CODREDUZIDO = " & Val(txtCod.text) & ") And (AnoExercicio = " & Val(txtAno.text) & ") And (CodLancamento = 20) And (SeqLancamento = " & nSeq & ") "
'                    cn.Execute Sql, rdExecDirect
                '   .Close
                End If
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
                Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,VALORTRIBUTO) VALUES("
                Sql = Sql & !CODREDUZIDO & "," & !AnoExercicio & "," & !CodLancamento & "," & !SeqLancamento & ","
                Sql = Sql & !NumParcela & "," & !CODCOMPLEMENTO & "," & 587 & "," & Virg2Ponto(CStr(nValorTributo)) & ")"
                cn.Execute Sql, rdExecDirect
               .MoveNext
           Loop
          .Close
        End With
proximo:
       .MoveNext
    Loop
   .Close
End With
MsgBox "FIM"

End Sub

Private Sub Form_Load()
Dim x As Integer

For x = 2008 To Year(Now) + 1
    txtAno.AddItem CStr(x)
Next

Centraliza Me
Me.Top = Me.Top + 500
If NomeDeLogin <> "SCHWARTZ" Then
    Opt(1).Enabled = False
End If

If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "ISRAEL" Or NomeDeLogin = "IORIO" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "GLEISE" Or _
NomeDeLogin = "SOLANGE" Or IsAtendente Then
    cmdCalculo.Enabled = True
End If
txtAno.ListIndex = 0
End Sub

Private Sub Opt_Click(Index As Integer)
If Index = 0 Then
    txtCod.Locked = False
    txtCod.BackColor = Branco
    txtNumProc.Locked = False
    txtNumProc.BackColor = Branco
Else
    txtCod.Locked = True
    txtCod.BackColor = Kde
    txtNumProc.Locked = True
    txtNumProc.BackColor = Kde
End If
End Sub

Private Sub txtAno_Click()
If txtAno.ListIndex = 0 Then
    txtPerc.Text = "4,15"
ElseIf txtAno.ListIndex = 1 Then
    txtPerc.Text = "6,25"
ElseIf txtAno.ListIndex = 2 Then
    txtPerc.Text = "4,34"
ElseIf txtAno.ListIndex = 3 Then
    txtPerc.Text = "4,70"
ElseIf txtAno.ListIndex = 4 Then
    txtPerc.Text = "7,31"
ElseIf txtAno.ListIndex = 5 Then
    txtPerc.Text = "5,2824"
ElseIf txtAno.ListIndex = 6 Then
    txtPerc.Text = "5,86"
ElseIf txtAno.ListIndex = 7 Then
    txtPerc.Text = "6,7465"
ElseIf txtAno.ListIndex = 8 Then
    txtPerc.Text = "9,4932"
ElseIf txtAno.ListIndex = 9 Then
    txtPerc.Text = "8,4764"
ElseIf txtAno.ListIndex = 10 Then
    txtPerc.Text = "2,5377"
ElseIf txtAno.ListIndex = 11 Then
    txtPerc.Text = "4,53"
End If

End Sub

Private Sub txtAno_KeyPress(KeyAscii As Integer)
Tweak txtAno, KeyAscii, IntegerPositive
End Sub

Private Sub txtPerc_KeyPress(KeyAscii As Integer)
Tweak txtPerc, KeyAscii, DecimalPositive
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

