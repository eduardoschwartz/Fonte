VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmISSMensal 
   Appearance      =   0  'Flat
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arrecadação Mensal de ISS"
   ClientHeight    =   4260
   ClientLeft      =   4545
   ClientTop       =   3390
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   5535
   Begin VB.CheckBox chkSimples 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Incluir Simples Nacional"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1860
      TabIndex        =   27
      Top             =   1320
      Width           =   3405
   End
   Begin VB.CheckBox chkMaior 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Caption         =   "Maiores Contribuintes de ISS Variável"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1860
      TabIndex        =   26
      Top             =   1020
      Width           =   3405
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00EEEEEE&
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   30
      TabIndex        =   22
      Top             =   1650
      Width           =   5445
      Begin VB.TextBox txtAtiv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   90
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Atividades:"
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   885
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Tipo de Relatório"
      ForeColor       =   &H00000080&
      Height          =   1260
      Left            =   30
      TabIndex        =   19
      Top             =   2970
      Width           =   1755
      Begin VB.OptionButton optPag 
         BackColor       =   &H00EEEEEE&
         Caption         =   "ISS Eletrônico"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   28
         Top             =   855
         Width           =   1440
      End
      Begin VB.OptionButton optPag 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Não Pagos"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   1395
      End
      Begin VB.OptionButton optPag 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Pagos"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Tipo de Empresa"
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   30
      TabIndex        =   16
      Top             =   2100
      Width           =   1755
      Begin VB.OptionButton OptEmp 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cadastrada"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.OptionButton OptEmp 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cidadão"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   570
         Width           =   1395
      End
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Sair da Tela"
      Top             =   3330
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frmISSMensal.frx":0000
      PICN            =   "frmISSMensal.frx":001C
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
      Left            =   2880
      TabIndex        =   7
      ToolTipText     =   "Emitir Relatório"
      Top             =   3330
      Width           =   1260
      _ExtentX        =   2223
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
      MICON           =   "frmISSMensal.frx":008A
      PICN            =   "frmISSMensal.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Faixa de Códigos"
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   1800
      TabIndex        =   11
      Top             =   2100
      Width           =   3675
      Begin VB.TextBox txtCod2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2580
         MaxLength       =   6
         TabIndex        =   6
         Top             =   510
         Width           =   885
      End
      Begin VB.TextBox txtCod1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   810
         MaxLength       =   6
         TabIndex        =   5
         Top             =   510
         Width           =   885
      End
      Begin VB.CheckBox chkCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Todos os Códigos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Final...:"
         Height          =   255
         Index           =   1
         Left            =   1950
         TabIndex        =   14
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicial...:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Periodo"
      ForeColor       =   &H00000080&
      Height          =   885
      Left            =   1800
      TabIndex        =   10
      Top             =   30
      Width           =   3675
      Begin VB.TextBox txtAno1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano...:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Tipo de ISS"
      ForeColor       =   &H00000080&
      Height          =   1605
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   1755
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00EEEEEE&
         Caption         =   "ISSQN"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1230
         Width           =   1395
      End
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00EEEEEE&
         Caption         =   "ISS Fixo"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   900
         Width           =   1395
      End
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00EEEEEE&
         Caption         =   "ISS Variável"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   1395
      End
      Begin VB.OptionButton OptTipo 
         BackColor       =   &H00EEEEEE&
         Caption         =   "ISS Estimado"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   240
      Left            =   2910
      TabIndex        =   15
      Top             =   3780
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmISSMensal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type REFS
    nAnoVT As Integer
    nMesVT As Integer
    nAnoRL As Integer
    nMesRL As Integer
End Type

Private Type tpMatriz
    nCodReduz As Long
    nAno As Integer
    nMes1 As Double
    nMes2 As Double
    nMes3 As Double
    nMes4 As Double
    nMes5 As Double
    nMes6 As Double
    nMes7 As Double
    nMes8 As Double
    nMes9 As Double
    nMes10 As Double
    nMes11 As Double
    nMes12 As Double
End Type

Private Type RELMOB2
    nCodReduz As Long
    sNome As String
    bSimples As Boolean
    nValor(12) As Double
End Type

Dim RdoAux As rdoResultset, Sql As String, aMatriz() As tpMatriz, bAchou As Boolean


Private Sub chkCodigo_Click()
If chkCodigo.Value = 1 Then
    txtCod1.Locked = True: txtCod2.Locked = True
    txtCod1.BackColor = Kde: txtCod2.BackColor = Kde
Else
    txtCod1.Locked = False: txtCod2.Locked = False
    txtCod1.BackColor = Branco: txtCod2.BackColor = Branco
End If
End Sub

Private Sub cmdPrint_Click()

If optPag(2).Value = True Then
    BuildReportMob2
    Exit Sub
End If

If optPag(1).Value = True And OptEmp(1).Value = True Then
    MsgBox "Relatório de débitos não pagos apenas para empresas cadastradas.", vbExclamation, "Atenção"
    Exit Sub
End If

If Val(txtAno1.Text) = 0 Then
    MsgBox "Digite ano de impressão.", vbExclamation, "Atenção"
    Exit Sub
End If

If chkCodigo.Value = 0 Then
    If Val(txtCod1.Text) = 0 Or Val(txtCod2.Text) = 0 Then
        MsgBox "Digite codigo inicial e final.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If Val(txtCod1.Text) > Val(txtCod2.Text) Then
        MsgBox "Código inicial maior que código final.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If OptEmp(0).Value = True Then
        If Val(txtCod1.Text) < 100000 Or Val(txtCod1.Text) > 300000 Then
            MsgBox "Código inicial fora do limite.", vbExclamation, "Atenção"
            Exit Sub
        End If
        If Val(txtCod2.Text) < 100000 Or Val(txtCod2.Text) > 300000 Then
            MsgBox "Código Final fora do limite.", vbExclamation, "Atenção"
            Exit Sub
        End If
    Else
        If Val(txtCod1.Text) < 500000 Or Val(txtCod1.Text) > 800000 Then
            MsgBox "Código inicial fora do limite.", vbExclamation, "Atenção"
            Exit Sub
        End If
        If Val(txtCod2.Text) < 500000 Or Val(txtCod2.Text) > 800000 Then
            MsgBox "Código Final fora do limite.", vbExclamation, "Atenção"
            Exit Sub
        End If
    End If
Else
    If OptEmp(0).Value = True Then
        txtCod1.Text = 100000
        txtCod2.Text = 300000
    Else
        txtCod1.Text = 500000
        txtCod2.Text = 800000
    End If
End If

'If optPag(0).Value = True Then
If MsgBox("Deseja gerar o relatório ?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") = vbYes Then
    GeraMatriz
Else
    Exit Sub
End If

If OptEmp(0).Value = True Then
    If chkMaior.Value = False Then
        frmReport.ShowReport "ISSMENSAL", frmMdi.hwnd, Me.hwnd
    Else
        frmReport.ShowReport "ISSMENSALVERTICAL", frmMdi.hwnd, Me.hwnd
    End If
Else
    frmReport.ShowReport "ISSMENSALFORA", frmMdi.hwnd, Me.hwnd
End If

Sql = "DELETE FROM ISSMENSAL WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
chkCodigo.Value = 1: Pb.Value = 0
txtAno1.Text = Year(Now)
Centraliza Me
End Sub

Private Sub GeraMatriz()
Dim xId As Long, nNumRec As Long, x As Integer, bAchou As Boolean, aRef() As REFS, nMesRef As Integer, nAnoRef As Integer, bSimples As Boolean
Ocupado
ReDim aRef(0)

Sql = "SELECT * FROM SIMPLESREF ORDER BY ANOVT,MESVT,ANORL,MESRL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aRef(UBound(aRef) + 1)
        aRef(UBound(aRef)).nAnoVT = !ANOVT
        aRef(UBound(aRef)).nMesVT = !MESVT
        aRef(UBound(aRef)).nAnoRL = !ANORL
        aRef(UBound(aRef)).nMesRL = !MESRL
       .MoveNext
    Loop
   .Close
End With

If OptEmp(0).Value = True Then
    Sql = "SELECT distinct codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, codtributo, valortributo, datarecebimento, restituido, SIMPLES , DataVencimento, dataencerramento FROM vwISSMENSAL2 "
    'Sql = Sql & "WHERE   YEAR(DATARECEBIMENTO) = " & Val(txtAno1.text) & " "
    'Sql = Sql & "WHERE   (CODREDUZIDO BETWEEN " & Val(txtCod1.text) & " AND " & Val(txtCod2.text) & ") and YEAR(DATARECEBIMENTO) = " & Val(txtAno1.text) & " AND YEAR(DATAVENCIMENTO) = " & Val(txtAno1.text) & " "
    Sql = Sql & "WHERE   (CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & ") and YEAR(DATARECEBIMENTO) = " & Val(txtAno1.Text) & " "
'    sql=SQL & "AND CODLANCAMENTO<>20"
    If optTipo(0).Value = True Then
        Sql = Sql & "AND CODTRIBUTO in (12,569) " 'ESTIMADO
    ElseIf optTipo(1).Value = True Then
        Sql = Sql & "AND CODTRIBUTO in (13,502) " 'VARIAVEL
    ElseIf optTipo(2).Value = True Then
        Sql = Sql & "AND CODTRIBUTO in (11,555) " 'FIXO
    ElseIf optTipo(3).Value = True Then
        Sql = Sql & "AND CODTRIBUTO IN (179,180,181,182,183,184,185,19) " 'ISSQN
    End If
    If txtAtiv.Text <> "" Then
        Sql = Sql & "AND CODATIVIDADE in (" & txtAtiv.Text & ") "
    End If
    Sql = Sql & "ORDER BY CODREDUZIDO"
Else
    Sql = "SELECT CODREDUZIDO,ANO,MES,VALORTOTAL FROM VWISSMENSALFORA "
    Sql = Sql & "WHERE (CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text) & ") "
    Sql = Sql & "AND (ANO BETWEEN " & Val(txtAno1.Text) & " AND " & Val(txtAno2.Text) & ") "
    If optTipo(0).Value = True Then
        Sql = Sql & "AND CODTRIBUTO = 569 " 'ESTIMADO
    Else
        Sql = Sql & "AND CODTRIBUTO = 502 " 'VARIAVEL
    End If
    Sql = Sql & "ORDER BY CODREDUZIDO,ANO,MES"
End If

If chkMaior.Value = vbChecked Then
    Sql = "SELECT * FROM VWISSMENSAL3 "
    Sql = Sql & "WHERE  CODREDUZIDO NOT in (SELECT CODMOBILIARIO FROM vwMOBILIARIOSUSPENSO WHERE CODTIPOEVENTO=2) AND  SIMPLES=0 AND DATAENCERRAMENTO IS NULL AND YEAR(DATARECEBIMENTO) = " & Val(txtAno1.Text) & " AND YEAR(DATAVENCIMENTO) = " & Val(txtAno1.Text) & " "
    'Sql = Sql & "WHERE SIMPLES=0 AND YEAR(DATAVENCIMENTO) = " & Val(txtAno1.text) & " "
    If optTipo(0).Value = True Then
        Sql = Sql & "AND CODTRIBUTO in (12,569) " 'ESTIMADO
    ElseIf optTipo(1).Value = True Then
        Sql = Sql & "AND CODTRIBUTO in (13,502) " 'VARIAVEL
    ElseIf optTipo(2).Value = True Then
        Sql = Sql & "AND CODTRIBUTO in (11,555) " 'FIXO
    ElseIf optTipo(3).Value = True Then
        Sql = Sql & "AND CODTRIBUTO IN (179,180,181,182,183,184,185,19) " 'ISSQN
    End If
    If chkCodigo.Value = vbUnchecked Then
        Sql = Sql & "AND CODREDUZIDO BETWEEN " & Val(txtCod1.Text) & " AND " & Val(txtCod2.Text)
    End If
    Sql = Sql & " ORDER BY CODREDUZIDO"
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    xId = 1: nNumRec = .RowCount: ReDim aMatriz(0)
    Do Until .EOF
        'nAnoRef = Year(!DataVencimento)
        nAnoRef = Year(!datarecebimento)
        If Val(SubNull(!SIMPLES)) = 1 Then
            bSimples = True
        Else
            bSimples = False
        End If
       ' If !CODREDUZIDO = 100070 Then MsgBox "teste"
        If xId Mod 10 = 0 Then
           CallPb xId, nNumRec
        End If
        bAchou = False
        For h = 1 To UBound(aMatriz)
            If aMatriz(h).nCodReduz = !CODREDUZIDO Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            x = UBound(aMatriz) + 1
            ReDim Preserve aMatriz(x)
            aMatriz(x).nCodReduz = !CODREDUZIDO
            'aMatriz(x).nAno = Year(!DataVencimento)
            aMatriz(x).nAno = Year(!datarecebimento)
        Else
            x = h
        End If
        
        If bSimples Then
            nMesRef = 0: nAnoRef = 0
            For t = 1 To UBound(aRef)
                If Year(!DataVencimento) = aRef(t).nAnoVT And Month(!DataVencimento) = aRef(t).nMesVT Then
                    nAnoRef = aRef(t).nAnoRL: nMesRef = aRef(t).nMesRL
'                    If nAnoRef = 2008 Then MsgBox "teste"
                    Exit For
                End If
            Next
            If nAnoRef <> Val(txtAno1.Text) Then GoTo PROXIMO2
            aMatriz(x).nAno = nAnoRef
            
        Else
            'nMesRef = Month(!DataVencimento)
            nMesRef = Month(!datarecebimento)
'            If nAnoRef = Val(txtAno1.text) Then
                
'                If nMesRef > 1 Then
'                    nMesRef = nMesRef - 1
'                    aMatriz(x).nAno = nAnoRef
'                ElseIf nMesRef = 1 Then
'                    GoTo PROXIMO2
''                    nMesRef = 12
''                    aMatriz(x).nAno = aMatriz(x).nAno - 1
'                End If
'            Else
'                MsgBox "teste"
'            End If
        End If
        
        Select Case nMesRef
            Case 1
                aMatriz(x).nMes1 = aMatriz(x).nMes1 + !ValorTributo
            Case 2
                aMatriz(x).nMes2 = aMatriz(x).nMes2 + !ValorTributo
            Case 3
                aMatriz(x).nMes3 = aMatriz(x).nMes3 + !ValorTributo
            Case 4
                aMatriz(x).nMes4 = aMatriz(x).nMes4 + !ValorTributo
            Case 5
                aMatriz(x).nMes5 = aMatriz(x).nMes5 + !ValorTributo
            Case 6
                aMatriz(x).nMes6 = aMatriz(x).nMes6 + !ValorTributo
            Case 7
                aMatriz(x).nMes7 = aMatriz(x).nMes7 + !ValorTributo
            Case 8
                aMatriz(x).nMes8 = aMatriz(x).nMes8 + !ValorTributo
            Case 9
                aMatriz(x).nMes9 = aMatriz(x).nMes9 + !ValorTributo
            Case 10
                aMatriz(x).nMes10 = aMatriz(x).nMes10 + !ValorTributo
            Case 11
                aMatriz(x).nMes11 = aMatriz(x).nMes11 + !ValorTributo
            Case 12
                aMatriz(x).nMes12 = aMatriz(x).nMes12 + !ValorTributo
        End Select
PROXIMO2:
        xId = xId + 1
       .MoveNext
    Loop
    Pb.Value = 100
   .Close
End With


''***** CRIA MES DE DEZEMBRO ****
'
'If OptEmp(0).Value = True Then
'    Sql = "SELECT distinct codreduzido, anoexercicio, codlancamento, seqlancamento, numparcela, codcomplemento, codtributo, valortributo, datarecebimento, restituido, SIMPLES , DataVencimento, dataencerramento FROM vwISSMENSAL2 "
'    Sql = Sql & "WHERE (CODREDUZIDO BETWEEN " & Val(txtCod1.text) & " AND " & Val(txtCod2.text) & ") and YEAR(DATARECEBIMENTO) = " & Val(txtAno1.text) + 1 & " AND MONTH(DATARECEBIMENTO)=1" & " "
'    If OptTipo(0).Value = True Then
'        Sql = Sql & "AND CODTRIBUTO in (12,569) " 'ESTIMADO
'    ElseIf OptTipo(1).Value = True Then
'        Sql = Sql & "AND CODTRIBUTO in (13,502) " 'VARIAVEL
'    ElseIf OptTipo(2).Value = True Then
'        Sql = Sql & "AND CODTRIBUTO in (11,555) " 'FIXO
'    ElseIf OptTipo(3).Value = True Then
'        Sql = Sql & "AND CODTRIBUTO IN (179,180,181,182,183,184,185,19) " 'ISSQN
'    End If
'    If txtAtiv.text <> "" Then
'        Sql = Sql & "AND CODATIVIDADE in (" & txtAtiv.text & ") "
'    End If
'    Sql = Sql & "ORDER BY CODREDUZIDO"
'End If
'
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
''    xId = 1: nNumRec = .RowCount: ReDim aMatriz(0)
'    Do Until .EOF
'        nAnoRef = Year(!DataRecebimento)
'        If Val(SubNull(!SIMPLES)) = 1 Then
'            bSimples = True
'        Else
'            bSimples = False
'        End If
'        bAchou = False
'        For h = 1 To UBound(aMatriz)
'            If aMatriz(h).nCodReduz = !CODREDUZIDO Then
'                bAchou = True
'                Exit For
'            End If
'        Next
'        If Not bAchou Then
'            x = UBound(aMatriz) + 1
'            ReDim Preserve aMatriz(x)
'            aMatriz(x).nCodReduz = !CODREDUZIDO
'            aMatriz(x).nAno = Year(!DataRecebimento)
'        Else
'            x = h
'        End If
'
'        If bSimples Then
'            nMesRef = 0: nAnoRef = 0
'            For t = 1 To UBound(aRef)
'                If Year(!DataVencimento) = aRef(t).nAnoVT And Month(!DataVencimento) = aRef(t).nMesVT Then
'                    nAnoRef = aRef(t).nAnoRL: nMesRef = aRef(t).nMesRL
''                    If nAnoRef = 2008 Then MsgBox "teste"
'                    Exit For
'                End If
'            Next
'            If nAnoRef <> Val(txtAno1.text) Then GoTo PROXIMO5
'            aMatriz(x).nAno = nAnoRef
'
'        Else
'            'nMesRef = Month(!DataVencimento)
'            nMesRef = 12
'        End If
'
'        Select Case nMesRef
'            Case 1
'                aMatriz(x).nMes1 = aMatriz(x).nMes1 + !ValorTributo
'            Case 2
'                aMatriz(x).nMes2 = aMatriz(x).nMes2 + !ValorTributo
'            Case 3
'                aMatriz(x).nMes3 = aMatriz(x).nMes3 + !ValorTributo
'            Case 4
'                aMatriz(x).nMes4 = aMatriz(x).nMes4 + !ValorTributo
'            Case 5
'                aMatriz(x).nMes5 = aMatriz(x).nMes5 + !ValorTributo
'            Case 6
'                aMatriz(x).nMes6 = aMatriz(x).nMes6 + !ValorTributo
'            Case 7
'                aMatriz(x).nMes7 = aMatriz(x).nMes7 + !ValorTributo
'            Case 8
'                aMatriz(x).nMes8 = aMatriz(x).nMes8 + !ValorTributo
'            Case 9
'                aMatriz(x).nMes9 = aMatriz(x).nMes9 + !ValorTributo
'            Case 10
'                aMatriz(x).nMes10 = aMatriz(x).nMes10 + !ValorTributo
'            Case 11
'                aMatriz(x).nMes11 = aMatriz(x).nMes11 + !ValorTributo
'            Case 12
'                aMatriz(x).nMes12 = aMatriz(x).nMes12 + !ValorTributo
'        End Select
'PROXIMO5:
'        xId = xId + 1
'       .MoveNext
'    Loop
'   .Close
'End With
'
'*******************************


Liberado

If chkMaior.Value = vbUnchecked Then GoTo fim

Sql = "SELECT DISTINCT mobiliario.codigomob FROM mobiliario INNER JOIN mobiliarioatividadeiss ON "
Sql = Sql & "mobiliario.codigomob = mobiliarioatividadeiss.codmobiliario Where  CODIGOMOB NOT in (SELECT CODMOBILIARIO FROM vwMOBILIARIOSUSPENSO WHERE CODTIPOEVENTO=2) and (mobiliario.dataencerramento Is Null) And "
Sql = Sql & "(mobiliarioatividadeiss.CodTributo = 13)  "
If chkSimples.Value = vbUnchecked Then
    Sql = Sql & " AND MOBILIARIO.SIMPLES = 0 "
End If
Sql = Sql & " ORDER BY mobiliario.codigomob"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        bAchou = False
        For x = 1 To UBound(aMatriz)
            If aMatriz(x).nCodReduz = !codigomob Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aMatriz(UBound(aMatriz) + 1)
            aMatriz(UBound(aMatriz)).nCodReduz = !codigomob
            aMatriz(UBound(aMatriz)).nMes1 = 0
            aMatriz(UBound(aMatriz)).nMes2 = 0
            aMatriz(UBound(aMatriz)).nMes3 = 0
            aMatriz(UBound(aMatriz)).nMes4 = 0
            aMatriz(UBound(aMatriz)).nMes5 = 0
            aMatriz(UBound(aMatriz)).nMes6 = 0
            aMatriz(UBound(aMatriz)).nMes7 = 0
            aMatriz(UBound(aMatriz)).nMes8 = 0
            aMatriz(UBound(aMatriz)).nMes9 = 0
            aMatriz(UBound(aMatriz)).nMes10 = 0
            aMatriz(UBound(aMatriz)).nMes11 = 0
            aMatriz(UBound(aMatriz)).nMes12 = 0
        End If
       .MoveNext
    Loop
   .Close
End With





fim:
Sql = "DELETE FROM ISSMENSAL WHERE COMPUTER='" & NomeDoUsuario & "'"
cn.Execute Sql, rdExecDirect

Pb.Value = 0
If chkMaior.Value = vbChecked Then
    For x = 1 To UBound(aMatriz)
        If x Mod 10 = 0 Then
           CallPb CLng(x), CLng(UBound(aMatriz))
        End If
        
        With aMatriz(x)
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 1 & "," & Virg2Ponto(CStr(.nMes1)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 2 & "," & Virg2Ponto(CStr(.nMes2)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 3 & "," & Virg2Ponto(CStr(.nMes3)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 4 & "," & Virg2Ponto(CStr(.nMes4)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 5 & "," & Virg2Ponto(CStr(.nMes5)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 6 & "," & Virg2Ponto(CStr(.nMes6)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 7 & "," & Virg2Ponto(CStr(.nMes7)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 8 & "," & Virg2Ponto(CStr(.nMes8)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 9 & "," & Virg2Ponto(CStr(.nMes9)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 10 & "," & Virg2Ponto(CStr(.nMes10)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 11 & "," & Virg2Ponto(CStr(.nMes11)) & ")"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 12 & "," & Virg2Ponto(CStr(.nMes12)) & ")"
            cn.Execute Sql, rdExecDirect
        End With
    Next
Else
    For x = 1 To UBound(aMatriz)
        With aMatriz(x)
            If .nAno <> Val(txtAno1.Text) Then
                GoTo proximo3
            End If
            Sql = "INSERT ISSMENSAL(COMPUTER,CODREDUZIDO,ANO,MES,MES1,MES2,MES3,MES4,MES5,MES6,MES7,MES8,MES9,MES10,MES11,MES12) VALUES('"
            Sql = Sql & NomeDoUsuario & "'," & .nCodReduz & "," & .nAno & "," & 0 & "," & Virg2Ponto(CStr(.nMes1)) & "," & Virg2Ponto(CStr(.nMes2)) & "," & Virg2Ponto(CStr(.nMes3)) & "," & Virg2Ponto(CStr(.nMes4))
            Sql = Sql & "," & Virg2Ponto(CStr(.nMes5)) & "," & Virg2Ponto(CStr(.nMes6)) & "," & Virg2Ponto(CStr(.nMes7)) & "," & Virg2Ponto(CStr(.nMes8)) & "," & Virg2Ponto(CStr(.nMes9)) & "," & Virg2Ponto(CStr(.nMes10))
            Sql = Sql & "," & Virg2Ponto(CStr(.nMes11)) & "," & Virg2Ponto(CStr(.nMes12)) & ")"
            cn.Execute Sql, rdExecDirect
        End With
proximo3:
    Next
End If

End Sub

Private Sub CallPb(nPos As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPos * 100) / nTotal) <= 100 Then
   Pb.Value = (nPos * 100) / nTotal
Else
   Pb.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String

If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
FillSpace = sTmp

End Function

Private Sub BuildReportMob2()
Dim RdoAux As rdoResultset, Sql As String, aCodigos() As Long, x As Long, bFind As Boolean
Dim nMesRef As Integer, nAnoRef As Integer, aReg() As RELMOB2, nPos As Long, y As Long, nExercicio As Integer
Dim sNomeArq As String, FF1 As Integer, ax As String, ret As Long

nExercicio = Val(txtAno1.Text)
If nExercicio < 2004 Or nExercicio > Year(Now) Then
    MsgBox "Ano inválido!", vbCritical, "Atenção"
    Exit Sub
End If

Pb.Value = 0
ReDim aCodigos(0): ReDim aReg(0)
Sql = "SELECT DISTINCT codreduzido From debitoparcela WHERE (debitoparcela.codreduzido >= 100000 AND debitoparcela.codreduzido < 300000) and (codlancamento = 5) AND (statuslanc = 2 or statuslanc=3 or statuslanc = 7) AND "
Sql = Sql & "(datavencimento BETWEEN '02/01/" & CStr(nExercicio) & "' AND '01/31/" & CStr(nExercicio + 1) & "')"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCodigos(UBound(aCodigos) + 1)
        aCodigos(UBound(aCodigos)) = !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With

nPos = 1
For x = 1 To UBound(aCodigos)
'    If X = 30 Then Exit For
    CallPb x, UBound(aCodigos)
'    If aCodigos(x) = 101073 Then MsgBox "teste"
    Sql = "SELECT DISTINCT debitoparcela.codreduzido, debitoparcela.anoexercicio, debitoparcela.codlancamento, debitoparcela.seqlancamento, debitoparcela.numparcela, "
    Sql = Sql & "debitoparcela.codcomplemento, debitoparcela.datavencimento, debitopago.restituido, debitopago.valorpagoreal, vwFULLEMPRESA.razaosocial,vwFULLEMPRESA.SIMPLES "
    Sql = Sql & "FROM debitoparcela LEFT OUTER JOIN debitopago ON debitoparcela.codreduzido = debitopago.codreduzido AND debitoparcela.anoexercicio = debitopago.anoexercicio AND "
    Sql = Sql & "debitoparcela.codlancamento = debitopago.codlancamento AND debitoparcela.seqlancamento = debitopago.seqlancamento AND "
    Sql = Sql & "debitoparcela.numparcela = debitopago.numparcela AND debitoparcela.codcomplemento = debitopago.codcomplemento LEFT OUTER JOIN "
    Sql = Sql & "vwFULLEMPRESA ON debitoparcela.codreduzido = vwFULLEMPRESA.codigomob WHERE (debitoparcela.codreduzido =" & aCodigos(x) & ") and (debitoparcela.codlancamento = 5) AND (statuslanc = 2 or statuslanc=3 or statuslanc = 7) AND (datavencimento BETWEEN '02/01/" & CStr(nExercicio) & "' AND '01/31/" & CStr(nExercicio + 1) & "')" & " AND "
    Sql = Sql & "(debitopago.restituido IS NULL) "
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            DoEvents
            If Month(!DataVencimento) = 1 Then
                nMesRef = 12
            Else
                nMesRef = Month(!DataVencimento) - 1
            End If
            bFind = False
            For y = 1 To UBound(aReg)
                If aReg(y).nCodReduz = aCodigos(x) Then
                    bFind = True
                    Exit For
                End If
            Next
            
            If Not bFind Then
                ReDim Preserve aReg(UBound(aReg) + 1)
                nPos = UBound(aReg)
                aReg(nPos).nCodReduz = aCodigos(x)
                aReg(nPos).sNome = !razaosocial
                aReg(nPos).bSimples = IIf(!SIMPLES = "S", True, False)
                If Not IsNull(!valorpagoreal) Then
                    aReg(nPos).nValor(nMesRef) = !valorpagoreal
                End If
                nPos = nPos + 1
            Else
                If Not IsNull(!valorpagoreal) Then
                    aReg(y).nValor(nMesRef) = aReg(y).nValor(nMesRef) + !valorpagoreal
                End If
            End If
           .MoveNext
        Loop
       .Close
    End With
Next

sNomeArq = sPathBin & "\REPORTMOB2.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

Print #FF1, "***********************************************************"
Print #FF1, "ARRECADAÇÃO ANUAL DE ISS VARIÁVEL PARA " & CStr(nExercicio)
Print #FF1, "IMPRESSO EM " & Format(Now, "dd/mm/yyyy") & " - Fonte: GTI"
Print #FF1, "***********************************************************"
ax = ""
Print #FF1, ax
ax = FillSpace("CÓDIGO", 8) & FillSpace("RAZÃO SOCIAL", 42) & "SNA" & FillLeft("JANEIRO", 12) & FillLeft("FEVEREIRO", 12)
ax = ax & FillLeft("MARÇO", 12) & FillLeft("ABRIL", 12) & FillLeft("MAIO", 12) & FillLeft("JUNHO", 12) & FillLeft("JULHO", 12)
ax = ax & FillLeft("AGOSTO", 12) & FillLeft("SETEMBRO", 12) & FillLeft("OUTUBRO", 12) & FillLeft("NOVEMBRO", 12) & FillLeft("DEZEMBRO", 12)
Print #FF1, ax
Print #FF1, "******************************************************************************************************************************************************************************************************"

For x = 1 To UBound(aReg)
    CallPb x, UBound(aReg)
    With aReg(x)
        ax = FillSpace(CStr(.nCodReduz), 8) & FillSpace(.sNome, 42) & IIf(.bSimples, " S ", " N ")
        For y = 1 To 12
            If .nValor(y) = 0 Then
                Sql = "SELECT * FROM nfisseletrosmov WHERE CODIGO=" & .nCodReduz & " AND ANO=" & nExercicio & " AND MES=" & y
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    ax = ax & FillLeft("SEM MOV.", 12)
                Else
                    ax = ax & FillLeft(Format(.nValor(y), "#0.00"), 12)
                End If
                RdoAux.Close
            Else
                ax = ax & FillLeft(Format(.nValor(y), "#0.00"), 12)
            End If
        Next
        Print #FF1, ax
    End With
Next

Print #FF1, ""

Sql = "SELECT DISTINCT mobiliarioatividadeiss.codmobiliario, mobiliario.razaosocial, mobiliario.dataencerramento,mobiliario.simples "
Sql = Sql & "FROM mobiliarioatividadeiss INNER JOIN mobiliario ON mobiliarioatividadeiss.codmobiliario = mobiliario.codigomob "
Sql = Sql & "Where (mobiliarioatividadeiss.CodTributo = 13) And (mobiliario.dataencerramento Is Null) "
Sql = Sql & "ORDER BY mobiliarioatividadeiss.codmobiliario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        bFind = False
        For x = 1 To UBound(aReg)
            If aReg(x).nCodReduz = !CODMOBILIARIO Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            ax = FillSpace(!CODMOBILIARIO, 8) & FillSpace(!razaosocial, 42) & IIf(!SIMPLES = "S", " S ", " N ")
            For x = 1 To 12
                ax = ax & FillLeft("0,00", 12)
            Next
            Print #FF1, ax
        End If
       .MoveNext
    Loop
   .Close
End With


Close #FF1
Pb.Value = 0

ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

Liberado
MsgBox "Relatório disponível em " & sPathBin & "\REPORTMOB2.TXT"

End Sub

