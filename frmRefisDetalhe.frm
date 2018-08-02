VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRefisDetalhe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refis DAM por tributo"
   ClientHeight    =   1605
   ClientLeft      =   13260
   ClientTop       =   6600
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1605
   ScaleWidth      =   4890
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1185
      TabIndex        =   0
      Top             =   240
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRefisDetalhe.frx":0000
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
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   3585
      TabIndex        =   1
      Top             =   255
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRefisDetalhe.frx":001C
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
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   3570
      TabIndex        =   2
      ToolTipText     =   "Imprimir esta Tela"
      Top             =   1050
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmRefisDetalhe.frx":0038
      PICN            =   "frmRefisDetalhe.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   165
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblPB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100 %"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2565
      TabIndex        =   4
      Top             =   300
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   150
      TabIndex        =   3
      Top             =   285
      Width           =   1035
   End
End
Attribute VB_Name = "frmRefisDetalhe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Documento
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    sDataVencto As String
    sSit As String
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
End Type

Private Type TRIBUTO
    nNumDoc As Long
    nSeqDoc As Integer
    nCodReduz As Long
    nAno As Integer
    nLanc As Integer
    nSeq As Integer
    nParc As Integer
    nCompl As Integer
    nCodTrib As Integer
    nValorPrincipal As Double
    nValorMulta As Double
    nValorJuros As Double
    nValorCorrecao As Double
    nValorTotal As Double
End Type


Private Sub cmdPrint_Click()
If Not IsDate(mskDataIni.Text) Then
    MsgBox "Data de Inicio inválido", vbExclamation, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataFim.Text) Then
    MsgBox "Data de Fim inválido", vbExclamation, "atenção"
    Exit Sub
End If

If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    MsgBox "Data de Inicio tem que ser maior que data de termino", vbExclamation, "atenção"
    Exit Sub
End If
GeraResumo

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   Pb.value = (nPosF * 100) / nTotal
Else
   Pb.value = 100
End If
lblPB.Caption = FormatNumber(Pb.value, 2)

'Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub

Private Sub GeraResumo()

Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, nPlano As Integer, nPerc As Double
Pb.value = 0
nPos = 1
Ocupado
Sql = "select * from vwrefisnovo2 where datapagamento between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "' order by numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        nPlano = !plano
        Sql = "select desconto from plano where codigo=" & nPlano
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            nPerc = !desconto
           .Close
        End With
        
        
        Sql = "select * from vwdocumentodetalhe where numdocumento=" & !NumDocumento
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                MsgBox !CodTributo & " - " & !desctributo
                
               .MoveNext
            Loop
           .Close
        End With
        
        nPos = nPos + 1
        DoEvents
       .MoveNext
    Loop
   .Close
End With



Liberado

End Sub

Private Function CarregaParcela(nNumDoc As Long) As Documento
'Dim qd As New rdoQuery, RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, x As Integer, nFicha As Long, nFichaJM As Long, nFichaC As Long
'Dim nLinha As Integer, nValorPago As Double, sDataPag As String, sDup As String, sBax As String, bNewDoc As Boolean, nSoma As Double
'Dim nValorPrincipal As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double, nValorDif As Double
'Dim nValorComp As Double, nValorTarifa As Double, nValorChecar As Double, bIsento As Boolean, nPercDesconto As Double, RdoAux3 As rdoResultset, RdoAux4 As rdoResultset
'Dim nValorTarifaGlobal As Double, nLast As Integer, nQtdeLanc As Integer, aDocTmp() As Documento, nSeqReg As Integer, sNumProc As String, dDataVencto As Date
'Dim nCodReduz As Long, RdoEicon As rdoResultset
'
'Dim aDoc(1) As Documento, aTrib() As TRIBUTO
'ReDim aTrib(0)
'nLast = 1
'
'Sql = "SELECT parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento, debitoparcela.numerolivro, debitoparcela.paginalivro, debitoparcela.dataajuiza, "
'Sql = Sql & "parceladocumento.seqlancamento, parceladocumento.numparcela, parceladocumento.codcomplemento,"
'Sql = Sql & "parceladocumento.numdocumento,parceladocumento.plano, debitoparcela.datavencimento, debitoparcela.statuslanc, numdocumento.datadocumento,"
'Sql = Sql & "NumDocumento.valortaxadoc,numdocumento.percisencao FROM parceladocumento INNER JOIN debitoparcela ON parceladocumento.codreduzido = debitoparcela.codreduzido AND "
'Sql = Sql & "parceladocumento.anoexercicio = debitoparcela.anoexercicio AND parceladocumento.codlancamento = debitoparcela.codlancamento AND "
'Sql = Sql & "parceladocumento.seqlancamento = debitoparcela.seqlancamento AND parceladocumento.numparcela = debitoparcela.numparcela AND "
'Sql = Sql & "parceladocumento.codcomplemento = debitoparcela.codcomplemento INNER JOIN numdocumento ON parceladocumento.numdocumento = numdocumento.numdocumento "
'Sql = Sql & "WHERE PARCELADOCUMENTO.NumDocumento = " & nNumDoc
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    If .RowCount > 0 Then
'        If Val(SubNull(!plano)) > 0 Then
'            If !plano = 16 Then
'                nPercDesconto = 100
'            ElseIf !plano = 17 Then
'                nPercDesconto = 80
'            ElseIf !plano = 18 Then
'                nPercDesconto = 60
'            ElseIf !plano = 19 Then
'                nPercDesconto = 50
'            ElseIf !plano = 20 Then
'                nPercDesconto = 50
'            ElseIf !plano = 21 Then
'                nPercDesconto = 40
'            ElseIf !plano = 22 Then
'                nPercDesconto = 30
'            ElseIf !plano = 23 Then
'                nPercDesconto = 100
'            ElseIf !plano = 24 Then
'                nPercDesconto = 60
'            ElseIf !plano = 25 Then
'                nPercDesconto = 40
'            Else
'                nPercDesconto = 0
'            End If
'        Else
'            nPercDesconto = 0
'        End If
'        If nPercDesconto > 0 Then
'            bIsento = True
'        End If
'        sDataPag = Format(!DATADOCUMENTO, "dd/mm/yyyy")
'        Do Until .EOF
'            nCodReduz = !CODREDUZIDO
'            'CARREGA AS PARCELAS DO DOCUMENTO
'            On Error Resume Next
'            RdoAux2.Close
'            On Error GoTo 0
'            nValorPrincipal = 0: nValorMulta = 0: nValorJuros = 0: nValorCorrecao = 0: nValorTotal = 0
'            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
'            qd(0) = nCodReduz: qd(1) = nCodReduz
'            qd(2) = !AnoExercicio: qd(3) = !AnoExercicio
'            qd(4) = !CodLancamento: qd(5) = !CodLancamento
'            qd(6) = !SeqLancamento: qd(7) = !SeqLancamento
'            qd(8) = !NumParcela: qd(9) = !NumParcela
'            qd(10) = !CODCOMPLEMENTO: qd(11) = !CODCOMPLEMENTO
'            qd(12) = 1: qd(13) = 99: qd(14) = Format(sDataPag, "mm/dd/yyyy")
'            qd(15) = NomeDoUsuario
'            Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
'            Do Until RdoAux2.EOF
'                dDataVencto = RdoAux2!DataVencimentoCalc
'                nValorPrincipal = nValorPrincipal + RdoAux2!ValorTributo
'                If Not bIsento Then
'                    nValorMulta = RdoAux2!ValorMulta
'                    nValorJuros = RdoAux2!ValorJuros
'                Else
'                    nValorMulta = 0
'                    nValorJuros = 0
'                End If
'                nValorCorrecao = nValorCorrecao + CDbl(SubNull(RdoAux2!ValorCorrecao))
'
'                If Not bIsento Then
'                    nValorTotal = nValorTotal + (RdoAux2!ValorTributo + RdoAux2!ValorMulta + RdoAux2!ValorJuros + CDbl(SubNull(RdoAux2!ValorCorrecao)))
'                Else
'                    If nPercDesconto > 0 Then
'                        nValorMulta = nValorMulta + ((100 - nPercDesconto) * RdoAux2!ValorMulta / 100)
'                        nValorJuros = nValorJuros + ((100 - nPercDesconto) * RdoAux2!ValorJuros / 100)
'                        nValorTotal = nValorTotal + (RdoAux2!ValorTributo + CDbl(SubNull(RdoAux2!ValorCorrecao) + nValorMulta + nValorJuros))
'                   Else
'                        nValorTotal = nValorTotal + (RdoAux2!ValorTributo + CDbl(SubNull(RdoAux2!ValorCorrecao)))
'                   End If
'
'                End If
'                'Carrega os tributos
'                ReDim Preserve aTrib(UBound(aTrib) + 1)
'                nLast = UBound(aTrib)
'                aTrib(nLast).nNumDoc = nNumDoc
'                aTrib(nLast).nSeqDoc = aRegistro(nLinha).nSeq
'                aTrib(nLast).nCodReduz = nCodReduz
'                aTrib(nLast).nAno = !AnoExercicio
'                aTrib(nLast).nLanc = !CodLancamento
'                aTrib(nLast).nSeq = !SeqLancamento
'                aTrib(nLast).nParc = !NumParcela
'                aTrib(nLast).nCompl = !CODCOMPLEMENTO
'                aTrib(nLast).nCodTrib = RdoAux2!CodTributo
'                aTrib(nLast).nValorPrincipal = RdoAux2!ValorTributo
'                If Not bIsento Then
'                    aTrib(nLast).nValorMulta = RdoAux2!ValorMulta
'                    aTrib(nLast).nValorJuros = RdoAux2!ValorJuros
'                Else
'                    aTrib(nLast).nValorMulta = 0
'                    aTrib(nLast).nValorJuros = 0
'                End If
'                aTrib(nLast).nValorCorrecao = RdoAux2!ValorCorrecao
'                aTrib(nLast).nValorTotal = aTrib(nLast).nValorPrincipal + aTrib(nLast).nValorJuros + aTrib(nLast).nValorMulta + aTrib(nLast).nValorCorrecao + aTrib(nLast).nValorTarifa
'
'                nSoma = nSoma + aTrib(nLast).nValorTotal
'                nLast = nLast + 1
'                RdoAux2.MoveNext
'            Loop
'
'            nLast = 1
'            aDoc(nLast).nNumDoc = nNumDoc
'            aDoc(nLast).nSeqDoc = aRegistro(nLinha).nSeq
'            aDoc(nLast).nCodReduz = nCodReduz
'            aDoc(nLast).nAno = !AnoExercicio
'            aDoc(nLast).nLanc = !CodLancamento
'            aDoc(nLast).nSeq = !SeqLancamento
'            aDoc(nLast).nParc = !NumParcela
'            aDoc(nLast).nCompl = !CODCOMPLEMENTO
'            'aDoc(nLast).sDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
'            aDoc(nLast).sDataVencto = Format(dDataVencto, "dd/mm/yyyy")
'            aDoc(nLast).sSit = !statuslanc
'            aDoc(nLast).nNumeroLivro = Val(SubNull(!numerolivro))
'            aDoc(nLast).nPaginaLivro = Val(SubNull(!paginalivro))
'            aDoc(nLast).bAjuizado = IIf(IsNull(!DATAAJUIZA), False, True)
'            aDoc(nLast).nValorPrincipal = nValorPrincipal
'            aDoc(nLast).nValorMulta = nValorMulta
'            aDoc(nLast).nValorJuros = nValorJuros
'            aDoc(nLast).nValorCorrecao = nValorCorrecao
'            aDoc(nLast).nValorTotal = nValorTotal
'           .MoveNext
'        Loop
'    End If
'    .Close
'End With

'CarregaParcela = aDoc(1)

End Function


