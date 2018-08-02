VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmBaixaManual 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa Manual"
   ClientHeight    =   5310
   ClientLeft      =   1110
   ClientTop       =   2730
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5310
   ScaleWidth      =   10350
   Begin esMaskEdit.esMaskedEdit mskCodBarra 
      Height          =   315
      Left            =   3945
      TabIndex        =   7
      Top             =   120
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   556
      MouseIcon       =   "frmBaixaManual.frx":0000
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
      MaxLength       =   61
      Mask            =   "###########-# | ###########-# | ###########-# | ###########-#"
      SelText         =   ""
      Text            =   "___________-_ | ___________-_ | ___________-_ | ___________-_"
      HideSelection   =   -1  'True
   End
   Begin VB.TextBox txtNumDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1305
      TabIndex        =   0
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cmbBanco 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmBaixaManual.frx":001C
      Left            =   1770
      List            =   "frmBaixaManual.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   4200
      Width           =   2475
   End
   Begin VB.TextBox txtAgencia 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1770
      TabIndex        =   4
      Top             =   4590
      Width           =   1275
   End
   Begin VB.TextBox txtValorPago 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1770
      TabIndex        =   5
      Top             =   4920
      Width           =   1275
   End
   Begin MSFlexGridLib.MSFlexGrid grdParc 
      Height          =   2925
      Left            =   30
      TabIndex        =   6
      Top             =   510
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   5159
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedCols       =   0
      BackColorSel    =   12582912
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmBaixaManual.frx":0020
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   8910
      TabIndex        =   24
      ToolTipText     =   "Sair da Tela"
      Top             =   4860
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
      MICON           =   "frmBaixaManual.frx":010B
      PICN            =   "frmBaixaManual.frx":0127
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   345
      Left            =   8910
      TabIndex        =   25
      ToolTipText     =   "Efetuar Baixa"
      Top             =   4470
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Efetuar Baixa"
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
      MICON           =   "frmBaixaManual.frx":0195
      PICN            =   "frmBaixaManual.frx":01B1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdLoadDeb 
      Height          =   300
      Left            =   3030
      TabIndex        =   26
      ToolTipText     =   "Carregar os Débitos"
      Top             =   3525
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   "&Carregar"
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
      MICON           =   "frmBaixaManual.frx":0250
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataPag 
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      Top             =   3540
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      MouseIcon       =   "frmBaixaManual.frx":026C
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
   Begin esMaskEdit.esMaskedEdit mskDataCred 
      Height          =   285
      Left            =   1770
      TabIndex        =   2
      Top             =   3870
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   503
      MouseIcon       =   "frmBaixaManual.frx":0288
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Taxa Doc.........:"
      Height          =   225
      Index           =   6
      Left            =   4410
      TabIndex        =   29
      Top             =   4110
      Width           =   1605
   End
   Begin VB.Label lblValorTaxa 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6090
      TabIndex        =   28
      Top             =   4110
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   195
      TabIndex        =   27
      Top             =   180
      Width           =   1035
   End
   Begin VB.Label lblDup 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6090
      TabIndex        =   23
      Top             =   4980
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pagto.Duplicado........:"
      Height          =   225
      Index           =   10
      Left            =   4410
      TabIndex        =   22
      Top             =   4980
      Width           =   1620
   End
   Begin VB.Label lblValLanc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6090
      TabIndex        =   21
      Top             =   4395
      Width           =   1185
   End
   Begin VB.Label lblNumLanc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6090
      TabIndex        =   20
      Top             =   3540
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total.................:"
      Height          =   225
      Index           =   9
      Left            =   4410
      TabIndex        =   19
      Top             =   4395
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Lançamentos..:"
      Height          =   225
      Index           =   8
      Left            =   4410
      TabIndex        =   18
      Top             =   3540
      Width           =   1605
   End
   Begin VB.Label lblValorDif 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6090
      TabIndex        =   17
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Diferença..........:"
      Height          =   225
      Index           =   7
      Left            =   4410
      TabIndex        =   16
      Top             =   4680
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Pago................:"
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   4950
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Agência....................:"
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   4620
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Crédito.........:"
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   3900
      Width           =   1605
   End
   Begin VB.Label lblValorCalc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6090
      TabIndex        =   12
      Top             =   3825
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Lançamento.....:"
      Height          =   225
      Index           =   2
      Left            =   4410
      TabIndex        =   11
      Top             =   3825
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco.......................:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   4260
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Pagamento..:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3570
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cód.Barras:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   2820
      TabIndex        =   8
      Top             =   180
      Width           =   1035
   End
End
Attribute VB_Name = "frmBaixaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type CodigoBarra
   PreCodBarra As String * 4
   ValorRecebido As String * 11
   CodigoMunic As String * 4
   DataVencto As String * 8
   NumDocumento As String * 9
   NumParcela As String * 2
   SituacaoRetorno As String * 2
   FillerSmar As String * 4
End Type

Dim sBloco1 As String, sBloco2 As String, sBloco3 As String, sBloco4 As String, sBloco5 As String
Dim sBloco As String

Dim aCodBarra() As CodigoBarra
Dim RdoAux As rdoResultset
Dim RdoAux2 As rdoResultset
Dim Sql As String, bExec As Boolean
Dim nCodReduz As Long, nAnoExercicio As Integer, nCodLanc As Integer
Dim nSeqLanc As Integer, nNumParc As Integer, nCodTributo As Integer
Dim nCompl As Integer, nStatus As Integer, sStatus As String

Private Sub cmdBaixa_Click()

If txtNumDoc.text = "" Then
    MsgBox "Selecione o Documento.", vbExclamation, "Atenção"
    Exit Sub
End If

If grdParc.Rows = 1 Then
    MsgBox "Carregue os lançamentos", vbExclamation, "Atenção"
    cmdLoadDeb.SetFocus
    Exit Sub
End If

If Not IsDate(mskDataPag.text) Then
   MsgBox "Data de Pagamento inválida.", vbExclamation, "Atenção"
   mskDataPag.SetFocus
   Exit Sub
End If

If Not IsDate(mskDataCred.text) Then
   MsgBox "Data de Crédito inválida.", vbExclamation, "Atenção"
   mskDataCred.SetFocus
   Exit Sub
End If

If CDate(mskDataCred.text) < CDate(mskDataPag.text) Then
   MsgBox "Data de crédito não pode ser menor que a data de pagamento.", vbExclamation, "Atenção"
   mskDataCred.SetFocus
   Exit Sub
End If
If cmbBanco.ListIndex = -1 Then
   MsgBox "Selecione um Banco.", vbExclamation, "Atenção"
  cmbBanco.SetFocus
   Exit Sub
End If
If Val(txtAgencia.text) = 0 Then
   txtAgencia.text = 0
End If

If CDbl(txtValorPago.text) = 0 Then
   MsgBox "Digite o Valor Pago.", vbExclamation, "Atenção"
   txtValorPago.SetFocus
   Exit Sub
End If

If Val(lblNumLanc.Caption) = 0 Then
   MsgBox "Carregue os Lançamentos.", vbExclamation, "Atenção"
   cmdLoadDeb.SetFocus
   Exit Sub
End If


Sql = "SELECT NUMDOCUMENTO,VALORPAGO FROM NUMDOCUMENTO "
Sql = Sql & " WHERE NUMDOCUMENTO=" & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If !ValorPago > 0 Then
        If MsgBox("Já foi efetuado a baixa para este Documento." & vbCrLf & "A baixa será entendida como um pagamento em duplicidade." & vbCrLf & "Deseja Continuar ?", vbQuestion + vbYesNo, "Atenção") = vbNo Then
           Exit Sub
        End If
    End If
   .Close
End With

If lblDup.Caption <> "Não" Then
   If MsgBox("Deseja efetuar baixa deste Documento ? " & vbCrLf & "ATENÇÃO: EXISTE(M) LANÇAMENTO(S) EM DUPLICIDADE !", vbQuestion + vbYesNo, "CONFIRMAÇÃO") = vbNo Then
      Exit Sub
   End If
Else
   If MsgBox("Deseja efetuar baixa deste Documento ? ", vbQuestion + vbYesNo, "CONFIRMAÇÃO") = vbNo Then
      Exit Sub
   End If
End If

Ocupado

For x = 1 To grdParc.Rows - 1
    If Left$(grdParc.TextMatrix(x, 2), 3) = "005" Then
       grdParc.TextMatrix(x, 6) = CDbl(grdParc.TextMatrix(x, 6)) + (CDbl(txtValorPago.text))
       grdParc.TextMatrix(x, 10) = grdParc.TextMatrix(x, 6)
       nCodReduz = grdParc.TextMatrix(x, 1)
       nAnoExercicio = grdParc.TextMatrix(x, 0)
       nCodLanc = Val(Left$(grdParc.TextMatrix(x, 2), 3))
       nSeqLanc = grdParc.TextMatrix(x, 3)
       nNumParc = grdParc.TextMatrix(x, 4)
       nCompl = grdParc.TextMatrix(x, 5)
       Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=" & Virg2Ponto(txtValorPago.text - lblValorTaxa.Caption)
       Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
       Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
       Sql = Sql & " AND NUMPARCELA=" & nNumParc & " AND CODTRIBUTO=13"
       cn.Execute Sql, rdExecDirect
       
       lblValorCalc.Caption = FormatNumber(txtValorPago.text - lblValorTaxa.Caption, 2)
       lblValLanc.Caption = txtValorPago.text
       lblValorDif.Caption = "0,00"
       
       Exit For
    End If
Next

EfetuaBaixa
Liberado
MsgBox "Baixa efetuada com sucesso.", vbExclamation, "Atenção"
'Limpa
End Sub

Private Sub EfetuaBaixa()
Dim x As Integer
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset, bJuros As Boolean, bMulta As Boolean
Dim nValorLanc As Double, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double
Dim nTotal As Double
Dim dDataBase As Date, dDataPag As Date, dDataVencto As Date, nValorPago As Double
Dim nSeqAdd As Integer, nSomaDoc As Double, sCodAgencia As String

With grdParc
    nValorTaxa = CDbl(lblValorTaxa.Caption)
    nTotal = CDbl(txtValorPago.text)
    nValorLanc = 0
    For x = 1 To .Rows - 1
        nValorLanc = nValorLanc + CDbl(.TextMatrix(x, 13))
    Next
    If nValorLanc > nTotal Then
       .TextMatrix(1, 13) = Abs(CDbl(.TextMatrix(1, 13)) - (nValorLanc - nTotal))
    ElseIf nValorLanc < nTotal Then
       .TextMatrix(1, 13) = CDbl(.TextMatrix(1, 13)) + (nTotal - nValorLanc)
    End If
    
End With

'Exit Sub
'EFETUA AS BAIXAS
dDataPag = CDate(mskDataPag.text)
sCodAgencia = txtAgencia.text
nSomaDoc = CDbl(lblValLanc.Caption)

With grdParc
    For x = 1 To .Rows - 1
        If .TextMatrix(x, 1) <> "N/A" Then
             nCodReduz = .TextMatrix(x, 1)
             nAnoExercicio = .TextMatrix(x, 0)
             nCodLanc = Val(Left$(.TextMatrix(x, 2), 3))
             nSeqLanc = .TextMatrix(x, 3)
             nNumParc = .TextMatrix(x, 4)
             nCompl = .TextMatrix(x, 5)
             dDataVencto = CDate(.TextMatrix(x, 12))
             nValorLanc = .TextMatrix(x, 10)
             nValorPago = CDbl(txtValorPago.text)
             nTotal = CDbl(.TextMatrix(x, 13))
             nStatus = 0
             Sql = "SELECT CODREDUZIDO,STATUSLANC FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND "
             Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND "
             Sql = Sql & "NUMPARCELA=" & nNumParc & " AND CODCOMPLEMENTO=" & nCompl
             Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             If RdoAux.RowCount > 0 Then
                If RdoAux!statuslanc = 5 Or RdoAux!statuslanc = 10 Or RdoAux!statuslanc = 12 Or RdoAux!statuslanc = 8 Then
                    nStatus = 15
                End If
             End If
             RdoAux.Close
             
             If nStatus <> 15 Then
                If nNumParc = 0 Then
                   If CDbl(lblValorDif.Caption) <= 0 Then
                       nStatus = 1 'UNICA SEM DIF
                   Else
                       nStatus = 9 'UNICA COM DIF
                   End If
                Else
                   'If nTotal - nValorLanc = 0 Then
                   If CDbl(lblValorDif.Caption) <= 0 Then
                       nStatus = 2 'PAGO SEM DIF
                   Else
                       nStatus = 7 'PAGO COM DIF
                   End If
                End If
             End If
             
             If UCase$(.TextMatrix(x, 11)) <> "SIM" Then 'não é duplicado
             
                'EFETUA BAIXA NA TABELA NUMDOCUMENTO
                 Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=" & cmbBanco.ItemData(cmbBanco.ListIndex) & " ,CODAGENCIA ='" & sCodAgencia & "' , VALORPAGO=" & Virg2Ponto(sTr(CDbl(txtValorPago.text)))
                 Sql = Sql & " WHERE NUMDOCUMENTO=" & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1))
                 cn.Execute Sql, rdExecDirect
                'EFETUA BAIXA NA TABELA DEBITOPARCELA
                 nValorLanc = grdParc.TextMatrix(x, 13)
                 Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=" & nStatus & " WHERE CODREDUZIDO=" & nCodReduz & " AND "
                 Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                 Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                 cn.Execute Sql, rdExecDirect
                'EFETUA BAIXA NA TABELA DEBITOPAGO
                 Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND "
                 Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                 Sql = Sql & " AND NUMPARCELA=" & nNumParc
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                      If IsNull(!MAXIMO) Then
                         nSeqAdd = 0
                      Else
                         If .RowCount = 0 Then
                            nSeqAdd = 0
                        Else
                           nSeqAdd = !MAXIMO + 1
                        End If
                     End If
                    .Close
                End With
                nValorLanc = grdParc.TextMatrix(x, 10)
                Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL) VALUES(" & nCodReduz & ","
                Sql = Sql & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nSeqAdd & ",'"
                'Sql = Sql & Format(dDataPag, "mm/dd/yyyy") & "','" & Format(mskDataCred.text, "mm/dd/yyyy") & "'," & Virg2Ponto(Str(nTotal)) & ","
                Sql = Sql & Format(dDataPag, "mm/dd/yyyy") & "','" & Format(mskDataCred.text, "mm/dd/yyyy") & "'," & Virg2Ponto(sTr(nValorLanc)) & ","
                If nValorTaxa > 3 Then
                   nValorTaxa = CDbl(lblValorDif.Caption)
                   Sql = Sql & Val(cmbBanco.ItemData(cmbBanco.ListIndex)) & "," & Val(txtAgencia.text) & "," & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1)) & "," & Virg2Ponto(sTr(nTotal)) & ")"
                   'Sql = Sql & Val(cmbBanco.ItemData(cmbBanco.ListIndex)) & "," & Val(txtAgencia.text) & "," & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1)) & "," & Virg2Ponto(Str(nValorLanc)) & ")"
                Else
                   Sql = Sql & Val(cmbBanco.ItemData(cmbBanco.ListIndex)) & "," & Val(txtAgencia.text) & "," & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1)) & "," & Virg2Ponto(sTr(nTotal)) & ")"
                End If
                cn.Execute Sql, rdExecDirect
                        
                'SE FOR PARCELA UNICA EFETUA BAIXA EM TODAS AS PARCELAS AUTOMATICAMENTO
                'SERA?
                 If nNumParc = 0 Then
                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=1  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                    Sql = Sql & " AND NUMPARCELA<>0"
                    cn.Execute Sql, rdExecDirect
                 Else
                    Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=5  WHERE CODREDUZIDO=" & nCodReduz & " AND "
                    Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                    Sql = Sql & " AND NUMPARCELA=0"
                    cn.Execute Sql, rdExecDirect
                 End If
                
                
                'EFETUA BAIXA NA TABELA DEBITOTRIBUTO
                 Sql = "SELECT DEBITOPARCELA.DATADEBASE,CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO INNER JOIN DEBITOPARCELA ON DEBITOTRIBUTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND "
                 Sql = Sql & "DEBITOTRIBUTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND DEBITOTRIBUTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND "
                 Sql = Sql & "DEBITOTRIBUTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND DEBITOTRIBUTO.NumParcela = DEBITOPARCELA.NumParcela AND "
                 Sql = Sql & "DEBITOTRIBUTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO "
                 Sql = Sql & " WHERE DEBITOPARCELA.CODREDUZIDO=" & nCodReduz & " AND DEBITOPARCELA.ANOEXERCICIO=" & nAnoExercicio & " AND DEBITOPARCELA.CODLANCAMENTO=" & nCodLanc & " AND DEBITOPARCELA.SEQLANCAMENTO=" & nSeqLanc & " AND DEBITOPARCELA.NUMPARCELA=" & nNumParc & " AND "
                 Sql = Sql & "DEBITOPARCELA.CODCOMPLEMENTO=" & nCompl
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                        Do Until .EOF
                            If !CodTributo = 13 Then
                                Sql = "UPDATE DEBITOTRIBUTO SET VALORTRIBUTO=" & Virg2Ponto(sTr(nValorPago - nValorTaxa))
                                Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                                Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=" & nCodTributo
                                cn.Execute Sql, rdExecDirect
                            End If
                            
                            bJuros = False: bMulta = False
                            Sql = "SELECT MULTA,JUROS FROM TRIBUTO WHERE CODTRIBUTO=" & !CodTributo
                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                            With RdoAux2
                                If .RowCount > 0 Then
                                    bJuros = !Juros
                                    bMulta = !Multa
                                End If
                               .Close
                            End With
                            
                            dDataBase = Format(!DATADEBASE, "dd/mm/yyyy")
                            nValorLanc = FormatNumber(!valortributo, 2)
                            nCodTributo = !CodTributo
                            nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorLanc, dDataVencto, dDataPag), 2)
                            nValorJuros = 0: nValorMulta = 0
                            If bJuros Then
                                nValorJuros = FormatNumber(CalculaJuros2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                            End If
                            If bMulta Then
                                nValorMulta = FormatNumber(CalculaMulta2(nValorLanc + nValorCorrecao, dDataVencto, dDataPag), 2)
                            End If
                            Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=" & Virg2Ponto(sTr(nValorCorrecao)) & " ,VALORMULTA=" & Virg2Ponto(sTr(nValorMulta)) & " ,VALORJUROS=" & Virg2Ponto(sTr(nValorJuros))
                            Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND NUMPARCELA=" & nNumParc & " AND "
                            Sql = Sql & "CODCOMPLEMENTO=" & nCompl & " AND CODTRIBUTO=" & nCodTributo
                            cn.Execute Sql, rdExecDirect
                           .MoveNext
                        Loop
                       .Close
                 End With
             Else
               'EFETUA BAIXA NA TABELA NUMDOCUMENTO
                Sql = "UPDATE NUMDOCUMENTO SET CODBANCO=" & cmbBanco.ItemData(cmbBanco.ListIndex) & " ,CODAGENCIA ='" & sCodAgencia & "' , VALORPAGO=" & Virg2Ponto(sTr(CDbl(txtValorPago.text)))
                Sql = Sql & " WHERE NUMDOCUMENTO=" & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1))
                cn.Execute Sql, rdExecDirect
               'GRAVA DEBITO ADICIONAL
                'EFETUA BAIXA NA TABELA DEBITOPAGO
                 Sql = "SELECT MAX(SEQPAG) AS MAXIMO FROM DEBITOPAGO WHERE CODREDUZIDO=" & nCodReduz & " AND "
                 Sql = Sql & "ANOEXERCICIO=" & nAnoExercicio & " AND CODLANCAMENTO=" & nCodLanc & " AND SEQLANCAMENTO=" & nSeqLanc & " AND CODCOMPLEMENTO=" & nCompl
                 Sql = Sql & " AND NUMPARCELA=" & nNumParc
                 Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux2
                      If IsNull(!MAXIMO) Then
                         nSeqAdd = 0
                      Else
                         If .RowCount = 0 Then
                            nSeqAdd = 0
                        Else
                           nSeqAdd = !MAXIMO + 1
                        End If
                     End If
                    .Close
                End With
                nValorLanc = grdParc.TextMatrix(x, 13)
                Sql = "INSERT DEBITOPAGO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,"
                Sql = Sql & "SEQPAG,DATAPAGAMENTO,DATARECEBIMENTO,VALORPAGO,CODBANCO,CODAGENCIA,NUMDOCUMENTO,VALORPAGOREAL) VALUES(" & nCodReduz & ","
                Sql = Sql & nAnoExercicio & "," & nCodLanc & "," & nSeqLanc & "," & nNumParc & "," & nCompl & "," & nSeqAdd & ",'"
                Sql = Sql & Format(dDataPag, "mm/dd/yyyy") & "','" & Format(mskDataCred.text, "mm/dd/yyyy") & "'," & Virg2Ponto(sTr(nValorLanc)) & ","
                'Sql = Sql & Val(cmbBanco.ItemData(cmbBanco.ListIndex)) & "," & Val(txtAgencia.text) & "," & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1)) & "," & Virg2Ponto(sTr(nValorLanc + nValorTaxa)) & ")"
                Sql = Sql & Val(cmbBanco.ItemData(cmbBanco.ListIndex)) & "," & Val(txtAgencia.text) & "," & Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1)) & "," & Virg2Ponto(sTr(nValorLanc)) & ")"
                cn.Execute Sql, rdExecDirect
             End If
        End If
    Next
End With

End Sub

Private Sub Executa()
Dim nNumDoc As Long
bExec = False
Limpa
If Val(txtNumDoc.text) = 0 And mskCodBarra.ClipText = "" Then
'   MsgBox "Digite o nº do documento ou o código de barras.", vbExclamation, "Atenção"
   Exit Sub
End If

If Val(txtNumDoc.text) > 990000000 Then
   MsgBox "Documento Inválido.", vbExclamation, "Atenção"
   Exit Sub
End If

If Val(txtNumDoc.text) = 0 Then
    If Not ValidaCodBarra(mskCodBarra.text) Then
       MsgBox "Código de Barra Inválido.", vbExclamation, "Atenção"
       mskCodBarra.SetFocus
       Exit Sub
    End If
    LoadMatrix
    If txtNumDoc.text = "" Then
       txtNumDoc.text = Val(aCodBarra(0).NumDocumento)
    End If
Else
   nNumDoc = Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1))
   If Val(Right$(txtNumDoc.text, 1)) <> RetornaDVNumDoc(nNumDoc) Then
       MsgBox "Digito Verificador Inválido", vbExclamation, "Atenção"
       Exit Sub
   End If
End If

bExec = True
grdParc.SetFocus
MontaResumo

End Sub


Private Sub cmdLoadDeb_Click()
If Not IsDate(mskDataPag.text) Then
   MsgBox "Data de Pagamento inválida.", vbExclamation, "Atenção"
   mskDataPag.SetFocus
   Exit Sub
End If

Executa
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Ocupado
frmMdi.AddWindow Me.Name, Me.Caption

Centraliza Me

Sql = "SELECT CODBANCO,NOMEBANCO FROM BANCO WHERE CODBANCO<>0 ORDER BY NOMEBANCO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         cmbBanco.AddItem !NomeBanco
         cmbBanco.ItemData(cmbBanco.NewIndex) = !CodBanco
        .MoveNext
    Loop
End With

Liberado

End Sub

Private Function ValidaCodBarra(sCodBarra As String) As Boolean
Dim nRetVal As Integer
sCodBarra = RetornaNumero(sCodBarra)
ValidaCodBarra = True
If Len(sCodBarra) < 48 Then
   ValidaCodBarra = False
   Exit Function
End If

sBloco1 = Left$(sCodBarra, 3)
sBloco2 = Mid$(sCodBarra, 5, 7)
sBloco3 = Mid$(sCodBarra, 13, 11)
sBloco4 = Mid$(sCodBarra, 25, 11)
sBloco5 = Mid$(sCodBarra, 37, 11)

sBloco = sBloco1 & sBloco2 & sBloco3 & sBloco4 & sBloco5
nRetVal = RetornaDV2of5(sBloco)
If nRetVal = 10 Then nRetVal = 0
If nRetVal <> Val(Mid$(sCodBarra, 4, 1)) Then
   ValidaCodBarra = False
   Exit Function
End If

End Function

Private Function RetornaDV2of5(sBloco As String) As Integer
Dim c As Integer
Dim d As Integer
Dim e As String
Dim nSoma As Integer
Dim nResto As Integer

For c = Len(sBloco) To 1 Step -1
      If c Mod 2 = 1 Then
         d = Val(Mid(sBloco, c, 1)) * 2
      Else
         d = Val(Mid(sBloco, c, 1)) * 1
      End If
      If d > 0 Then
         If d > 9 Then
            e = CStr(d)
            d = Val(Left$(e, 1)) + Val(Right$(e, 1))
         End If
         nSoma = nSoma + d
      End If
Next

nResto = nSoma Mod 10
RetornaDV2of5 = 10 - nResto

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
End Sub

Private Sub Limpa()
LimpaMascara mskDataCred
'LimpaMascara mskDataPag
cmbBanco.ListIndex = -1
txtAgencia.text = ""
txtValorPago.text = ""
lblNumLanc.Caption = 0
lblValorCalc.Caption = "0,00"
lblValLanc.Caption = "0,00"
lblValorTaxa.Caption = "0,00"
lblDup.Caption = ""
lblValorDif.Caption = "0,00"
grdParc.Rows = 1
End Sub

Private Sub LoadMatrix()
ReDim aCodBarra(0)
With aCodBarra(0)
    .PreCodBarra = Mid$(sBloco, 1, 3)
    .ValorRecebido = Mid$(sBloco, 4, 11)
    .CodigoMunic = Mid$(sBloco, 15, 4)
    .DataVencto = Mid$(sBloco, 19, 8)
    .NumDocumento = Mid$(sBloco, 27, 9)
    .NumParcela = Mid$(sBloco, 36, 2)
    .SituacaoRetorno = Mid$(sBloco, 38, 2)
    .FillerSmar = Mid$(sBloco, 40, 4)
End With

End Sub

Private Sub MontaResumo()
Dim x As Long
Dim nNumDoc As Long, dDataDoc As Date
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim nStatus As Integer
Dim nValorJuros As Double
Dim nValorMulta As Double
Dim nValorCorrecao As Double
Dim bDupl As Boolean
Dim dDataBase As Date
Dim dDataPag As Date
Dim dDataVencto As Date
Dim nSoma As Double
Dim nValorPrincipal As Double
Dim nSomaTotal As Double
Dim nSomaPrincipal As Double
Dim bDupS As Boolean, bDupN As Boolean, bCalculaJurosMulta As Boolean

grdParc.Rows = 1
nNumDoc = Val(Left$(txtNumDoc.text, Len(txtNumDoc.text) - 1))
dDataPag = CDate(mskDataPag.text)
'CARREGA OS DEBITOS DESTE DOCUMENTO
Sql = "SELECT PARCELADOCUMENTO.CODREDUZIDO,PARCELADOCUMENTO.ANOEXERCICIO,PARCELADOCUMENTO.CODLANCAMENTO,LANCAMENTO.DESCREDUZ,PARCELADOCUMENTO.SEQLANCAMENTO,"
Sql = Sql & "PARCELADOCUMENTO.NUMPARCELA,PARCELADOCUMENTO.CODCOMPLEMENTO,PARCELADOCUMENTO.NUMDOCUMENTO,NUMDOCUMENTO.DATADOCUMENTO,NUMDOCUMENTO.CODBANCO,NUMDOCUMENTO.VALORTAXADOC,"
Sql = Sql & "PARCELADOCUMENTO.VALORJUROS,PARCELADOCUMENTO.VALORMULTA,PARCELADOCUMENTO.VALORCORRECAO,"
Sql = Sql & "NUMDOCUMENTO.CODAGENCIA,NUMDOCUMENTO.VALORPAGO,DEBITOPARCELA.STATUSLANC,SITUACAOLANCAMENTO.DESCSITUACAO,DEBITOPARCELA.DATAVENCIMENTO,DEBITOPARCELA.DATADEBASE "
Sql = Sql & "FROM PARCELADOCUMENTO INNER JOIN DEBITOPARCELA ON PARCELADOCUMENTO.CODREDUZIDO = DEBITOPARCELA.CODREDUZIDO AND PARCELADOCUMENTO.ANOEXERCICIO = DEBITOPARCELA.ANOEXERCICIO AND "
Sql = Sql & "PARCELADOCUMENTO.CODLANCAMENTO = DEBITOPARCELA.CODLANCAMENTO AND PARCELADOCUMENTO.SEQLANCAMENTO = DEBITOPARCELA.SEQLANCAMENTO AND PARCELADOCUMENTO.NumParcela = DEBITOPARCELA.NumParcela AND "
Sql = Sql & "PARCELADOCUMENTO.CODCOMPLEMENTO = DEBITOPARCELA.CODCOMPLEMENTO Inner Join LANCAMENTO ON DEBITOPARCELA.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO Inner Join SITUACAOLANCAMENTO ON "
Sql = Sql & "DEBITOPARCELA.STATUSLANC = SITUACAOLANCAMENTO.CODSITUACAO Inner Join NUMDOCUMENTO ON PARCELADOCUMENTO.NUMDOCUMENTO = NUMDOCUMENTO.NUMDOCUMENTO Where PARCELADOCUMENTO.NumDocumento = " & nNumDoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
'        If .RowCount > 1 Then
            bCalculaJurosMulta = False
'        Else
'            bCalculaJurosMulta = True
'        End If
        MsgBox "Documento não encontrado.", vbExclamation, "Atenção"
        If mskCodBarra.ClipText <> "" Then
           mskCodBarra.Locked = False
           mskCodBarra.BackColor = Branco
        Else
           txtNumDoc.Locked = False
           txtNumDoc.BackColor = Branco
        End If
        Exit Sub
    Else
        If Not IsNull(!DATADOCUMENTO) Then
            dDataDoc = !DATADOCUMENTO
        End If
        bCalculaJurosMulta = True
    End If
    'SE NÃO TIVER TAXADOC SINAL QUE VEIO DA SMARK ENTÃO PEGAMOS A TAXADOC DO 1º LANCAMENTO
    Sql = "SELECT * FROM DEBITOTRIBUTO "
    Sql = Sql & "WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
    Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO=3 "
    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux2.RowCount > 0 Then
        If IsNull(!VALORTAXADOC) Or !VALORTAXADOC = 0 Then
           Sql = "SELECT CODTRIBUTO,VALORTRIBUTO FROM DEBITOTRIBUTO WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
           Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO=3"
           Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux2
               If .RowCount > 0 Then
                  lblValorTaxa.Caption = FormatNumber(!valortributo, 2)
               Else
                  lblValorTaxa.Caption = "0,00"
               End If
           End With
        Else
           lblValorTaxa.Caption = FormatNumber(!VALORTAXADOC, 2)
        End If
    Else
        Sql = "SELECT VALORTAXADOC FROM NUMDOCUMENTO WHERE NUMDOCUMENTO = " & nNumDoc
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux2.RowCount > 0 Then
            lblValorTaxa.Caption = Format(!VALORTAXADOC, "#0.00")
        Else
            lblValorTaxa.Caption = "0,00"
        End If
        
    End If
    Do Until .EOF
        nStatus = !statuslanc
        sStatus = !DescSituacao
        dDataBase = Format(!DATADEBASE, "dd/mm/yyyy")
        dDataVencto = Format(!DataVencimento, "dd/mm/yyyy")
        
        Sql = "SELECT SUM(VALORTRIBUTO) AS TOTAL, SUM(VALORJUROS) AS JUROS, SUM(VALORMULTA) AS MULTA, SUM(VALORCORRECAO) AS CORRECAO FROM DEBITOTRIBUTO "
        Sql = Sql & "WHERE CODREDUZIDO = " & !CODREDUZIDO & " AND ANOEXERCICIO = " & !AnoExercicio & " AND CODLANCAMENTO = " & !CodLancamento & " AND "
        Sql = Sql & "SEQLANCAMENTO = " & !SeqLancamento & " AND NUMPARCELA = " & !NumParcela & " AND CODCOMPLEMENTO = " & !CODCOMPLEMENTO & " AND CODTRIBUTO<>3 AND CODTRIBUTO<>3 "
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If Not IsNull(!Total) Then
               nValorPrincipal = !Total
               nValorJuros = !Juros
               nValorMulta = !Multa
               nValorCorrecao = !CORRECAO
            Else
               nValorPrincipal = 0
               nValorMulta = 0
               nValorJuros = 0
               nValorCorrecao = 0
            End If
           .Close
        End With
        
        nValorCorrecao = FormatNumber(CalculaCorrecao2(nValorPrincipal, dDataVencto, dDataPag), 2)
        If bCalculaJurosMulta Then
            If nStatus = 20 Then 'JULGAMENTO
                nValorJuros = FormatNumber(CalculaJuros2(nValorPrincipal + nValorCorrecao, dDataVencto, dDataDoc), 2)
            Else
                nValorJuros = FormatNumber(CalculaJuros2(nValorPrincipal + nValorCorrecao, dDataVencto, dDataPag), 2)
            End If
            nValorMulta = FormatNumber(CalculaMulta2(nValorPrincipal + nValorCorrecao, dDataVencto, dDataPag), 2)
        Else
            nValorJuros = 0
            nValorMulta = 0
        End If
        nValorTotal = nValorPrincipal + nValorCorrecao + nValorJuros + nValorMulta
        nSomaTotal = nSomaTotal + nValorTotal
        nSomaPrincipal = nSomaPrincipal + nValorPrincipal
         
        If nStatus = 1 Or nStatus = 2 Or nStatus = 7 Or nStatus = 9 Then
           bDupS = True
           bDupl = True
        Else
           bDupN = True
           bDupl = False
        End If
        
        grdParc.AddItem !AnoExercicio & Chr(9) & Format(!CODREDUZIDO, "000000") & Chr(9) & Format(!CodLancamento, "000") & " - " & !descreduz & Chr(9) & Format(!SeqLancamento, "00") & Chr(9) & Format(!NumParcela, "00") & Chr(9) & _
           !CODCOMPLEMENTO & Chr(9) & FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorMulta, 2) & Chr(9) & _
           FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & FormatNumber(nValorTotal, 2) & Chr(9) & IIf(bDupl, "Sim", "Não") & Chr(9) & Format(!DataVencimento, "dd/mm/yyyy")
       .MoveNext
    Loop
   .Close
End With
Fim:

bDupl = False
nSoma = 0
nSomaTaxa = FormatNumber((CDbl(lblValorTaxa.Caption) / (grdParc.Rows - 1)), 2)
With grdParc
    For x = 1 To .Rows - 1
       .TextMatrix(x, 14) = nSomaTaxa
       .TextMatrix(x, 10) = FormatNumber(CDbl(.TextMatrix(x, 6)) + CDbl(.TextMatrix(x, 7)) + CDbl(.TextMatrix(x, 8)) + CDbl(.TextMatrix(x, 9)), 2)
        nSoma = nSoma + CDbl(.TextMatrix(x, 10))
    Next
End With

If (nSomaTaxa * (grdParc.Rows - 1)) > CDbl(lblValorTaxa.Caption) Then
    grdParc.TextMatrix(grdParc.Rows - 1, 14) = grdParc.TextMatrix(grdParc.Rows - 1, 14) - ((nSomaTaxa * (grdParc.Rows - 1)) - CDbl(lblValorTaxa.Caption))
End If

With grdParc
    For x = 1 To .Rows - 1
       .TextMatrix(x, 10) = FormatNumber(CDbl(.TextMatrix(x, 6)) + CDbl(.TextMatrix(x, 7)) + CDbl(.TextMatrix(x, 8)) + CDbl(.TextMatrix(x, 9)), 2)
       .TextMatrix(x, 13) = FormatNumber(CDbl(.TextMatrix(x, 10)) + CDbl(.TextMatrix(x, 14)), 2)
    Next
End With


lblValLanc.Caption = FormatNumber(nSoma + CDbl(lblValorTaxa.Caption), 2)
'lblValLanc.Caption = FormatNumber(nSoma, 2)
lblValorCalc.Caption = FormatNumber(nSoma, 2)
lblNumLanc.Caption = Format(grdParc.Rows - 1, "00")
If bDupS = False Then
   lblDup.Caption = "Não"
ElseIf bDupS = True And bDupN = False Then
   lblDup.Caption = "Sim"
ElseIf bDupS = True And bDupN = True Then
   lblDup.Caption = "Parcial"
End If

lblValorDif.Caption = lblValLanc.Caption

End Sub

Private Sub mskCodBarra_Change()

If Val(mskCodBarra.ClipText) > 0 Then
    If bExec Then txtNumDoc.text = ""
    txtNumDoc.Locked = True
    txtNumDoc.BackColor = Kde
Else
    txtNumDoc.Locked = False
    txtNumDoc.BackColor = Branco
End If

End Sub

Private Sub mskCodBarra_GotFocus()
mskCodBarra.SetFocus
mskCodBarra.SelStart = 0
mskCodBarra.SelLength = Len(mskCodBarra.text)
End Sub

Private Sub mskCodBarra_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    If Not ValidaCodBarra(mskCodBarra.text) Then
        MsgBox "Código de Barra inválido.", vbExclamation, "Atenção"
        mskCodBarra.SetFocus
        Exit Sub
    Else
        LoadMatrix
        txtNumDoc.text = aCodBarra(0).NumDocumento
'        Executa
'        txtNumDoc_KeyPress vbKeyReturn
    End If
End If

End Sub

Private Sub mskDataCred_GotFocus()
mskDataCred.SetFocus
mskDataCred.SelStart = 0
mskDataCred.SelLength = Len(mskDataCred.text)

End Sub

Private Sub mskDataPag_GotFocus()
mskDataPag.SetFocus
mskDataPag.SelStart = 0
mskDataPag.SelLength = Len(mskDataPag.text)

End Sub

Private Sub mskDataPag_LostFocus()
'cmdLoadDeb_Click
End Sub

Private Sub txtNumDoc_Change()

If txtNumDoc.text = "" Then
    mskCodBarra.Locked = False
    mskCodBarra.BackColor = Branco
End If
If Val(txtNumDoc.text) > 0 Then
    If bExec Then LimpaMascara mskCodBarra
    If mskCodBarra.ClipText = "" Then
        mskCodBarra.Locked = True
        mskCodBarra.BackColor = Kde
    End If
Else
    mskCodBarra.Locked = False
    mskCodBarra.BackColor = Branco
End If

End Sub

Private Sub txtNumDoc_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   txtNumDoc_LostFocus
Else
   Tweak txtNumDoc, KeyAscii, IntegerPositive
End If
End Sub

Private Sub txtNumDoc_LostFocus()
mskDataPag.SetFocus
'Executa
End Sub

Private Sub txtValorPago_Change()
On Error Resume Next
If txtValorPago.text = "" Then
   lblValorDif.Caption = FormatNumber(CDbl(lblValLanc.Caption), 2)
Else
   lblValorDif.Caption = FormatNumber(CDbl(txtValorPago.text) - CDbl(lblValLanc.Caption), 2)
End If
End Sub

Private Sub txtValorPago_KeyPress(KeyAscii As Integer)
Tweak txtValorPago, KeyAscii, DecimalPositive
End Sub

Private Function CalculaJuros2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

If dDataPagto = "00:00:00" Then
    dDataPagto = Now
End If

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= dDataPagto Then
    CalculaJuros2 = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(dDataPagto) And Year(dDataVencto) = Year(dDataPagto) Then
    CalculaJuros2 = 0
    Exit Function
End If

If Not dcJuros.Exists(Year(dDataPagto)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJuros2 = 0
   Exit Function
End If

'MONTA O NOVO VENCIMENTO A PARTIR DO DIA 1 DO MES SUBSEQUENTE
nDia = Day(dDataVencto)
nMes = Month(dDataVencto)
nAno = Year(dDataVencto)
nDia = 1
If nMes = 12 Then
    nMes = 1
    nAno = nAno + 1
Else
    nMes = nMes + 1
End If

sDataVencto = Format(nDia, "00") & "/" & Format(nMes, "00") & "/" & Format(nAno, "0000")
dDataVencto = Format(sDataVencto, "dd/mm/yyyy")
nNumMes = Int(DateDiff("d", dDataVencto, dDataPagto) / 30) + 1



'If dDataVencto >= dDataPagto Then
'    CalculaJuros2 = 0
'    Exit Function
'End If
'nNumMes = Int((DateDiff("d", dDataVencto, dDataPagto)) / 30)
Sql = "SELECT PERCJUROS FROM JUROS WHERE ANOJUROS=" & Year(dDataPagto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
        CalculaJuros2 = 0
        Exit Function
    Else
        nValorPerc = !PERCJUROS
    End If
   .Close
End With
nValorPerc = nValorPerc / 100

CalculaJuros2 = nValorDebito * nValorPerc * nNumMes
If CalculaJuros2 > 0 Then
   CalculaJuros2 = FormatNumber(CalculaJuros2, 3)
End If

End Function

Private Function CalculaMulta2(nValorDebito As Double, dDataVencto As Date, dDataPagto As Date) As Double
Dim nNumDia As Integer
Dim nValorPerc As Double
Dim RdoAux As rdoResultset, Sql As String

'If dDataVencto <= Now Then
'    CalculaMulta2 = 0
'    Exit Function
'End If


If dDataVencto >= dDataPagto Then
    CalculaMulta2 = 0
    Exit Function
End If

nNumDia = Abs(DateDiff("d", dDataPagto, dDataVencto))

If nNumDia = 0 Then
   CalculaMulta2 = 0
   Exit Function
End If

Sql = "SELECT MINDIA,MAXDIA,PERCDIA FROM MULTA WHERE ANOMULTA=" & Year(dDataVencto)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         If nNumDia >= !MINDIA And nNumDia <= !MAXDIA Then
             nValorPerc = !PERCDIA
             Exit Do
         ElseIf nNumDia >= !MINDIA And !MAXDIA = 0 Then
             nValorPerc = !PERCDIA
             Exit Do
         End If
        .MoveNext
    Loop
End With

nValorPerc = nValorPerc / 100
CalculaMulta2 = nValorDebito * nValorPerc
If CalculaMulta2 > 0 Then
   CalculaMulta2 = FormatNumber(CalculaMulta2, 3)
End If

End Function

Private Function CalculaCorrecao2(nValorDebito As Double, dDataBase As Date, dDataVencto As Date) As Double

Dim UfirAtual As Double
Dim UfirBase As Double

UfirAtual = RetornaUFIR(Year(dDataVencto))
UfirBase = RetornaUFIR(Year(dDataBase))

If UfirAtual = 0 Or UfirBase = 0 Then
    CalculaCorrecao2 = 0
    Exit Function
End If

CalculaCorrecao2 = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecao2 > 0 Then
   CalculaCorrecao2 = FormatNumber(CalculaCorrecao2, 2)
End If
End Function

