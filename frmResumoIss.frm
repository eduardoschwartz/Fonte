VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmResumoIss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo das Guias de Iss emitidas por período pela Giss"
   ClientHeight    =   8190
   ClientLeft      =   7980
   ClientTop       =   3960
   ClientWidth     =   10515
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   10515
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   60
      Left            =   90
      TabIndex        =   16
      Top             =   5010
      Width           =   10335
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   3825
      Left            =   60
      TabIndex        =   3
      Top             =   510
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   6747
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº Guia"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Vencto."
         Object.Width           =   2187
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "St."
         Object.Width           =   776
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Código"
         Object.Width           =   1692
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Razão Social"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CPF/CNPJ"
         Object.Width           =   2823
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   345
      Left            =   4770
      TabIndex        =   2
      ToolTipText     =   "Consultar relatório selecionado"
      Top             =   90
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmResumoIss.frx":0000
      PICN            =   "frmResumoIss.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   120
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmResumoIss.frx":0176
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
      Top             =   120
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmResumoIss.frx":0192
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
   Begin MSComctlLib.ListView lvEmpresa 
      Height          =   2415
      Left            =   60
      TabIndex        =   18
      Top             =   5370
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   4260
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1692
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "CPF/CNPJ"
         Object.Width           =   2823
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Pb 
      Height          =   165
      Left            =   7590
      TabIndex        =   21
      Top             =   7890
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
      Left            =   9750
      TabIndex        =   22
      Top             =   7890
      Width           =   480
   End
   Begin VB.Label lblQtdeF 
      Alignment       =   1  'Right Justify
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1260
      TabIndex        =   20
      Top             =   7890
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Quantidade..:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   19
      Top             =   7890
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Prestadores de serviço sem faturamento no período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   180
      TabIndex        =   17
      Top             =   5100
      Width           =   4275
   End
   Begin VB.Label Label1 
      Caption         =   "Guias total............:"
      Height          =   195
      Index           =   4
      Left            =   3420
      TabIndex        =   15
      Top             =   4650
      Width           =   1485
   End
   Begin VB.Label lblTotal 
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4920
      TabIndex        =   14
      Top             =   4650
      Width           =   1965
   End
   Begin VB.Label lblOutro 
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1350
      TabIndex        =   13
      Top             =   4650
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Guias outros..:"
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   12
      Top             =   4650
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Guias canceladas..:"
      Height          =   195
      Index           =   2
      Left            =   6960
      TabIndex        =   11
      Top             =   4410
      Width           =   1485
   End
   Begin VB.Label lblCancel 
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   8460
      TabIndex        =   10
      Top             =   4410
      Width           =   1965
   End
   Begin VB.Label lblNPago 
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   4920
      TabIndex        =   9
      Top             =   4410
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "Guias não pagas..:"
      Height          =   195
      Index           =   0
      Left            =   3420
      TabIndex        =   8
      Top             =   4410
      Width           =   1485
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2565
      TabIndex        =   7
      Top             =   165
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   330
      TabIndex        =   6
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Guias pagas..:"
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   5
      Top             =   4410
      Width           =   1125
   End
   Begin VB.Label lblPago 
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   1350
      TabIndex        =   4
      Top             =   4410
      Width           =   1965
   End
End
Attribute VB_Name = "frmResumoIss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()
Dim Sql As String, RdoAux As rdoResultset, itmX As ListItem, sDoc As String, nQtdePago As Integer, nValorPago As Double, nQtdeNPago As Integer, nValorNPago As Double
Dim nQtdeCancel As Integer, nValorCancel As Double, nQtdeOutro As Integer, nValorOutro As Double, nQtdeTotal As Integer, nValorTotal As Double, RdoaAux2 As rdoResultset
Dim aSuspensoCod() As Long, lResult As Long, nQtdeF As Integer, aEmpresa() As Long, x As Long, bFind As Boolean, nPos As Long, nTot As Long

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

ReDim aSuspensoCod(0)
Sql = "SELECT codmobiliario, DataEv, codtipoevento From vwMOBILIARIOSUSPENSO Where (codtipoevento = 2)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aSuspensoCod(UBound(aSuspensoCod) + 1)
        aSuspensoCod(UBound(aSuspensoCod)) = !codmobiliario
       .MoveNext
    Loop
   .Close
End With

ReDim aEmpresa(0)
Pb.value = 0: lblPB.Caption = "0%"
lvMain.ListItems.Clear: lvEmpresa.ListItems.Clear
nQtdePago = 0: nValorPago = 0: nQtdeNPago = 0: nValorNPago = 0: nQtdeCancel = 0: nValorCancel = 0: nQtdeOutro = 0: nValorOutro = 0: nQtdeTotal = 0: nValorTotal = 0: nQtdeF = 0
lblPago.Caption = "": lblNPago.Caption = "": lblCancel.Caption = "": lblOutro.Caption = "": lblTotal.Caption = "": lblQtdeF.Caption = ""

DoEvents

Sql = "SELECT debitoparcela.codreduzido, debitoparcela.datavencimento, debitoparcela.datadebase, debitotributo.valortributo, debitoparcela.statuslanc,parceladocumento.NumDocumento , mobiliario.RazaoSocial,"
Sql = Sql & "mobiliario.Cnpj, mobiliario.CPF FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND "
Sql = Sql & "debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND debitoparcela.numparcela = parceladocumento.numparcela AND "
Sql = Sql & "debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN mobiliario ON debitoparcela.codreduzido = mobiliario.codigomob WHERE (debitoparcela.codreduzido BETWEEN 100000 AND 300000) AND "
Sql = Sql & "(debitoparcela.datavencimento BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "') AND (debitoparcela.codlancamento = 5) AND (debitoparcela.usuario = 'Giss Online') AND "
Sql = Sql & "(parceladocumento.numdocumento BETWEEN 2000000 AND 3000000) ORDER BY parceladocumento.numdocumento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "A consulta não gerou resultado.", vbExclamation, "Atenção"
        Exit Sub
    End If
    Ocupado
    DoEvents
    Do Until .EOF
        sDoc = Trim(SubNull(!Cnpj))
        If sDoc = "" Then
            sDoc = Trim(SubNull(!CPF))
        End If
        
        bFind = False
        For x = 1 To UBound(aEmpresa)
            If aEmpresa(x) = !CODREDUZIDO Then
                bFind = True
                Exit For
            End If
        Next
        If Not bFind Then
            ReDim Preserve aEmpresa(UBound(aEmpresa) + 1)
            aEmpresa(UBound(aEmpresa)) = !CODREDUZIDO
        End If
        
        Set itmX = lvMain.ListItems.Add(, , Format(!NumDocumento, "0000000"))
        itmX.SubItems(1) = Format(!DataVencimento, "dd/mm/yyyy")
        itmX.SubItems(2) = Format(!statuslanc, "00")
        itmX.SubItems(3) = !CODREDUZIDO
        itmX.SubItems(4) = SubNull(!RazaoSocial)
        itmX.SubItems(5) = sDoc
        nQtdeTotal = nQtdeTotal + 1
        nValorTotal = nValorTotal + !ValorTributo
        
        If !statuslanc = 2 Then
            nQtdePago = nQtdePago + 1
            nValorPago = nValorPago + !ValorTributo
        ElseIf !statuslanc = 3 Then
            nQtdeNPago = nQtdeNPago + 1
            nValorNPago = nValorNPago + !ValorTributo
        ElseIf !statuslanc = 5 Then
            nQtdeCancel = nQtdeCancel + 1
            nValorCancel = nValorCancel + !ValorTributo
        Else
            nQtdeOutro = nQtdeOutro + 1
            nValorOutro = nValorOutro + !ValorTributo
        End If
        
       .MoveNext
    Loop
   .Close
End With


Sql = "SELECT codigomob,razaosocial,cnpj FROM  mobiliario WHERE cnpj<>'' and cnpj is not null and (dataencerramento IS NULL) order by codigomob"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nPos = 1
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 20 = 0 Then
            CallPb nPos, nTot
        End If
    
        bFind = False
        For x = 1 To UBound(aSuspensoCod)
            If aSuspensoCod(x) = !codigomob Then
                bFind = True
                GoTo proximo
            End If
        Next
        
        
        bFind = False
        For x = 1 To UBound(aEmpresa)
            If aEmpresa(x) = !codigomob Then
                bFind = True
                GoTo proximo
            End If
        Next
        
        If SNCheck(!codigomob) Then GoTo proximo
        
        
        Set itmX = lvEmpresa.ListItems.Add(, "C" & Format(!codigomob, "000000"), Format(!codigomob, "000000"))
        itmX.SubItems(1) = SubNull(!RazaoSocial)
        itmX.SubItems(2) = SubNull(!Cnpj)
        nQtdeF = nQtdeF + 1
        
proximo:
    nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With


Liberado
lblPago.Caption = "R$ " & FormatNumber(nValorPago, 2) & " (" & nQtdePago & ")"
lblNPago.Caption = "R$ " & FormatNumber(nValorNPago, 2) & " (" & nQtdeNPago & ")"
lblCancel.Caption = "R$ " & FormatNumber(nValorCancel, 2) & " (" & nQtdeCancel & ")"
lblOutro.Caption = "R$ " & FormatNumber(nValorOutro, 2) & " (" & nQtdeOutro & ")"
lblTotal.Caption = "R$ " & FormatNumber(nValorTotal, 2) & " (" & nQtdeTotal & ")"
lblQtdeF.Caption = nQtdeF
Pb.value = 0
lblPB.Caption = "0%"

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub lvEmpresa_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvEmpresa.SortKey = ColumnHeader.Position - 1
lvEmpresa.Sorted = True
lvEmpresa.SortOrder = lvwAscending

End Sub

Private Sub lvMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvMain.SortKey = ColumnHeader.Position - 1
lvMain.Sorted = True
lvMain.SortOrder = lvwAscending
End Sub

Private Function SNCheck(nCodigo As Long) As Boolean
Dim RdoAux As rdoResultset, Sql As String
Sql = "SELECT " & NomeBaseDados & ".dbo.RETORNASN(" & Format(nCodigo, "000000") & ",'" & Format(Now, "mm/dd/yyyy") & "') AS RETORNO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
     If RdoAux!RETORNO = 1 Then
        SNCheck = True
     Else
        SNCheck = False
     End If
    .Close
End With

End Function

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

