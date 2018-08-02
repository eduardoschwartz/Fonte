VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSimulaSimples 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simulador de Cálculo - Simples Nacional"
   ClientHeight    =   6450
   ClientLeft      =   3675
   ClientTop       =   3285
   ClientWidth     =   8265
   FillColor       =   &H00404040&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   8265
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   5805
      Left            =   30
      TabIndex        =   4
      Top             =   570
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   10239
      _Version        =   393217
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSimulaSimples.frx":0000
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1275
      MaxLength       =   12
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin prjChameleon.chameleonButton cmdGerar 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   6870
      TabIndex        =   3
      ToolTipText     =   "Cancelar Edição"
      Top             =   150
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Simular"
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
      MICON           =   "frmSimulaSimples.frx":0082
      PICN            =   "frmSimulaSimples.frx":009E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataVencto 
      Height          =   285
      Left            =   3450
      TabIndex        =   1
      Top             =   180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmSimulaSimples.frx":013D
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
   Begin esMaskEdit.esMaskedEdit mskDataPagto 
      Height          =   285
      Left            =   5610
      TabIndex        =   2
      Top             =   180
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmSimulaSimples.frx":0159
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
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Pagto...:"
      Height          =   225
      Index           =   1
      Left            =   4710
      TabIndex        =   7
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Vencto...:"
      Height          =   225
      Index           =   0
      Left            =   2490
      TabIndex        =   6
      Top             =   210
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor tributo..:"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Width           =   1035
   End
End
Attribute VB_Name = "frmSimulaSimples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Selic
    nAno As Integer
    nMes As Integer
    nValor As Double
End Type

Private Sub cmdGerar_Click()
Rtb.Text = ""

If Val(txtValor.Text) = 0 Then
    MsgBox "Digite o valor.", vbCritical, "Erro"
    Exit Sub
End If

If Not IsDate(mskDataVencto.Text) Then
    MsgBox "Digite a data de vencimento.", vbCritical, "Erro"
    Exit Sub
End If

If Not IsDate(mskDataPagto.Text) Then
    MsgBox "Digite a data de pagamento.", vbCritical, "Erro"
    Exit Sub
End If

If CDate(mskDataVencto.Text) > CDate(mskDataPagto.Text) Then
    MsgBox "Data de vencimento maior que data de pagamento.", vbCritical, "Erro"
    Exit Sub
End If

Simulador

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub mskDataPagto_Change()
Rtb.Text = ""
End Sub

Private Sub mskDataPagto_GotFocus()
mskDataPagto.SelStart = 0
mskDataPagto.SelLength = 10
mskDataPagto.SetFocus
End Sub

Private Sub mskDataPagto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGerar_Click
End If
End Sub

Private Sub mskDataVencto_Change()
Rtb.Text = ""
End Sub

Private Sub mskDataVencto_GotFocus()
mskDataVencto.SelStart = 0
mskDataVencto.SelLength = 10
mskDataVencto.SetFocus
End Sub

Private Sub mskDataVencto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGerar_Click
End If
End Sub

Private Sub txtValor_Change()
Rtb.Text = ""
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGerar_Click
Else
    Tweak txtValor, KeyAscii, DecimalPositive
End If
End Sub

Private Sub Simulador()
Dim nSemana As Integer, sSemana As String, dDataVencto As Date, dDataPagto As Date
Dim nPercMulta As Double, nValorMulta As Double, bJuros As Boolean, nValorJuros As Double, nValorTotal As Double
Dim aSelic() As Selic, aSelic2() As Selic, Sql As String, RdoAux As rdoResultset, nPercSelic As Double, x As Integer
Dim sDataSelic As String

ReDim aSelic(0): ReDim aSelic2(0)
Sql = "SELECT * FROM TAXASELIC ORDER BY ANO,MES"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aSelic(UBound(aSelic) + 1)
        aSelic(UBound(aSelic)).nAno = !Ano
        aSelic(UBound(aSelic)).nMes = !Mes
        aSelic(UBound(aSelic)).nValor = !Valor
       .MoveNext
    Loop
   .Close
End With

dDataVencto = CDate(mskDataVencto.Text)
dDataPagto = CDate(mskDataPagto.Text)
nSemana = Weekday(mskDataVencto.Text)

If nSemana < 6 Then 'de domingo a 5ª aumenta 1 dia
    dDataVencto = DateAdd("d", 1, dDataVencto)
ElseIf nSemana = 6 Then 'sexta aumenta 3 dias
    dDataVencto = DateAdd("d", 3, dDataVencto)
ElseIf nSemana = 7 Then 'sabado aumenta 2 dias
    dDataVencto = DateAdd("d", 2, dDataVencto)
End If

Select Case nSemana
    Case 1
        sSemana = "(Domingo)"
    Case 2
        sSemana = "(2ª feira)"
    Case 3
        sSemana = "(3ª feira)"
    Case 4
        sSemana = "(4ª feira)"
    Case 5
        sSemana = "(5ª feira)"
    Case 6
        sSemana = "(6ª feira)"
    Case 7
        sSemana = "(Sábado)"
End Select

Rtb.SelBold = True
Rtb.SelColor = vbBlack
Rtb.SelUnderline = True
Rtb.SelText = "Valor do tributo:"
Rtb.SelUnderline = False
Rtb.SelText = " R$ "
Rtb.SelColor = vbBlue
Rtb.SelText = FormatNumber(txtValor.Text, 2) & "  "
Rtb.SelColor = vbBlack
Rtb.SelUnderline = True
Rtb.SelText = "Data Vencto:"
Rtb.SelUnderline = False
Rtb.SelColor = vbBlue
Rtb.SelText = " " & mskDataVencto.Text & " "
Rtb.SelColor = &H8000&
Rtb.SelText = sSemana & " "
Rtb.SelColor = vbBlack
Rtb.SelUnderline = True
Rtb.SelText = "Data Pagto:"
Rtb.SelUnderline = False
Rtb.SelColor = vbBlue
Rtb.SelText = " " & mskDataPagto.Text & vbCrLf
Rtb.SelBold = True
Rtb.SelColor = vbBlack
Rtb.SelUnderline = True
Rtb.SelText = "Nº de Dias:"
Rtb.SelUnderline = False
Rtb.SelColor = vbBlue
nSemana = DateDiff("d", CDate(dDataVencto), CDate(dDataPagto)) + 1
Rtb.SelText = " " & CStr(nSemana) & vbCrLf & vbCrLf
Rtb.SelBold = True
Rtb.SelColor = vbRed
Rtb.SelUnderline = True
Rtb.SelText = "CÁLCULO DE MULTA:" & vbCrLf
Rtb.SelUnderline = False
Rtb.SelColor = vbBlack
Rtb.SelBold = True
Rtb.SelColor = vbBlack
Rtb.SelUnderline = False
nPercMulta = nSemana * 0.33
Rtb.SelText = CStr(nSemana) & " dia(s) de multa = " & CStr(nSemana) & " * 0,33% = " & FormatNumber(nPercMulta, 2) & "%" & vbCrLf
Rtb.SelUnderline = False
If nPercMulta > 20 Then
    Rtb.SelColor = &H80FF&
    Rtb.SelText = "% de Multa=" & FormatNumber(nPercMulta, 2) & "% -> Bloqueado em 20%" & vbCrLf
    nPercMulta = 20
End If
Rtb.SelBold = True
Rtb.SelColor = vbBlue
Rtb.SelUnderline = False
Rtb.SelText = "Valor da multa = "
Rtb.SelColor = vbBlack
Rtb.SelText = "(" & FormatNumber(txtValor.Text, 2) & " * " & FormatNumber(nPercMulta, 2) & "%" & ")" & " --> "
Rtb.SelBold = True
Rtb.SelColor = &H8000&
nValorMulta = CDbl(txtValor.Text) * (nPercMulta / 100)
Rtb.SelUnderline = False
Rtb.SelText = "Valor da multa = R$ " & FormatNumber(nValorMulta, 2) & vbCrLf & vbCrLf
Rtb.SelBold = True
Rtb.SelColor = vbRed
Rtb.SelUnderline = True
Rtb.SelText = "CÁLCULO DE JUROS:"
Rtb.SelUnderline = False
bJuros = True
If Year(dDataVencto) = Year(dDataPagto) And Month(dDataVencto) = Month(dDataPagto) Then bJuros = False
If Not bJuros Then
    Rtb.SelColor = vbBlack
    Rtb.SelText = " PAGAMENTO NO MESMO MÊS DO VENCIMENTO" & vbCrLf
    Rtb.SelBold = True
    Rtb.SelColor = &H8000&
    Rtb.SelUnderline = False
    Rtb.SelText = "Valor do juros = R$ 0,00" & vbCrLf & vbCrLf
    GoTo FINALIZA
End If

dDataVencto = CDate(mskDataVencto.Text)
nPercSelic = 1
For x = 1 To UBound(aSelic)
    If aSelic(x).nAno < Year(dDataVencto) Then
        GoTo proximo
    End If
    If aSelic(x).nAno = Year(dDataVencto) And aSelic(x).nMes <= Month(dDataVencto) Then
        GoTo proximo
    End If
    If aSelic(x).nAno > Year(dDataPagto) Then
        GoTo proximo
    End If
    If aSelic(x).nAno = Year(dDataPagto) And aSelic(x).nMes >= Month(dDataPagto) Then
        GoTo proximo
    End If
    
    ReDim Preserve aSelic2(UBound(aSelic2) + 1)
    aSelic2(UBound(aSelic2)).nAno = aSelic(x).nAno
    aSelic2(UBound(aSelic2)).nMes = aSelic(x).nMes
    aSelic2(UBound(aSelic2)).nValor = aSelic(x).nValor
proximo:
Next

nPercSelic = 1
If UBound(aSelic2) > 0 Then
    sDataSelic = ""
    For x = 1 To UBound(aSelic2)
        sDataSelic = sDataSelic & Format(aSelic2(x).nMes, "00") & "/" & Right(aSelic2(x).nAno, 2) & " + "
    Next
    sDataSelic = Left(sDataSelic, Len(sDataSelic) - 3)
    Rtb.SelText = vbCrLf
    Rtb.SelColor = vbBlack
    Rtb.SelText = "Taxa Selic de (" & sDataSelic & ") + 1%"
    sDataSelic = " = ("
    For x = 1 To UBound(aSelic2)
        sDataSelic = sDataSelic & FormatNumber(aSelic2(x).nValor, 2) & "% + "
        nPercSelic = nPercSelic + aSelic2(x).nValor
    Next
    sDataSelic = Left(sDataSelic, Len(sDataSelic) - 3)
    Rtb.SelText = sDataSelic & ") + 1% = "
    Rtb.SelText = FormatNumber(nPercSelic - 1) & "% + 1% = " & FormatNumber(nPercSelic) & "%"
Else
    Rtb.SelText = vbCrLf
    Rtb.SelColor = vbBlack
    Rtb.SelText = "APENAS JUROS DO MES DE 1%"
End If
Rtb.SelText = vbCrLf
Rtb.SelBold = True
Rtb.SelColor = vbBlue
Rtb.SelUnderline = False
Rtb.SelText = "Valor do juros = "
Rtb.SelColor = vbBlack
Rtb.SelText = "(" & FormatNumber(txtValor.Text, 2) & " * " & FormatNumber(nPercSelic, 2) & "%" & ")" & " --> "
Rtb.SelBold = True
Rtb.SelColor = &H8000&
nValorJuros = CDbl(txtValor.Text) * (nPercSelic / 100)
Rtb.SelUnderline = False
Rtb.SelText = "Valor do juros = R$ " & FormatNumber(nValorJuros, 2) & vbCrLf & vbCrLf
Rtb.SelBold = True

FINALIZA:
Rtb.SelBold = True
Rtb.SelColor = vbRed
Rtb.SelUnderline = True
Rtb.SelText = "VALORFINAL:" & vbCrLf
Rtb.SelBold = True
Rtb.SelUnderline = False
Rtb.SelColor = vbBlue
Rtb.SelText = "Valor Final = "
Rtb.SelColor = vbBlack
Rtb.SelText = "Valor do Tributo + Valor da Multa + Valor do Juros" & vbCrLf
Rtb.SelBold = True
Rtb.SelColor = vbBlue
Rtb.SelText = "Valor Final = "
Rtb.SelColor = vbBlack
Rtb.SelText = "R$ " & FormatNumber(txtValor.Text, 2) & " + " & FormatNumber(nValorMulta, 2) & " + " & FormatNumber(nValorJuros, 2) & " --> "
Rtb.SelBold = True
Rtb.SelColor = &H8000&
Rtb.SelUnderline = False
Rtb.SelText = "Valor Final = R$ " & FormatNumber(CDbl(txtValor.Text) + nValorJuros + nValorMulta, 2)


End Sub

Private Sub Fonte(Alinhamento As RichTextLib.SelAlignmentConstants, Cor As Long, Size As Integer, Negrito As Boolean, Italico As Boolean, Sublinhado As Boolean)

With Rtb
    .SelAlignment = Alinhamento
    .SelColor = Cor
    .SelFontSize = Size
    .SelBold = Negrito
    .SelUnderline = Sublinhado
    .SelItalic = Italico
End With

End Sub

Private Sub Negrito()
Rtb.SelBold = True
End Sub

Private Sub Normal()
Rtb.SelBold = False
End Sub

