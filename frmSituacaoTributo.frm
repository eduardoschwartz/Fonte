VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSituacaoTributo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situação dos tributos lançados"
   ClientHeight    =   2025
   ClientLeft      =   2025
   ClientTop       =   4455
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   5940
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4590
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1035
      Width           =   1185
   End
   Begin VB.ComboBox cmbSituacao 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1035
      Width           =   2490
   End
   Begin VB.ComboBox cmbTributo 
      Height          =   315
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   4740
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1215
      TabIndex        =   1
      Top             =   585
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmSituacaoTributo.frx":0000
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
      Left            =   3555
      TabIndex        =   2
      Top             =   600
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmSituacaoTributo.frx":001C
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
      Left            =   4680
      TabIndex        =   5
      ToolTipText     =   "Imprimir esta Tela"
      Top             =   1620
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
      MICON           =   "frmSituacaoTributo.frx":0038
      PICN            =   "frmSituacaoTributo.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   1620
      Width           =   3795
      _ExtentX        =   6694
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
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Código..:"
      Height          =   195
      Index           =   1
      Left            =   3735
      TabIndex        =   11
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Situação..:"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   10
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2595
      TabIndex        =   8
      Top             =   645
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Tributo....:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   780
   End
End
Attribute VB_Name = "frmSituacaoTributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

If cmbSituacao.ListIndex = -1 Then
    MsgBox "Selecione a situação", vbExclamation, "atenção"
    Exit Sub
End If

GeraArquivo

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaTributo
End Sub

Private Sub CarregaTributo()
Dim RdoAux As rdoResultset, Sql As String

Sql = "select codtributo,desctributo from tributo order by desctributo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTributo.AddItem !desctributo
        cmbTributo.ItemData(cmbTributo.NewIndex) = !CodTributo
        
       .MoveNext
    Loop
   .Close
End With

Sql = "select codsituacao,descsituacao from situacaolancamento order by descsituacao"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbSituacao.AddItem !DescSituacao
        cmbSituacao.ItemData(cmbSituacao.NewIndex) = !Codsituacao
        
       .MoveNext
    Loop
   .Close
End With


cmbTributo.ListIndex = 0
cmbSituacao.ListIndex = -1
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.Value = (nVal * 100) / nTot
Else
   PBar.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

Private Sub GeraArquivo()
Dim RdoAux As rdoResultset, Sql As String, nCodTributo As Integer, nPos As Long, nTot As Long
Dim nValorPago As Double, nCodReduz As Long, sNome As String, RdoAux2 As rdoResultset, nStatus As Integer

On Error GoTo Erro

cmdPrint.Enabled = False
Sql = "delete from relsituacaotributo where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

nCodTributo = cmbTributo.ItemData(cmbTributo.ListIndex)
nStatus = cmbSituacao.ItemData(cmbSituacao.ListIndex)

Sql = "SELECT debitotributo.codreduzido, debitopago.datarecebimento, debitopago.valorpagoreal, debitotributo.valortributo,debitoparcela.statuslanc, "
Sql = Sql & "debitoparcela.DataVencimento FROM debitotributo INNER JOIN debitoparcela ON debitotributo.codreduzido = debitoparcela.codreduzido AND debitotributo.anoexercicio = debitoparcela.anoexercicio AND "
Sql = Sql & "debitotributo.codlancamento = debitoparcela.codlancamento AND debitotributo.seqlancamento = debitoparcela.seqlancamento AND "
Sql = Sql & "debitotributo.numparcela = debitoparcela.numparcela AND debitotributo.codcomplemento = debitoparcela.codcomplemento LEFT OUTER JOIN "
Sql = Sql & "debitopago ON debitotributo.codreduzido = debitopago.codreduzido AND debitotributo.anoexercicio = debitopago.anoexercicio AND "
Sql = Sql & "debitotributo.codlancamento = debitopago.codlancamento AND debitotributo.seqlancamento = debitopago.seqlancamento AND "
Sql = Sql & "debitotributo.NumParcela = debitopago.NumParcela And debitotributo.CODCOMPLEMENTO = debitopago.CODCOMPLEMENTO "
Sql = Sql & "WHERE ((debitopago.datarecebimento between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "') or debitopago.datarecebimento is null) "
Sql = Sql & "AND debitotributo.codtributo = " & nCodTributo & " AND debitoparcela.statuslanc =" & nStatus
If Val(txtCod.Text) > 0 Then
    Sql = Sql & " and debitotributo.codreduzido=" & Val(txtCod.Text)
End If
Sql = Sql & " order by debitotributo.codreduzido,datavencimento"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    nPos = 1
    Ocupado
    Me.Refresh
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTot
        End If
        
        If Not IsNull(!datarecebimento) Then
            nValorPago = !ValorTributo
        Else
            nValorPago = 0
        End If
        
        nCodReduz = !CODREDUZIDO
        
        If nCodReduz < 100000 Then
            Sql = "select nomecidadao as nome from vwfullimovel where codreduzido=" & nCodReduz
        ElseIf nCodReduz >= 100000 And nCodReduz <= 300000 Then
            Sql = "select razaosocial as nome from mobiliario where codmobiliario=" & nCodReduz
        Else
            Sql = "select nomecidadao as nome from cidadao where codcidadao=" & nCodReduz
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = SubNull(RdoAux2!nome)
        RdoAux2.Close
        
        Sql = "insert relsituacaotributo(usuario,seq,codreduzido,nome,codtributo,desctributo,sit,datavencto,datarecebimento,valorlanc,valorpago) values('"
        Sql = Sql & NomeDeLogin & "'," & nPos & "," & nCodReduz & ",'" & Mask(sNome) & "'," & nCodTributo & ",'" & Mask(cmbTributo.Text) & "',"
        Sql = Sql & !statuslanc & ",'" & Format(!DataVencimento, "mm/dd/yyyy") & "','" & Format(!datarecebimento, "mm/dd/yyyy") & "',"
        Sql = Sql & Virg2Ponto(Format(!ValorTributo, "#0.00")) & "," & Virg2Ponto(Format(nValorPago, "#0.00")) & ")"
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

PBar.Value = 0: PBar.Color = vbWhite
Liberado
Me.Refresh
cmdPrint.Enabled = True

frmReport.ShowReport2 "SITUACAOTRIBUTO", frmMdi.hwnd, Me.hwnd

Sql = "delete from relsituacaotributo where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Exit Sub

Erro:
cmdPrint.Enabled = True
MsgBox Err.Description
Resume Next
End Sub

Private Sub mskDataFim_GotFocus()
mskDataFim.SetFocus
End Sub

Private Sub mskDataIni_GotFocus()
mskDataIni.SetFocus
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub
