VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmDAM 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emissão de D.A.M."
   ClientHeight    =   4845
   ClientLeft      =   3240
   ClientTop       =   2220
   ClientWidth     =   8970
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCobranca 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Cobrança"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3900
      TabIndex        =   24
      Top             =   4500
      Width           =   1860
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Boleto bancário"
      Height          =   195
      Index           =   1
      Left            =   6885
      TabIndex        =   26
      Top             =   3195
      Width           =   1545
   End
   Begin VB.OptionButton optTipo 
      Caption         =   "Normal"
      Height          =   195
      Index           =   0
      Left            =   5850
      TabIndex        =   25
      Top             =   3195
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.CheckBox chkCorrecao 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Isenção de Correção"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3900
      TabIndex        =   23
      Top             =   4170
      Width           =   1905
   End
   Begin prjChameleon.chameleonButton cmdAnistia 
      Height          =   240
      Left            =   3195
      TabIndex        =   22
      ToolTipText     =   "Vencimentos da Ansitia"
      Top             =   3240
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   423
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   14869218
      BCOLO           =   14869218
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmDAM.frx":0000
      PICN            =   "frmDAM.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskVencimento 
      Height          =   285
      Left            =   2070
      TabIndex        =   21
      Top             =   3195
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmDAM.frx":0176
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
   Begin VB.CheckBox chkAnistia 
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   3900
      TabIndex        =   19
      Top             =   3510
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CheckBox chkJulgamento 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Em Julgamento"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3900
      TabIndex        =   16
      Top             =   3840
      Width           =   1455
   End
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Left            =   2070
      TabIndex        =   15
      Top             =   4140
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmDAM.frx":0192
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
   Begin VB.CheckBox chkVenctoAtual 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Calcular com data de:"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   4170
      Width           =   1905
   End
   Begin VB.CheckBox chkMulta 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Isenção Total de Juros e Multa"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   3840
      Width           =   3555
   End
   Begin VB.CheckBox chkTx 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Remover Taxa de Expediente da DAM !!!"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   120
      TabIndex        =   11
      Top             =   3540
      Width           =   3555
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   7530
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   4410
      Width           =   1365
      _ExtentX        =   2408
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
      MICON           =   "frmDAM.frx":01AE
      PICN            =   "frmDAM.frx":01CA
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
      Left            =   6105
      TabIndex        =   3
      ToolTipText     =   "Emissão da DAM Informada"
      Top             =   4410
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Emitir DAM"
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
      MICON           =   "frmDAM.frx":0238
      PICN            =   "frmDAM.frx":0254
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   1230
      Left            =   270
      TabIndex        =   1
      Top             =   5310
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   2170
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDAM.frx":03AE
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2700
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   4763
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   128
      ForeColorSel    =   16777215
      GridColorFixed  =   16777215
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmDAM.frx":044A
   End
   Begin VB.Label lblSid 
      Caption         =   "0"
      Height          =   195
      Left            =   135
      TabIndex        =   27
      Top             =   4545
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Isenção dos juros e multa conforme REFIS-IV em :"
      Enabled         =   0   'False
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4200
      TabIndex        =   20
      Top             =   3510
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   8340
      TabIndex        =   18
      Top             =   3540
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAnistia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   7680
      TabIndex        =   17
      Top             =   3540
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblValorExp2 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3780
      TabIndex        =   12
      Top             =   3195
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Vencimento......:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da DAM..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5850
      TabIndex        =   9
      Top             =   2865
      Width           =   1935
   End
   Begin VB.Label lblValorTotal 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7845
      TabIndex        =   8
      Top             =   2865
      Width           =   1065
   End
   Begin VB.Label lblValorExp 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4725
      TabIndex        =   7
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label lblTotalLanc 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1995
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Taxa Expediênte...:"
      Height          =   195
      Index           =   3
      Left            =   3255
      TabIndex        =   5
      Top             =   2880
      Width           =   1470
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor dos Lançamentos..:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1800
   End
   Begin VB.Menu Anistia 
      Caption         =   "mnuAnistia"
      Visible         =   0   'False
      Begin VB.Menu mnuA1 
         Caption         =   "até 31/10/2012 (100%)"
      End
      Begin VB.Menu mnuA2 
         Caption         =   "até 30/11/2012 (95%)"
      End
      Begin VB.Menu mnuA3 
         Caption         =   "até 28/12/2012 (90%)"
      End
   End
End
Attribute VB_Name = "frmDAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sLANCAMENTO As String, sTributo As String, nSomaPrincipal As Double
Dim bCorrecao As Boolean, dVencto As Date, bHonorario As Boolean

Public Property Let Honorarios(bValor As Boolean)
    bHonorario = bValor
End Property

Private Sub chkAnistia_Click()
Dim x As Integer

With grdTemp
    For x = 1 To .Rows - 1
        If Val(.TextMatrix(x, 0)) <> 2012 Then
            If Year(CDate(.TextMatrix(x, 6))) > 2010 Then
                lblAnistia.Caption = "0,00"
                Exit Sub
            End If
        End If
    Next
End With


If chkAnistia.Value = vbUnchecked Then
    lblAnistia.Caption = "0,00"
    CarregaLista2
Else
    If dVencto <= CDate("31/10/2012") Then
       lblAnistia.Caption = "100,00"
    ElseIf dVencto <= CDate("30/11/2012") Then
       lblAnistia.Caption = "95,00"
    ElseIf dVencto <= CDate("28/12/2012") Then
       lblAnistia.Caption = "90,00"
    ElseIf dVencto >= CDate("28/12/2012") Then
       If chkMulta.Value = 0 Then
          lblAnistia.Caption = "0,00"
       Else
          lblAnistia.Caption = "100,00"
       End If
       Exit Sub
    End If
    nPerc = 100 - CDbl(lblAnistia.Caption)
    With grdTemp
        For x = 1 To grdTemp.Rows - 1
            .TextMatrix(x, 11) = FormatNumber(CDbl(.TextMatrix(x, 11)) * nPerc / 100, 2)
            .TextMatrix(x, 12) = FormatNumber(CDbl(.TextMatrix(x, 12)) * nPerc / 100, 2)
            .TextMatrix(x, 13) = FormatNumber(CDbl(.TextMatrix(x, 9)) + CDbl(.TextMatrix(x, 10)) + CDbl(.TextMatrix(x, 11)) + CDbl(.TextMatrix(x, 12)), 2)
        Next
    End With
End If
CalculaTotal
End Sub

Private Sub chkCorrecao_Click()
Dim x As Integer, bAchou As Boolean
CarregaLista2
End Sub

Private Sub chkJulgamento_Click()
CarregaLista2
End Sub

Private Sub chkMulta_Click()

Dim x As Integer, bAchou As Boolean
bAchou = False
If chkMulta.Value = 1 Then
    With grdTemp
        For x = 1 To .Rows - 2
            If .TextMatrix(x, 7) = "N" Then
                bAchou = True
                Exit For
            End If
        Next
    End With
    lblAnistia.Caption = "100,00"
Else
    If Not bAnistia Then
        lblAnistia.Caption = "0,00"
    End If
End If

CarregaLista2

End Sub

Private Sub chkTx_Click()

If chkTx.Value = 1 Then
    'remover
    grdTemp.TextMatrix(grdTemp.Rows - 1, 9) = "0,00"
    grdTemp.TextMatrix(grdTemp.Rows - 1, 13) = "0,00"
    nSomaPrincipal = nSomaPrincipal - CDbl(lblValorExp.Caption)
    lblValorExp.Caption = "0,00"
    
Else
    'adicionar
    grdTemp.TextMatrix(grdTemp.Rows - 1, 9) = lblValorExp2.Caption
    grdTemp.TextMatrix(grdTemp.Rows - 1, 13) = lblValorExp2.Caption
    lblValorExp.Caption = lblValorExp2.Caption
    nSomaPrincipal = nSomaPrincipal + CDbl(lblValorExp.Caption)
End If
CalculaTotal
End Sub

Private Sub chkVenctoAtual_Click()

If Not IsDate(mskVenc.Text) And chkVenctoAtual.Value = 1 Then
    MsgBox "Digite uma data válida.", vbExclamation, "Atenção"
    chkVenctoAtual.Value = 0
    Exit Sub
End If

CarregaLista2
End Sub

Private Sub cmdAnistia_Click()
PopupMenu Anistia
End Sub

Private Sub cmdBaixa_Click()
Dim nCodReduz As Long, nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer
Dim nCompl As Integer, x As Integer, aAno() As Integer, Y As Integer, bAchou As Boolean

If NomeDeLogin <> "SCHWARTZ" And optTipo(1).Value = True Then
    MsgBox "No momento você só pode emitir DAM normal.", vbCritical, "Acesso negado"
    Exit Sub
ElseIf NomeDeLogin = "SCHWARTZ" And optTipo(1).Value = True Then
    EmiteBoleto
    Exit Sub
End If

If bAnistia And chkAnistia.Value = vbChecked Then
    
    bAchou = False
    With grdTemp
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) <> 2012 Then
                bAchou = True
            End If
        Next
    End With
    If bAchou Then
        With grdTemp
            For x = 1 To .Rows - 1
                If Val(.TextMatrix(x, 0)) = 2012 And Val(.TextMatrix(x, 1)) <> 4 And Val(.TextMatrix(x, 1)) <> 41 And Val(.TextMatrix(x, 1)) <> 69 Then
                    MsgBox "Não é permitido emitir débitos de 2012 junto com outros anos.", vbExclamation, "Atenção"
                    lblAnistia.Caption = "0,00"
                    Exit Sub
                End If
            Next
        End With
    End If
End If


nCodReduz = Val(frmDebitoImob.txtCod.Text)
If (nCodReduz < 100000 Or nCodReduz > 300000) And chkJulgamento.Value = 1 Then
    MsgBox "Apuração Fiscal apenas para débitos mobiliários.", vbCritical, "Atenção"
    Liberado
    Exit Sub
End If

If Not IsDate(mskVencimento.Text) Then
    MsgBox "Data de Vencimento inválido.", vbCritical, "Atenção"
    Exit Sub
End If

'carrega os anos distintos
ReDim aAno(0)
For x = 1 To grdTemp.Rows - 1
    If Val(grdTemp.TextMatrix(x, 1)) <> 4 And Val(grdTemp.TextMatrix(x, 1)) <> 41 Then
        nAno = grdTemp.TextMatrix(x, 0)
        bAchou = False
        For Y = 1 To UBound(aAno)
            If aAno(Y) = nAno Then
                bAchou = True
                Exit For
            End If
        Next
        If Not bAchou Then
            ReDim Preserve aAno(UBound(aAno) + 1)
            aAno(UBound(aAno)) = nAno
        End If
    End If
Next
'verifica se 2007 esta junto com outros anos
bAchou = False
For x = 1 To UBound(aAno)
    If aAno(x) = 2007 Then
        bAchou = True
        Exit For
    End If
Next
If bAchou And UBound(aAno) > 1 Then
'    MsgBox "De acordo com a Lei Complementar No 83, de 22 de junho de 2.007, o exercício de 2.007 não pode ser emitido com outros exercícios." & vbCrLf & "favor emitir em guias separadas.", vbExclamation, "Atenção"
'    Exit Sub
End If



Ocupado
GravaDam

'MUDA O STATUS DAS PARCELAS SELECIONADAS PARA EM JULGAMENTO
If chkJulgamento.Value = 1 Then
    With frmDebitoImob.grdExtrato
        For x = 1 To .Rows
            If .CellText(x, 12) = "S" Then
                nAno = Val(.CellText(x, 1))
                nLanc = Val(Left$(.CellText(x, 2), 3))
                nSeq = Val(.CellText(x, 3))
                nParc = Val(.CellText(x, 4))
                nCompl = Val(.CellText(x, 5))
               .CellText(x, 12) = ""
               .CellText(x, 6) = "20 - EM JULGAMENTO"
               .CellForeColor(.Rows, 6) = &H80FF&
                Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=20"
                Sql = Sql & " WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
                Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
                Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
            End If
        Next
    End With
End If

Liberado
Unload Me
End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()

Centraliza Me
Me.Top = Me.Top + 1200
Ocupado
   
If Not bAnistia Then
    cmdAnistia.Visible = False
    chkAnistia.Value = vbUnchecked
End If
'If Val(frmDebitoImob.txtCod) = 523872 Then bCorrecao = False

mskVencimento.Text = frmDebitoImob.lblDataVencto.Caption
'Sql = "DELETE FROM DAM WHERE usuario='" & NomeDeLogin & "'"
'cn.Execute Sql, rdExecDirect

CarregaLista2

Liberado
Select Case Mid(frmMdi.Sbar.Panels(2).Text, 10, Len(frmMdi.Sbar.Panels(2).Text) - 8)
    Case "SCHWARTZ", "SIMONE", "RENATA", "GLEISE", "ROSE", "EDUARDO", "RITA", "LUIZH"
        chkVenctoAtual.Enabled = True
        mskVenc.Enabled = True
    Case Else
        chkVenctoAtual.Enabled = False
        mskVenc.Enabled = False
End Select

'chkJulgamento.Enabled = frmDebitoImob.mnuReativarJ.Enabled

'If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "ROSE" Or NomeDeLogin = "LETICIA" Or NomeDeLogin = "CARMELINO" Then
'    chkAnistia.Enabled = True
'Else
'    chkAnistia.Enabled = False
'End If



'If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "SIMONE" Or NomeDeLogin = "EDUARDO" Or NomeDeLogin = "RITA" Or NomeDeLogin = "JOSEANE" Or NomeDeLogin = "GLEISE" Then
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "SIMONE" Or NomeDeLogin = "EDUARDO" Or NomeDeLogin = "RITA" Or NomeDeLogin = "IORIO" Or NomeDeLogin = "LUIZH" Or _
    NomeDeLogin = "GLEISE" Or NomeDeLogin = "LIAMAR" Or NomeDeLogin = "HELOISA" Or NomeDeLogin = "SANDRA" Or NomeDeLogin = "CINTIA" Or NomeDeLogin = "LEILA" Or _
    NomeDeLogin = "REGIANE" Or NomeDeLogin = "SOLANGE" Or NomeDeLogin = "ELAINE" Or NomeDeLogin = "ALESSANDRA" Then
    chkMulta.Enabled = True
    chkCorrecao.Enabled = True
Else
    chkMulta.Enabled = False
    chkCorrecao.Enabled = False
End If

End Sub

Private Sub CarregaLista2()
Dim x As Integer, Y As Integer, nCodReduz As Long, aLanc() As String, Achou As Boolean
Dim sAno As String, sLanc As String, sSeq As String, sParc As String, aAno() As Integer
Dim sComp As String, sSit As String, sVencto As String, sDA As String
Dim sAj As String, nValorPrincipal As Double, sDataBase As String, bDA As Boolean
Dim nValorCorrecao As Double, nValorJuros As Double, nValorMulta As Double, nValorTotal As Double, bITBI As Boolean, nPerc As Double, nValorHon As Double
Dim nSomaTotal As Double, nSomaHon As Double, bJuros As Boolean, bMulta As Boolean, nCodTrib As Integer, qd As New rdoQuery, nSID As Long


Dim Sql As String, RdoAux As rdoResultset
On Error Resume Next
ReDim aLanc(0)



If chkVenctoAtual.Value = 0 Then
    If mskVencimento.ClipText = "" Then mskVencimento.Text = Right(frmMdi.Sbar.Panels(6).Text, 10)
    dVencto = Format(mskVencimento.Text, "dd/mm/yyyy")
Else
    dVencto = Format(mskVenc.Text, "dd/mm/yyyy")
End If


grdTemp.Rows = 1
nSomaHon = 0
nSomaTotal = 0
nSomaPrincipal = 0
With frmDebitoImob.grdExtrato
    nCodReduz = Val(frmDebitoImob.txtCod.Text)
    For x = 1 To .Rows
        bJuros = False: bMulta = False
        If .CellText(x, 12) = "S" Then
           sAno = .CellText(x, 1)
           sLanc = Left$(.CellText(x, 2), 3)
           sLANCAMENTO = Right$(.CellText(x, 2), Len(.CellText(x, 2)) - 5)
           Achou = False
           For Y = 1 To UBound(aLanc)
               If aLanc(Y) = sLANCAMENTO Then
                  Achou = True
                  Exit For
               End If
           Next
           If Not Achou Then
              ReDim Preserve aLanc(UBound(aLanc) + 1)
              aLanc(UBound(aLanc)) = sLANCAMENTO
           End If
           sSeq = .CellText(x, 3)
           sParc = IIf(.CellText(x, 4) = "Unica", "00", .CellText(x, 4))
           sComp = .CellText(x, 5)
           sSit = Left$(.CellText(x, 6), 2)
           sVencto = .CellText(x, 7)
'
           sDA = .CellText(x, 8)
           sAj = .CellText(x, 9)
           
            '***********************
                       'CARREGA O EXTRATO
            Set qd.ActiveConnection = cn
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
            qd(0) = nCodReduz
            qd(1) = nCodReduz
            qd(2) = Val(sAno)
            qd(3) = Val(sAno)
            qd(4) = Val(sLanc)
            qd(5) = Val(sLanc)
            qd(6) = Val(sSeq)
            qd(7) = Val(sSeq)
            qd(8) = Val(sParc)
            qd(9) = Val(sParc)
            qd(10) = Val(sComp)
            qd(11) = Val(sComp)
            qd(12) = 1
            qd(13) = 99
            qd(14) = Format(dVencto, "mm/dd/yyyy")
            qd(15) = NomeDeLogin
            Set RdoAux = qd.OpenResultset(rdOpenKeyset)
            With RdoAux
                sDataBase = !DATADEBASE
                sTributo = "": nValorPrincipal = 0: nValorJuros = 0: nValorMulta = 0: nValorCorrecao = 0: nValorTotal = 0: bITBI = False
                Do Until .EOF
                    If !CodTributo = 84 Then bITBI = True
                    sTributo = sTributo & Format(!CodTributo, "000") & "-" & !ABREVTRIBUTO & "/ "
                    nValorPrincipal = nValorPrincipal + !ValorTributo
                    nValorJuros = nValorJuros + !ValorJuros
                    nValorCorrecao = nValorCorrecao + !VALORCORRECAO
                    If MI And !CodLancamento = 5 Then
                        nValorMulta = 0
                        nValorTotal = nValorTotal + !ValorTributo + !ValorJuros + !VALORCORRECAO
                    Else
                        nValorMulta = nValorMulta + !ValorMulta
                        nValorTotal = nValorTotal + !ValorTributo + !ValorMulta + !ValorJuros + !VALORCORRECAO
                    End If
                    
                   .MoveNext
                Loop
               .Close
            End With
           
'**************************************************************
           
           If bITBI Then
                sTributo = ""
                Sql = "SELECT * FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & Val(sAno)
                Sql = Sql & " AND CODLANCAMENTO=" & Val(sLanc) & " AND SEQLANCAMENTO=" & Val(sSeq) & " AND NUMPARCELA=" & Val(sParc)
                Sql = Sql & " AND CODCOMPLEMENTO=" & Val(sComp)
                
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount > 0 Then
                        Do Until .EOF
                             If UCase$(Left(!obs, 2)) <> "LA" Then
                                sTributo = sTributo & SubNull(!obs) & " "
                             End If
                            .MoveNext
                        Loop
                    End If
                   .Close
                End With
           End If
           
           nSomaTotal = nSomaTotal + nValorTotal
           nSomaPrincipal = nSomaPrincipal + nValorPrincipal
          If sTributo <> "" Then
            sTributo = Left(sTributo, Len(sTributo) - 1)
          End If
           grdTrib.AddItem "001" & Chr(9) & sTributo
                       
           If chkMulta.Value = 1 Then
                nValorTotal = nValorTotal - nValorJuros - nValorMulta
                nValorJuros = 0: nValorMulta = 0
           End If
           If chkCorrecao.Value = 1 Then
                nValorTotal = nValorTotal - nValorCorrecao
                nValorCorrecao = 0
           End If
                       
                       
           grdTemp.AddItem sAno & Chr(9) & sLanc & Chr(9) & sSeq & Chr(9) & sParc & Chr(9) & _
           sComp & Chr(9) & sSit & Chr(9) & sVencto & Chr(9) & sDA & Chr(9) & sAj & Chr(9) & _
           FormatNumber(nValorPrincipal, 2) & Chr(9) & FormatNumber(nValorCorrecao, 2) & Chr(9) & _
           FormatNumber(nValorMulta, 2) & Chr(9) & FormatNumber(nValorJuros, 2) & Chr(9) & FormatNumber(nValorTotal, 2)
           
        End If
    Next
    
    sLANCAMENTO = ""
    For Y = 1 To UBound(aLanc)
        sLANCAMENTO = sLANCAMENTO & aLanc(Y) & "/ "
    Next
    sLANCAMENTO = Left(sLANCAMENTO, Len(sLANCAMENTO) - 2)

End With

'***** ANISTIA *****
If bAnistia And chkAnistia.Value = vbChecked Then
    
    Achou = False
    With grdTemp
        For x = 1 To .Rows - 1
            If Val(.TextMatrix(x, 0)) <> 2012 Then
                If Year(CDate(.TextMatrix(x, 6))) < 2012 Then
                    Achou = True
                End If
            End If
        Next
    End With
    If Not Achou Then
        lblAnistia.Caption = "0,00"
        GoTo CONTINUA
    End If
    

    With grdTemp
        For x = 1 To .Rows - 1
            If CDate(.TextMatrix(x, 6)) > CDate("31/12/2011") And .TextMatrix(x, 1) <> 41 And .TextMatrix(x, 1) <> 69 And .TextMatrix(x, 1) <> 5 Then
                lblAnistia.Caption = "0,00"
                GoTo CONTINUA
            End If
        Next
    End With
    
    If dVencto <= CDate("31/10/2012") Then
       lblAnistia.Caption = "100,00"
    ElseIf dVencto <= CDate("30/11/2012") Then
       lblAnistia.Caption = "95,00"
    ElseIf dVencto <= CDate("28/12/2012") Then
       lblAnistia.Caption = "90,00"
    ElseIf dVencto >= CDate("28/12/2012") Then
       If chkMulta.Value = 0 Then
          lblAnistia.Caption = "0,00"
       Else
          lblAnistia.Caption = "100,00"
       End If
       GoTo CONTINUA
    End If
    nPerc = 100 - CDbl(lblAnistia.Caption)
    With grdTemp
        For x = 1 To grdTemp.Rows - 1
            If .TextMatrix(x, 1) <> 41 And .TextMatrix(x, 1) <> 69 Then
                .TextMatrix(x, 11) = FormatNumber(CDbl(.TextMatrix(x, 11)) * nPerc / 100, 2)
                .TextMatrix(x, 12) = FormatNumber(CDbl(.TextMatrix(x, 12)) * nPerc / 100, 2)
            End If
            .TextMatrix(x, 13) = FormatNumber(CDbl(.TextMatrix(x, 9)) + CDbl(.TextMatrix(x, 10)) + CDbl(.TextMatrix(x, 11)) + CDbl(.TextMatrix(x, 12)), 2)
        Next
    End With
Else
    lblAnistia.Caption = "0,00"
End If
'*******************
CONTINUA:

Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE ANOEXPED = " & Year(Now) & " AND CODLANCAMENTO = 1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     lblValorExp.Caption = FormatNumber(!VALORDAM, 2)
     lblValorExp2.Caption = FormatNumber(!VALORDAM, 2)
     lblTotalLanc.Caption = FormatNumber(nSomaTotal, 2)
    .Close
End With

'taxa expediente
grdTemp.AddItem Year(Now) & Chr(9) & "004" & Chr(9) & "00" & Chr(9) & "01" & Chr(9) & _
  "0" & Chr(9) & "03" & Chr(9) & Format(Now, "dd/mm/yyyy") & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & _
  lblValorExp.Caption & Chr(9) & FormatNumber(0, 2) & Chr(9) & _
  FormatNumber(0, 2) & Chr(9) & FormatNumber(0, 2) & Chr(9) & lblValorExp.Caption
sTributo = "003" & "-" & "TAXA EXP.DOC."
grdTrib.AddItem "003" & Chr(9) & sTributo

'HONORARIOS
nSomaHon = 0
If bHonorario Then
    For x = 1 To grdTemp.Rows - 2
        nSomaHon = nSomaHon + (grdTemp.TextMatrix(x, 13) * 10 / 100)
    Next
    nSomaHon = FormatNumber(nSomaHon, 2)
    grdTemp.AddItem Year(Now) & Chr(9) & "041" & Chr(9) & "00" & Chr(9) & "01" & Chr(9) & _
    "0" & Chr(9) & "03" & Chr(9) & Format(Now, "dd/mm/yyyy") & Chr(9) & "N" & Chr(9) & "N" & Chr(9) & _
    nSomaHon & Chr(9) & FormatNumber(0, 2) & Chr(9) & _
    FormatNumber(0, 2) & Chr(9) & FormatNumber(0, 2) & Chr(9) & nSomaHon
    sTributo = "041" & "-" & "HONORÁRIOS"
    grdTrib.AddItem "041" & Chr(9) & sTributo
End If

If lblValorExp.Caption > 0 Then
    nSomaPrincipal = nSomaPrincipal + CDbl(lblValorExp.Caption)
End If
CalculaTotal

End Sub

Private Sub GravaDam()

On Error GoTo Erro

Dim x As Integer
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, qd As New rdoQuery, RdoAux3 As rdoResultset
Dim sNumInsc As String, sValorParc As String, sData As String, sObs As String
Dim nCodReduz As Long
Dim sNomeResp As String
Dim sEndImovel As String
Dim nNumImovel As Integer
Dim sComplImovel As String
Dim sBairroImovel As String
Dim nCodCidade As Integer
Dim nCodBairro As Integer
Dim sCidadeEntrega As String
Dim sUFEntrega As String
Dim sCPF As String
Dim nAno As Integer
Dim nNumDoc As Long
Dim sQuadra As String
Dim sLote As String
Dim nNumParc As Integer
Dim dDataVencto As Date
Dim nCodLanc As Integer
Dim nSeq As Integer, nSeq2 As Integer
Dim nComplemento As Integer
Dim nValorTotal As Double
Dim NumBarra2 As String
Dim NumBarra2a As String
Dim NumBarra2b As String
Dim NumBarra2c As String
Dim NumBarra2d As String
Dim StrBarra2 As String
Dim nLastCod As Long
Dim nValorTaxa As Double
Dim nLanc As Integer, nParc As Integer
Dim nComp As Integer, bMulta As Boolean
Dim nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nSID As Long

If MsgBox("Confirma Emissão da DAM ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nSID = Int(Rnd(10) * 1000000)
lblSid.Caption = nSID

'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "' and usuario='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSID
cn.Execute Sql, rdExecDirect

'RETORNA VALOR EXPEDIENTE
Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValorTaxa = RdoAux!VALORDAM
RdoAux.Close

nCodReduz = Val(frmDebitoImob.txtCod.Text)
sNomeResp = frmDebitoImob.lblProp.Caption

Select Case nCodReduz
    Case 1 To 99999
        Sql = "SELECT * FROM vwCnsImovel WHERE CODREDUZIDO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
             If .RowCount > 0 Then
                sNumInsc = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
                sEndImovel = Trim$(!AbrevTipoLog) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                nNumImovel = Val(SubNull(!Li_Num))
                sComplImovel = SubNull(!Li_Compl)
                nCodBairro = !codbairro
                sCidadeEntrega = SubNull(!desccidade)
                sUFEntrega = SubNull(!LI_UF)
                nCodCidade = !LI_CODCIDADE
                sQuadra = SubNull(!Li_Quadras)
                sLote = SubNull(!Li_Lotes)
                Sql = "SELECT CODREDUZIDO,CPF,CNPJ,RG,ORGAO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodReduz
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If Not IsNull(!CPF) Then
                       sCPF = !CPF
                    ElseIf Not IsNull(!Cnpj) Then
                       sCPF = !Cnpj
                    ElseIf Not IsNull(!rg) Then
                       sCPF = !rg
                    Else
                       sCPF = ""
                    End If
                End With
            End If
        End With
     Case 100000 To 500000
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,CNPJ,CPF,RAZAOSOCIAL,DESCCIDADE,SIGLAUF,CODCIDADE,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO,NOMELOGR "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                sNumInsc = !INSCESTADUAL
                sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                If Trim(sEndImovel) = "" Then
                    sEndImovel = SubNull(!NomeLogr)
                End If
                nNumImovel = Val(SubNull(!Numero))
                sComplImovel = SubNull(!COMPLEMENTO)
                nCodBairro = !codbairro
                sCidadeEntrega = SubNull(!desccidade)
                sUFEntrega = SubNull(!siglauf)
                nCodCidade = !CODCIDADE
                sQuadra = "0"
                sLote = "0"
                If Not IsNull(!CPF) Then
                   sCPF = !CPF
                ElseIf Not IsNull(!Cnpj) Then
                   sCPF = !Cnpj
                Else
                    sCPF = ""
                End If
            End If
         End With
     Case 500000 To 800000
        Sql = "SELECT cidadao.codcidadao, cidadao.cpf, cidadao.cnpj, cidadao.rg, cidadao.numimovel, cidadao.complemento, cidadao.codbairro, cidadao.codcidade, "
        Sql = Sql & "cidadao.siglauf, cidade.desccidade, bairro.descbairro, cidadao.nomelogradouro AS nomerua, cidadao.nomebairro, cidadao.nomecidade,"
        Sql = Sql & "Cidadao.codlogradouro , vwLOGRADOURO.AbrevTipoLog, vwLOGRADOURO.AbrevTitLog, vwLOGRADOURO.NomeLogradouro FROM bairro RIGHT OUTER JOIN "
        Sql = Sql & "cidade RIGHT OUTER JOIN cidadao ON cidade.siglauf = cidadao.siglauf AND cidade.codcidade = cidadao.codcidade LEFT OUTER JOIN "
        Sql = Sql & "vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO ON bairro.siglauf = cidadao.siglauf AND bairro.codcidade = Cidadao.codcidade And bairro.codbairro = Cidadao.codbairro "
        Sql = Sql & "WHERE CIDADAO.CODCIDADAO=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                If Not IsNull(!NomeLogradouro) Then
                    sEndImovel = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
                Else
                    sEndImovel = SubNull(!nomerua)
                End If
                nNumImovel = Val(SubNull(!NUMIMOVEL))
                sComplImovel = SubNull(!COMPLEMENTO)
                nCodBairro = Val(SubNull(!codbairro))
                nCodCidade = Val(SubNull(!CODCIDADE))
                sCidadeEntrega = SubNull(!desccidade)
                sUFEntrega = SubNull(!siglauf)
                If Not IsNull(!CPF) And Trim$(SubNull(!CPF)) <> "" Then
                   sCPF = !CPF
                ElseIf Not IsNull(!Cnpj) And Trim$(SubNull(!Cnpj)) <> "" Then
                   sCPF = !Cnpj
                ElseIf Not IsNull(!rg) Then
                   sCPF = !rg
                Else
                   sCPF = ""
                End If
             Else
                sCPF = ""
             End If
             If sCidadeEntrega = "" Then
                sCidadeEntrega = SubNull(!NomeCidade)
             End If
             If nCodBairro = 0 Then
                sBairroImovel = SubNull(!NOMEBairro)
                
                GoTo FIMBAIRRO
             End If
        End With
End Select


If nCodCidade = 0 Then GoTo FIMBAIRRO
Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & sUFEntrega & "' AND CODCIDADE=" & nCodCidade & " AND CODBAIRRO=" & nCodBairro
Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux3
    If .RowCount > 0 Then
         sBairroImovel = !DescBairro
    Else
         sBairroImovel = ""
    End If
   .Close
End With
FIMBAIRRO:
'TOTAL
nValorTotal = CDbl(lblTotalLanc.Caption)


'RETORNA ULTIMO DOCUMENTO
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!MAXIMO) Then
   nLastCod = 1
Else
   nLastCod = RdoAux!MAXIMO + 1
End If
RdoAux.Close
If dVencto = "00:00:00" Then dVencto = Format(Now, "dd/mm/yyyy")
'GERAÇÃO DOS DÉBITOS
With grdTemp
   'GRAVA NUMDOCUMENTO
    If chkMulta.Value = 1 Then
       bMulta = True
    Else
       If Val(lblAnistia.Caption) > 0 Then
           bMulta = True
       Else
           bMulta = False
       End If
    End If
       
    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO) VALUES("
    If chkJulgamento.Value = 0 Then
        Sql = Sql & nLastCod & ",'" & Format(dVencto, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ")"
    Else
        Sql = Sql & nLastCod & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ")"
    End If
    cn.Execute Sql, rdExecDirect
    For x = 1 To .Rows - 2
        nAno = Val(.TextMatrix(x, 0))
        nLanc = Val(.TextMatrix(x, 1))
        nSeq = Val(.TextMatrix(x, 2))
        nParc = Val(.TextMatrix(x, 3))
        nComp = Val(.TextMatrix(x, 4))
        nValorJuros = FormatNumber(CDbl(grdTemp.TextMatrix(x, 12)), 2)
        nValorMulta = FormatNumber(CDbl(.TextMatrix(x, 11)), 2)
        nValorCorrecao = FormatNumber(CDbl(.TextMatrix(x, 10)), 2)
       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,VALORJUROS,VALORMULTA,VALORCORRECAO) VALUES(" & nCodReduz & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nComp & "," & nLastCod & ","
        Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & ")"
        cn.Execute Sql, rdExecDirect
        If Val(lblAnistia.Caption) > 0 And bAnistia Then
            'GRAVA OBS PARCELA
             Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
             Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
             Sql = Sql & " AND CODCOMPLEMENTO=" & nComp
             Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux
                 If IsNull(!MAXIMO) Then
                     nSeq2 = 1
                 Else
                     nSeq2 = !MAXIMO + 1
                 End If
                .Close
             End With
             sData = Right$(frmMdi.Sbar.Panels(6).Text, 10)
             sObs = "Lancamento incluido na DAM número " & nLastCod & " com " & lblAnistia.Caption & "% de desconto em multa e juros conforme REFIS-IV"
             Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduz & "," & nAno & ","
             Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nComp & "," & nSeq2 & ",'" & sObs & "','" & NomeDeLogin & "','" & Format(sData, "mm/dd/yyyy") & "')"
             cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'CRIA VINCULO COM ISSELETRONICO SE HOUVER
If bAnistia Then
    If UBound(aDocDAM) > 0 Then
        For x = 1 To UBound(aDocDAM)
            Sql = "INSERT DAMISS(DOCDAM,DOCISS) VALUES(" & nLastCod & "," & aDocDAM(x) & ")"
            cn.Execute Sql, rdExecDirect
        Next
        ReDim aDocDAM(0)
    End If
End If

'ATUALIZA DEBITOTRIBUTO
With grdTemp
    For x = 1 To .Rows - 1
        nAno = Val(.TextMatrix(x, 0))
        nLanc = Val(.TextMatrix(x, 1))
        nSeq = Val(.TextMatrix(x, 2))
        nParc = Val(.TextMatrix(x, 3))
        nComp = Val(.TextMatrix(x, 4))
        If nLanc <> 4 Then
            Sql = "UPDATE DEBITOTRIBUTO SET VALORCORRECAO=" & Virg2Ponto(RemovePonto(.TextMatrix(x, 10))) & ","
            Sql = Sql & "VALORMULTA=" & Virg2Ponto(RemovePonto(.TextMatrix(x, 11))) & ","
            Sql = Sql & "VALORJUROS=" & Virg2Ponto(RemovePonto(.TextMatrix(x, 12))) & " WHERE CODREDUZIDO=" & nCodReduz
            Sql = Sql & " AND ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc
            Sql = Sql & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
            Sql = Sql & " AND CODCOMPLEMENTO=" & nComp
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With

'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "' and usuario='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSID
cn.Execute Sql, rdExecDirect

Set qd.ActiveConnection = cn

'GRAVA TEMPORARIO
With grdTemp
    For x = 1 To .Rows - 1
        nAno = Year(dVencto)
        nCodLanc = 1
        nSeq = 0
        nNumParc = 1
        nComplemento = 0
        dDataVencto = Format(dVencto, "dd/mm/yyyy")
        sValorParc = FormatNumber(lblValorTotal.Caption, 2)
        nNumDoc = nLastCod
        NumBarra2 = Gera2of5Cod(sValorParc, dDataVencto, nNumDoc, nNumParc, nCodLanc, nSeq, nComplemento)
        NumBarra2a = Left$(NumBarra2, 13)
        NumBarra2b = Mid$(NumBarra2, 14, 13)
        NumBarra2c = Mid$(NumBarra2, 27, 13)
        NumBarra2d = Right$(NumBarra2, 13)
        StrBarra2 = Gera2of5Str(Left$(NumBarra2a, 11) & Left$(NumBarra2b, 11) & Left$(NumBarra2c, 11) & Left$(NumBarra2d, 11))
        
        Sql = "INSERT DAM(COMPUTER,SEQ,INSCRICAO,CODREDUZIDO,TIPOIMPOSTO,NOMECONTRIBUINTE,CPF,ENDERECO,NUMERO,COMPLEMENTO,"
        Sql = Sql & "BAIRRO,CIDADE,UF,QUADRA,LOTE,FULLLANC,FULLTRIB,NUMDAM,ANOEXERC,LANC,NUMSEQ,NUMPARCELA,COMP,DATAVENCTO,"
        Sql = Sql & "SIT,AJ,DA,PRINCIPAL,CORRECAO,MULTA,JUROS,TOTAL,STRBARRA2,NUMBARRA2A,NUMBARRA2B,NUMBARRA2C,NUMBARRA2D,"
        Sql = Sql & "VALORDAM,VALORPRINCDAM,CODTRIBUTO,USUARIO,SID) VALUES('" & NomeDoComputador & "'," & x & ",'" & sNumInsc & "','"
        Sql = Sql & Format(nCodReduz, "000000") & "','" & "DAM" & "','" & Mask(Left$(sNomeResp, 40)) & "','" & Left(sCPF, 20) & "','" & Left$(sEndImovel, 40) & "',"
        Sql = Sql & nNumImovel & ",'" & Left$(sComplImovel, 30) & "','" & Left$(sBairroImovel, 25) & "','" & sCidadeEntrega & "','" & sUFEntrega & "','"
        Sql = Sql & Left(Mask(sQuadra), 15) & "','" & Left$(Mask(sLote), 10) & "','" & sLANCAMENTO & "','" & Left$(Mask(grdTrib.TextMatrix(x, 1)), 2000) & "','"
        Sql = Sql & CStr(nLastCod) & CStr(RetornaDVNumDoc(nLastCod)) & "','" & .TextMatrix(x, 0) & "','" & .TextMatrix(x, 1) & "','"
        Sql = Sql & .TextMatrix(x, 2) & "','" & .TextMatrix(x, 3) & "','" & .TextMatrix(x, 4) & "','" & Format(.TextMatrix(x, 6), "mm/dd/yyyy") & "','"
        Sql = Sql & .TextMatrix(x, 5) & "','" & .TextMatrix(x, 7) & "','" & .TextMatrix(x, 8) & "'," & Virg2Ponto(RemovePonto(.TextMatrix(x, 9))) & ","
        Sql = Sql & Virg2Ponto(RemovePonto(.TextMatrix(x, 10))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 11))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 12))) & "," & Virg2Ponto(RemovePonto(.TextMatrix(x, 13))) & ",'"
        Sql = Sql & Mask(StrBarra2) & "','" & NumBarra2a & "','" & NumBarra2b & "','" & NumBarra2c & "','" & NumBarra2d & "'," & Virg2Ponto(RemovePonto(sValorParc)) & ","
        Sql = Sql & Virg2Ponto(CStr(Format(nSomaPrincipal, "#0.00"))) & "," & Val(Left$(Mask(grdTrib.TextMatrix(x, 1)), 3)) & ",'" & NomeDeLogin & "'," & nSID & ")"
        cn.Execute Sql, rdExecDirect
    Next
End With

modLg "Emissão de DAM nº " & CStr(nLastCod)

nNumDoc = lblSid.Caption
'EXIBE RELATORIO
If bHonorario Then
    frmReport.ShowReport "DAMHONORARIO", frmMdi.hwnd, Me.hwnd, nNumDoc
Else
    If frmMdi.frTeste.Visible = True Then
        frmReport.ShowReport "DAMTMP", frmMdi.hwnd, Me.hwnd, nNumDoc
    Else
        frmReport.ShowReport "DAM", frmMdi.hwnd, Me.hwnd, nNumDoc
    End If
End If

'DELETA TEMPORARIO
'Sql = "DELETE FROM DAM WHERE COMPUTER='" & NomeDoUsuario & "' and usuario='" & NomeDeLogin & "'"
Sql = "DELETE FROM DAM WHERE SID=" & nSID
cn.Execute Sql, rdExecDirect

Exit Sub

Erro:
For Y = 0 To rdoErrors.Count - 1
     MsgBox rdoErrors(Y).Description
Next
Resume Next

End Sub

Private Sub CalculaTotal()
Dim nTotal As Double, x As Integer

For x = 1 To grdTemp.Rows - 1
    nTotal = nTotal + grdTemp.TextMatrix(x, 13)
Next

lblValorTotal.Caption = FormatNumber(nTotal, 2)
End Sub

Private Function CalculaCorrecaoDAM(nValorDebito As Double, dDataBase As Date) As Double

Dim RdoAux As rdoResultset, Sql As String
Dim UfirAtual As Double
Dim UfirBase As Double, dDataVencto As Date

If chkVenctoAtual.Value = 0 Then
    dDataVencto = Format(mskVencimento.Text)
Else
    dDataVencto = Format(mskVenc.Text)
End If

If Year(dDataBase) > Year(dDataVencto) Then
    CalculaCorrecaoDAM = 0
    Exit Function
End If

UfirAtual = RetornaUFIR(Year(dDataVencto))
If UfirAtual = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano atual.", vbCritical, "Alerta !!!"
    CalculaCorrecaoDAM = 0
    Exit Function
End If

UfirBase = RetornaUFIR(Year(dDataBase))
If UfirBase = 0 Then
    MsgBox "Não foi cadastrado o valor da Ufir para o ano base.", vbCritical, "Alerta !!!"
    CalculaCorrecaoDAM = 0
    Exit Function
End If

CalculaCorrecaoDAM = (nValorDebito * UfirAtual / UfirBase) - nValorDebito
If CalculaCorrecaoDAM > 0 Then
   CalculaCorrecaoDAM = FormatNumber(CalculaCorrecaoDAM, 2)
End If
End Function

Private Function CalculaJurosDAM(nValorDebito As Double, dDataVencto As Date) As Double
Dim nNumMes As Integer
Dim nValorPerc As Double
Dim sDataVencto As String, nDia As Integer, nMes As Integer, nAno As Integer

'If dDataNow = "00:00:00" Then
 dDataNow = Now
'End If

'SE O VENCIMENTO FOR MAIOR OU IGUAL A DATA ATUAL, NÃO EXISTE JUROS
If dDataVencto >= dDataNow Then
    CalculaJurosDAM = 0
    Exit Function
End If

'SE ESTIVER NO MESMO MES E ANO QUE A DATA ATUAL, NAO EXISTE JUROS
If Month(dDataVencto) = Month(dDataNow) And Year(dDataVencto) = Year(dDataNow) Then
    CalculaJurosDAM = 0
    Exit Function
End If

If Not dcJuros.Exists(Year(dDataNow)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJurosDAM = 0
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
nNumMes = Int(DateDiff("d", dDataVencto, dDataNow) / 30) + 1


'If chkVenctoAtual.Value = 0 Then
'    dDataNow = Now
'Else
'    dDataNow = Format(mskVenc.text, "dd/mm/yyyy")
'End If

'If dDataVencto >= dDataNow Then
'    CalculaJurosDAM = 0
'    Exit Function
'End If
'nNumMes = Int((DateDiff("d", dDataVencto, dDataNow)) / 30)

If Not dcJuros.Exists(Year(dDataNow)) Then
   MsgBox "Não foi cadastrado o valor do juros para o ano atual.", vbCritical, "Alerta !!!"
   CalculaJurosDAM = 0
   Exit Function
End If
nValorPerc = dcJuros.Item(Year(dDataNow))

nValorPerc = nValorPerc / 100

CalculaJurosDAM = nValorDebito * nValorPerc * nNumMes
If CalculaJurosDAM > 0 Then
   CalculaJurosDAM = FormatNumber(CalculaJurosDAM, 3)
End If

End Function

Private Sub mnuA1_Click()
mskVencimento.Text = "31/10/2012"
mskVencimento_LostFocus
End Sub

Private Sub mnuA2_Click()
mskVencimento.Text = "30/11/2012"
mskVencimento_LostFocus
End Sub

Private Sub mnuA3_Click()
mskVencimento.Text = "28/12/2012"
mskVencimento_LostFocus
End Sub

Private Sub mskVencimento_GotFocus()
mskVencimento.SelStart = 0
mskVencimento.SelLength = Len(mskVencimento.Text)
End Sub

Private Sub mskVencimento_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    mskVencimento_LostFocus
End If
End Sub

Private Sub mskVencimento_LostFocus()
Dim bValid As Boolean
bValid = False

If Not IsDate(mskVencimento.Text) Then
   MsgBox "Data inválida.", vbCritical, "Atenção"
   LimpaMascara mskVencimento
   bValid = False
Else
   If CDate(mskVencimento.Text) < Format(Now, "dd/mm/yyyy") Then
      MsgBox "Data de vencimento não pode ser retroativa.", vbCritical, "Atenção"
      LimpaMascara mskVencimento
      bValid = False
   Else
      bValid = True
   End If
End If

If bValid Then
    dDataVencto = Format(mskVencimento.Text)
    CarregaLista2
End If
End Sub

Private Sub EmiteBoleto()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, Sql As String, nPos As Integer, sDataDam As String, sDataVencto As String
Dim nCodReduz As Long, sInsc As String, sNome As String, sDoc As String, sEnd As String, nNum As Integer, nValorDoc As Double
Dim sCompl As String, sBairro As String, sCidade As String, sUF As String, sQuadras As String, sLotes As String
Dim sUsuario As String, nNumDoc As Long, bMulta As Boolean, nValorTaxa As Double, sNumDoc As String, bGerado As Boolean
Dim sLanc As String, sFullTrib As String, nAno As Integer, nSeq As Integer, nLanc As Integer, nParc As Integer, nCompl As Integer, nValorJuros As Double, nValorMulta As Double, nValorCorrecao As Double, nValorTotal As Double
Dim nSeq2 As Integer, sAj As String, sDA As String, nValorPrincipal As Double, sNumDoc2 As String, sNumDoc3 As String, nFatorVencto As Long
Dim nSID As Long, sDigitavel As String, sNossoNumero As String, sDv As String, sQuintoGrupo As String, dDataBase As Date

If MsgBox("Confirma Emissão da DAM ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then
   bGerado = False
   Exit Sub
End If

nSID = Int(Rnd(100) * 1000000)

Sql = "delete from boleto where sid=" & nSID
cn.Execute Sql, rdExecDirect

'RETORNA VALOR EXPEDIENTE
Sql = "SELECT VALORDAM FROM EXPEDIENTE WHERE CODLANCAMENTO=3 AND ANOEXPED=" & Year(Now)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nValorTaxa = RdoAux!VALORDAM
RdoAux.Close

bMulta = False
sDoc = ""
nPos = 0
nValorDoc = 0
nCodReduz = Val(frmDebitoImob.txtCod.Text)
sUsuario = NomeDeLogin
sDataDam = mskVencimento.Text

Select Case nCodReduz
    
    Case 1 To 99999
        Sql = "select * from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            sInsc = !Inscricao
            sNome = !nomecidadao
            sDoc = SubNull(!CPF)
            If sDoc = "" Then
                sDoc = SubNull(!Cnpj)
                If sDoc = "" Then
                    sDoc = SubNull(!rg)
                End If
            End If
            sEnd = SubNull(!Logradouro)
            nNum = Val(SubNull(!Li_Num))
            sCompl = Left(SubNull(!Li_Compl), 30)
            sBairro = SubNull(!DescBairro)
            sCidade = SubNull(!desccidade)
            sUF = SubNull(!LI_UF)
            sQuadras = Left(SubNull(!Li_Quadras), 15)
            sLotes = Left(SubNull(!Li_Lotes), 10)
           .Close
        End With

End Select

'grava documento
Sql = "SELECT MAX(NUMDOCUMENTO) AS MAXIMO FROM NUMDOCUMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!MAXIMO) Then
   nNumDoc = 0
Else
   nNumDoc = RdoAux!MAXIMO + 1
End If
RdoAux.Close

If chkMulta.Value = 1 Then
   bMulta = True
Else
   If Val(lblAnistia.Caption) > 0 Then
       bMulta = True
   Else
       bMulta = False
   End If
End If

With grdTemp
   'GRAVA NUMDOCUMENTO
    If chkMulta.Value = 1 Then
       bMulta = True
    Else
       If Val(lblAnistia.Caption) > 0 Then
           bMulta = True
       Else
           bMulta = False
       End If
    End If

    Sql = "INSERT NUMDOCUMENTO (NUMDOCUMENTO,DATADOCUMENTO,CODBANCO,CODAGENCIA,VALORPAGO,VALORTAXADOC,ISENTOMJ,PERCISENCAO) VALUES("
    If chkJulgamento.Value = 0 Then
        Sql = Sql & nNumDoc & ",'" & Format(sDataDam, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ")"
    Else
        Sql = Sql & nNumDoc & ",'" & Format(Now, "mm/dd/yyyy") & "'," & 0 & "," & 0 & "," & 0 & "," & Virg2Ponto(CStr(Round(nValorTaxa, 2))) & "," & IIf(bMulta, 1, 0) & "," & Virg2Ponto(lblAnistia.Caption) & ")"
    End If
    cn.Execute Sql, rdExecDirect

    sNumDoc = CStr(nNumDoc) & "-" & RetornaDVNumDoc(nNumDoc)
    sNumDoc2 = CStr(nNumDoc) & RetornaDVNumDoc(nNumDoc)
    sNumDoc3 = CStr(nNumDoc) & Modulo11(nNumDoc)
    
    For x = 1 To .Rows - 1
        nAno = Val(.TextMatrix(x, 0))
        nLanc = Val(.TextMatrix(x, 1))
        nSeq = Val(.TextMatrix(x, 2))
        nParc = Val(.TextMatrix(x, 3))
        nCompl = Val(.TextMatrix(x, 4))
        sDataVencto = .TextMatrix(x, 6)
        sDA = .TextMatrix(x, 7)
        sAj = .TextMatrix(x, 8)
        nValorPrincipal = FormatNumber(CDbl(.TextMatrix(x, 9)), 2)
        nValorJuros = FormatNumber(CDbl(grdTemp.TextMatrix(x, 12)), 2)
        nValorMulta = FormatNumber(CDbl(.TextMatrix(x, 11)), 2)
        nValorCorrecao = FormatNumber(CDbl(.TextMatrix(x, 10)), 2)
        nValorTotal = FormatNumber(CDbl(.TextMatrix(x, 13)), 2)
        nValorDoc = nValorDoc + nValorTotal
        sFullTrib = Left$(Mask(grdTrib.TextMatrix(x, 1)), 2000)
        sFullTrib = Left(sFullTrib, Len(sFullTrib) - 1)
       'GRAVA PARCELADOCUMENTO
        Sql = "INSERT PARCELADOCUMENTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,"
        Sql = Sql & "NUMPARCELA,CODCOMPLEMENTO,NUMDOCUMENTO,VALORJUROS,VALORMULTA,VALORCORRECAO) VALUES(" & nCodReduz & ","
        Sql = Sql & nAno & "," & nLanc & "," & nSeq & "," & nParc & "," & nCompl & "," & nNumDoc & ","
        Sql = Sql & Virg2Ponto(CStr(nValorJuros)) & "," & Virg2Ponto(CStr(nValorMulta)) & "," & Virg2Ponto(CStr(nValorCorrecao)) & ")"
        cn.Execute Sql, rdExecDirect
        If Val(lblAnistia.Caption) > 0 And bAnistia Then
            'GRAVA OBS PARCELA
             Sql = "SELECT MAX(SEQ) AS MAXIMO FROM OBSPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND ANOEXERCICIO=" & nAno
             Sql = Sql & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc
             Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
             Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
             With RdoAux
                 If IsNull(!MAXIMO) Then
                     nSeq2 = 1
                 Else
                     nSeq2 = !MAXIMO + 1
                 End If
                .Close
             End With
             sObs = "Lancamento incluido na DAM número " & nLastCod + 1 & " com " & lblAnistia.Caption & "% de desconto em multa e juros conforme REFIS-IV"
             Sql = "INSERT OBSPARCELA(CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,SEQ,OBS,USUARIO,DATA) VALUES(" & nCodReduz & "," & nAno & ","
             Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nComp & "," & nSeq2 & ",'" & sObs & "','" & NomeDeLogin & "','" & Format(Now, "mm/dd/yyyy") & "')"
             cn.Execute Sql, rdExecDirect
        End If
    
        Sql = "insert boleto(usuario,computer,sid,seq,inscricao,codreduzido,nome,cpf,endereco,numimovel,complemento,bairro,cidade,uf,quadra,lote,numdoc,nomefunc,datadam,fulllanc,fulltrib,"
        Sql = Sql & "anoexercicio,codlancamento,seqlancamento,numparcela,codcomplemento,datavencto,aj,da,principal,juros,multa,correcao,total,numdoc2) values('"
        Sql = Sql & NomeDeLogin & "','" & NomeDoComputador & "'," & nSID & "," & nPos & ",'" & sInsc & "'," & nCodReduz & ",'" & Mask(sNome) & "','" & sDoc & "','"
        Sql = Sql & Mask(sEnd) & "'," & nNum & ",'" & Mask(sCompl) & "','" & Mask(sBairro) & "','" & Mask(sCidade) & "','" & sUF & "','" & Mask(sQuadras) & "','"
        Sql = Sql & Mask(sLotes) & "','" & sNumDoc & "','" & NomeDeLogin & "','" & Format(sDataDam, "mm/dd/yyyy") & "','" & sLANCAMENTO & "','" & sFullTrib & "'," & nAno & ","
        Sql = Sql & nLanc & "," & nSeq & "," & nParc & "," & nCompl & ",'" & Format(sDataVencto, "mm/dd/yyyy") & "','" & sAj & "','" & sDA & "'," & Virg2Ponto(Format(nValorPrincipal, "#0.00")) & ","
        Sql = Sql & Virg2Ponto(Format(nValorJuros, "#0.00")) & "," & Virg2Ponto(Format(nValorMulta, "#0.00")) & "," & Virg2Ponto(Format(nValorCorrecao, "#0.00")) & "," & Virg2Ponto(Format(nValorTotal, "#0.00")) & ",'" & sNumDoc2
        Sql = Sql & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
    
    Next

End With

sNossoNumero = Format(sNumDoc3, "0000000000000")
sDv = Trim(Calculo_DV10("028" & Left(sNossoNumero, 7)))
sDigitavel = "0339912354028" & Left(sNossoNumero, 7) & sDv

sDigitavel = sDigitavel & Right(sNossoNumero, 6) & "0102"
sDv = Trim(Calculo_DV10(Right(sDigitavel, 10)))
sDigitavel = sDigitavel & sDv

dDataBase = "07/10/1997"
nFatorVencto = CDate(sDataDam) - dDataBase
sQuintoGrupo = Format(nFatorVencto, "0000")
sQuintoGrupo = sQuintoGrupo & Format(RetornaNumero(FormatNumber(nValorDoc, 2)), "0000000000")

sDv = Calculo_DV11(sDigitavel & sQuintoGrupo)
sDigitavel = sDigitavel & sDv & sQuintoGrupo

Sql = "update boleto set digitavel='" & sDigitavel & "' where sid=" & nSID
cn.Execute Sql, rdExecDirect

frmReport.ShowReport2 "BOLETODAM", frmMdi.hwnd, Me.hwnd, nSID

Sql = "delete from boleto where sid=" & nSID
cn.Execute Sql, rdExecDirect

End Sub
