VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmConfissaoDivida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Termo de Confissão de Dívida Fiscal"
   ClientHeight    =   3525
   ClientLeft      =   12915
   ClientTop       =   3840
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   8970
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCPF 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2655
      MaxLength       =   20
      TabIndex        =   31
      Top             =   2520
      Width           =   2670
   End
   Begin VB.TextBox txtRequerente 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2655
      MaxLength       =   100
      TabIndex        =   30
      Top             =   1845
      Width           =   6225
   End
   Begin VB.TextBox txtNumDoc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2655
      MaxLength       =   9
      TabIndex        =   29
      Top             =   585
      Width           =   1695
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Documento"
      Height          =   195
      Index           =   1
      Left            =   1620
      TabIndex        =   28
      Top             =   135
      Width           =   1365
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Processo"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   27
      Top             =   135
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.TextBox txtNumProc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2670
      TabIndex        =   0
      Top             =   585
      Width           =   1695
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   8400
      TabIndex        =   1
      ToolTipText     =   "Consulta Cidadão"
      Top             =   1845
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      MICON           =   "frmConfissaoDivida.frx":0000
      PICN            =   "frmConfissaoDivida.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   345
      Left            =   7770
      TabIndex        =   12
      ToolTipText     =   "Sair da Tela"
      Top             =   3105
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmConfissaoDivida.frx":0176
      PICN            =   "frmConfissaoDivida.frx":0192
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
      Left            =   7770
      TabIndex        =   13
      ToolTipText     =   "Gera as guias informadas"
      Top             =   2715
      Width           =   1140
      _ExtentX        =   2011
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
      MICON           =   "frmConfissaoDivida.frx":0200
      PICN            =   "frmConfissaoDivida.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskVenc 
      Height          =   285
      Left            =   2655
      TabIndex        =   32
      Top             =   3150
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      BackColor       =   16777215
      MouseIcon       =   "frmConfissaoDivida.frx":0376
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
   Begin VB.Label lblDI 
      Caption         =   "N"
      Height          =   255
      Left            =   7650
      TabIndex        =   34
      Top             =   150
      Width           =   525
   End
   Begin VB.Label lblTipoDoc 
      Caption         =   "Documento nº (sem DV).:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   135
      TabIndex        =   33
      Top             =   630
      Width           =   2505
   End
   Begin VB.Label lblSid 
      Caption         =   "0"
      Height          =   195
      Left            =   6660
      TabIndex        =   26
      Top             =   630
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblDataProc 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   4500
      TabIndex        =   25
      Top             =   615
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblCod 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   24
      Top             =   945
      Width           =   1635
   End
   Begin VB.Label lblVenc 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   23
      Top             =   3195
      Width           =   1695
   End
   Begin VB.Label lblValor 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   22
      Top             =   2865
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Vencto 1ª Parcela.....:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   10
      Left            =   120
      TabIndex        =   21
      Top             =   3195
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Qtde Parcela.:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   9
      Left            =   4500
      TabIndex        =   20
      Top             =   3195
      Width           =   1575
   End
   Begin VB.Label lblQtdeParc 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6120
      TabIndex        =   19
      Top             =   3195
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Valor Débito Corrigido:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   120
      TabIndex        =   18
      Top             =   2865
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "CPF / CNPJ.(Só número):"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   120
      TabIndex        =   17
      Top             =   2535
      Width           =   2505
   End
   Begin VB.Label lblCPF 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   16
      Top             =   2535
      Width           =   2595
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço p/Correspond.:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   2205
      Width           =   2505
   End
   Begin VB.Label lblEndCor 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   14
      Top             =   2205
      Width           =   6165
   End
   Begin VB.Label Label1 
      Caption         =   "Exercício(s).:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   4500
      TabIndex        =   11
      Top             =   2865
      Width           =   1575
   End
   Begin VB.Label lblTipoDoc 
      Caption         =   "Processo nº...........:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   615
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Requerente (Cidadão)..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   1890
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Contribuinte..:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   1260
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Código/Inscrição......:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   945
      Width           =   2505
   End
   Begin VB.Label Label1 
      Caption         =   "Endereço do Imóvel....:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1575
      Width           =   2505
   End
   Begin VB.Label lblAno 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6120
      TabIndex        =   5
      Top             =   2865
      Width           =   1575
   End
   Begin VB.Label lblEnd 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   4
      Top             =   1575
      Width           =   6165
   End
   Begin VB.Label lblProp 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   3
      Top             =   1260
      Width           =   6165
   End
   Begin VB.Label lblRequerente 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2700
      TabIndex        =   2
      Top             =   1890
      Width           =   5655
   End
End
Attribute VB_Name = "frmConfissaoDivida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim xImovel As clsImovel, bDV As Boolean

Private Sub cmdCnsImovel_Click()

Set frm = frmCnsCidadao
frm.sForm = "frmConfissaoDivida"
frm.show
frm.ZOrder 0

End Sub

Private Sub cmdPrint_Click()

If Opt(0).value = True Then
    If Val(lblCod.Caption) = 0 Then
        MsgBox "Código/Inscrição inválido.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If lblProp.Caption = "" Then
        MsgBox "Proprietário inválido.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If lblRequerente.Caption = "" Then
        MsgBox "Requerente inválido.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    If frmMdi.frTeste.Visible = True Then
        frmReport.ShowReport "CONFDIVIDATMP", frmMdi.HWND, Me.HWND
    Else
        frmReport.ShowReport "CONFDIVIDA", frmMdi.HWND, Me.HWND
    End If
Else
    If lblProp.Caption = "" Then
        MsgBox "Selecione um documento.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If Trim(txtRequerente.Text) = "" Then
        MsgBox "Digite o nome do requerente.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If ValidaCPF(Trim(txtCPF.Text)) = 0 And ValidaCGC(Trim(txtCPF.Text)) = 0 Then
        MsgBox "CPF/CNPJ inválido.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If Not IsDate(mskVenc.Text) Then
        MsgBox "Data de vencimento inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
    frmReport.ShowReport2 "CONFDIVIDADAM", frmMdi.HWND, Me.HWND
End If

End Sub

Private Sub cmdSair_Click()
Dim bBoleto As Boolean

'If bFichaCompensacao Then GoTo fim

bBoleto = False
If txtNumProc.Locked = True Then
    If lblSid.Caption = "0" Then
        'EXIBE RELATORIO
'        If frmMdi.frTeste.Visible = False Then
'            frmReport.ShowReport "Carne", frmMdi.hwnd, Me.hwnd
'        Else
'            frmReport.ShowReport "CarneTmp", frmMdi.hwnd, Me.hwnd
'        End If
        'LIMPA TEMPORARIO
'        Sql = "DELETE FROM CARNETMP WHERE COMPUTER='" & NomeDoUsuario & "'"
'        cn.Execute Sql, rdExecDirect
    Else
        If bBoleto Then
            If frmMdi.frTeste.Visible = False Then
                frmReport.ShowReport2 "BOLETOGUIA", frmMdi.HWND, Me.HWND, lblSid.Caption
            Else
                frmReport.ShowReport2 "BOLETOGUIATMP", frmMdi.HWND, Me.HWND, lblSid.Caption
            End If
        Else
            If frmMdi.frTeste.Visible = False Then
'                If bFichaCompensacao Then
                    frmReport.ShowReport2 "BOLETOGUIA_V5", frmMdi.HWND, Me.HWND, lblSid.Caption
'                Else
'                    frmReport.ShowReport2 "BOLETOGUIA_V4", frmMdi.HWND, Me.HWND, lblSid.Caption
'                End If
            Else
                frmReport.ShowReport2 "BOLETOGUIA_V4TMP", frmMdi.HWND, Me.HWND, lblSid.Caption
            End If
        End If
        Sql = "delete from boletoguiacapa where sid=" & lblSid.Caption
        cn.Execute Sql, rdExecDirect
        
        Sql = "delete from boletoguia where sid=" & lblSid.Caption
        cn.Execute Sql, rdExecDirect
    End If
End If

fim:
Unload Me

End Sub


Private Sub Form_Load()
Set xImovel = New clsImovel

txtNumDoc.Visible = False
lblTipoDoc(1).Visible = False
txtRequerente.Visible = False
txtCPF.Visible = False
mskVenc.Visible = False

Centraliza Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload frmCnsCidadao
Set xImovel = Nothing
End Sub

Private Sub Opt_Click(Index As Integer)
Limpa
If Index = 0 Then
    txtNumProc.Visible = True
    lblTipoDoc(0).Visible = True
    txtNumDoc.Visible = False
    lblTipoDoc(1).Visible = False
    txtRequerente.Visible = False
    lblRequerente.Visible = True
    txtCPF.Visible = False
    lblCPF.Visible = True
    mskVenc.Visible = False
    lblVenc.Visible = True
    
    txtNumProc.SetFocus
Else
    txtNumProc.Visible = False
    lblTipoDoc(0).Visible = False
    txtNumDoc.Visible = True
    lblTipoDoc(1).Visible = True
    txtRequerente.Visible = True
    lblRequerente.Visible = False
    txtCPF.Visible = True
    lblCPF.Visible = False
    mskVenc.Visible = True
    lblVenc.Visible = False
    
    txtNumDoc.SetFocus
End If

End Sub


Private Sub txtCPF_KeyPress(KeyAscii As Integer)
Tweak txtCPF, KeyAscii, IntegerPositive
End Sub

Private Sub txtNumDoc_Change()
If lblProp.Caption <> "" Then
    Limpa
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
Dim nNumdoc As Long, nCodReduz As Long, RdoAux2 As rdoResultset, sNome As String, sEnd As String, sDoc As String
Dim qd As New rdoQuery, sExercicio As String, aExercicio() As Integer, x As Integer, bFind As Boolean
Dim sDataVencto As String, nValorAnistia As Double, nValorTributo As Double, nValorJuros As Double, nValorMulta As Double
Dim nValorCorrecao As Double, nValorTotal As Double

Set qd.ActiveConnection = cn
nNumdoc = Val(txtNumDoc.Text)
If nNumdoc = 0 Then Exit Sub

ReDim aExercicio(0)

Sql = "select numdocumento from numdocumento where numdocumento=" & nNumdoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        RdoAux.Close
        MsgBox "Documento não localizado.", vbExclamation, "Atenção"
        Exit Sub
    End If
   .Close
End With

Sql = "SELECT numdocumento.numdocumento, numdocumento.datadocumento,isentomj,percisencao, parceladocumento.codreduzido, parceladocumento.anoexercicio, parceladocumento.codlancamento,"
Sql = Sql & "parceladocumento.seqlancamento, parceladocumento.numparcela, parceladocumento.codcomplemento "
Sql = Sql & "FROM numdocumento INNER JOIN parceladocumento ON numdocumento.numdocumento = parceladocumento.numdocumento WHERE numdocumento.numdocumento=" & nNumdoc
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nCodReduz = !CODREDUZIDO
    bFind = False
    For x = 1 To UBound(aExercicio)
        If aExercicio(x) = !AnoExercicio Then
            bFind = True
            Exit For
        End If
    Next
    If Not bFind Then
        ReDim Preserve aExercicio(UBound(aExercicio) + 1)
        aExercicio(UBound(aExercicio)) = !AnoExercicio
    End If
    
    If nCodReduz < 100000 Then
        Sql = "select logradouro,li_num,nomecidadao,cpf,cnpj,descbairro,desccidade,li_uf from vwfullimovel2 where codreduzido=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = SubNull(RdoAux2!nomecidadao)
        sEnd = SubNull(RdoAux2!Logradouro) & ", " & SubNull(RdoAux2!Li_Num) & " " & SubNull(RdoAux2!DescBairro) & " " & SubNull(RdoAux2!descCidade) & "/" & SubNull(RdoAux2!li_uf)
        If Not IsNull(RdoAux2!Cnpj) Then
            sDoc = RdoAux2!Cnpj
        End If
        If sDoc = "" Then
            If Not IsNull(RdoAux2!CPF) Then
                sDoc = RdoAux2!CPF
            End If
        Else
            sDoc = ""
        End If
        RdoAux2.Close
    ElseIf nCodReduz >= 100000 And nCodReduz <= 500000 Then
        Sql = "SELECT razaosocial, LOGRADOURO, numero, descbairro, desccidade, siglauf, cpf, cnpj FROM vwFULLEMPRESA3 where codigomob=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = SubNull(RdoAux2!razaosocial)
        sEnd = SubNull(RdoAux2!Logradouro) & ", " & SubNull(RdoAux2!Numero) & " " & SubNull(RdoAux2!DescBairro) & " " & SubNull(RdoAux2!descCidade) & "/" & SubNull(RdoAux2!SiglaUF)
        If Not IsNull(RdoAux2!Cnpj) Then
            sDoc = Format(RdoAux2!Cnpj, "00\.000\.000/0000-00")
        Else
            If Not IsNull(RdoAux2!CPF) Then
                sDoc = Format(RdoAux2!CPF, "000\.000\.000-00")
            Else
                sDoc = ""
            End If
        End If
        RdoAux2.Close
    Else
        Sql = "SELECT nomecidadao, cpf, cnpj, ENDERECO, numimovel, DESCCIDADE, siglauf, descbairro FROM vwFULLCIDADAO where codcidadao=" & nCodReduz
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = SubNull(RdoAux2!nomecidadao)
        sEnd = SubNull(RdoAux2!Endereco) & ", " & SubNull(RdoAux2!NUMIMOVEL) & " " & SubNull(RdoAux2!DescBairro) & " " & SubNull(RdoAux2!descCidade) & " " & SubNull(RdoAux2!descCidade) & "/" & SubNull(RdoAux2!SiglaUF)
        If Not IsNull(RdoAux2!Cnpj) Then
            sDoc = Format(RdoAux2!Cnpj, "00\.000\.000/0000-00")
        Else
            If Not IsNull(RdoAux2!CPF) Then
                sDoc = Format(RdoAux2!CPF, "000\.000\.000-00")
            Else
                sDoc = ""
            End If
        End If
        RdoAux2.Close
    End If
    
    lblCod.Caption = Format(nCodReduz, "000000")
    lblProp.Caption = sNome
    txtRequerente.Text = sNome
    lblEnd.Caption = sEnd
    txtCPF.Text = sDoc
    mskVenc.Text = Format(!Datadocumento, "dd/mm/yyyy")
    
    Do Until .EOF
        On Error Resume Next
        RdoAux2.Close
        On Error GoTo 0
        qd.Sql = "{ Call spEXTRATONEW(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) }"
        qd(0) = nCodReduz
        qd(1) = nCodReduz
        qd(2) = !AnoExercicio
        qd(3) = !AnoExercicio
        qd(4) = !CodLancamento
        qd(5) = !CodLancamento
        qd(6) = !SeqLancamento
        qd(7) = !SeqLancamento
        qd(8) = !NumParcela
        qd(9) = !NumParcela
        qd(10) = !CODCOMPLEMENTO
        qd(11) = !CODCOMPLEMENTO
        qd(12) = 1
        qd(13) = 99
        qd(14) = Format(!Datadocumento, "mm/dd/yyyy")
        qd(15) = NomeDoUsuario
        Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
        With RdoAux2
            Do Until .EOF
            
                nValorTributo = !ValorTributo
                nValorMulta = !ValorMulta
                nValorJuros = !ValorJuros
                nValorCorrecao = !ValorCorrecao
                
                sDataVencto = Format(RdoAux!Datadocumento, "dd/mm/yyyy")
                If CDate(sDataVencto) >= CDate("01/10/2013") And CDate(sDataVencto) <= CDate("31/10/2013") Then
                    If Val(SubNull(RdoAux!isentomj)) = 0 Then
                        nValorAnistia = 0
                    Else
                        nValorAnistia = 100
                    End If
                ElseIf CDate(sDataVencto) >= CDate("01/11/2013") And CDate(sDataVencto) <= CDate("29/11/2013") Then
                    If Val(SubNull(RdoAux!isentomj)) = 0 Then
                        nValorAnistia = 0
                    Else
                        nValorAnistia = 95
                    End If
                ElseIf CDate(sDataVencto) >= CDate("30/11/2013") And CDate(sDataVencto) <= CDate("27/12/2013") Then
                    If Val(SubNull(RdoAux!isentomj)) = 0 Then
                        nValorAnistia = 0
                    Else
                        nValorAnistia = 90
                    End If
                ElseIf CDate(sDataVencto) < CDate("01/10/2013") Or CDate(sDataVencto) >= CDate("28/12/2013") Then
                   nValorAnistia = 0
                End If
                nPerc = 100 - nValorAnistia
                nValorMulta = nValorMulta * nPerc / 100
                nValorJuros = nValorJuros * nPerc / 100
                nValorTotal = nValorTotal + nValorTributo + nValorMulta + nValorJuros + nValorCorrecao
                
               .MoveNext
            Loop
           .Close
        End With
       .MoveNext
    Loop
    lblValor.Caption = FormatNumber(nValorTotal, 2)
    For x = 1 To UBound(aExercicio)
        sExercicio = sExercicio & aExercicio(x) & ", "
    Next
    If Len(sExercicio) > 0 Then
        sExercicio = Left(sExercicio, Len(sExercicio) - 1)
    End If
    lblAno.Caption = sExercicio
End With



End Sub

Private Sub txtNumProc_Change()

If lblProp.Caption <> "" Then
    Limpa
End If

End Sub

Private Sub txtNumProc_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    txtNumProc_LostFocus
End If

End Sub

Private Sub txtNumProc_LostFocus()
'On Error Resume Next
Dim nCodCidadao As Long

If Trim(txtNumProc.Text) = "" Then Exit Sub
sValidaProc = ValidaProcesso(txtNumProc.Text)
   
    
If NovoProtocolo = 0 Then
    Sql = "SELECT CODCIDAPRO FROM PROCESSO WHERE ANOPROCESS=" & Val(Right$(txtNumProc.Text, 4)) & " AND NUMEROPROC=" & Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            MsgBox "Cidadão não localizado no protocolo.", vbExclamation, "Atenção"
            Exit Sub
        Else
            bDV = False
            nCodCidadao = !CODCIDAPRO
        End If
       .Close
    End With
Else
    Sql = "SELECT CODCIDADAO,DATAENTRADA FROM PROCESSOGTI WHERE ANO=" & Val(Right$(txtNumProc.Text, 4)) & " AND NUMERO=" & Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            
            GoTo REMOVEDIGITO
    '        MsgBox "Cidadão não localizado no protocolo.", vbExclamation, "Atenção"
    '        Exit Sub
        Else
            bDV = False
            nCodCidadao = !CodCidadao
            lblDataProc.Caption = Format(!DATAENTRADA, "dd/mm/yyyy")
        End If
       .Close
    End With
End If
GoTo continua
REMOVEDIGITO:
    Sql = "SELECT CODCIDADAO,DATAENTRADA FROM PROCESSOGTI WHERE ANO=" & Val(Right$(txtNumProc.Text, 4)) & " AND NUMERO=" & Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 6))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
            lblDataProc.Caption = ""
            MsgBox "Cidadão não localizado no protocolo.", vbExclamation, "Atenção"
            Exit Sub
        Else
            bDV = True
            nCodCidadao = !CodCidadao
            lblDataProc.Caption = Format(!DATAENTRADA, "dd/mm/yyyy")
        End If
       .Close
    End With
    
    
    On Error Resume Next
'    Sql = "SELECT CODCIDAPRO FROM PROCESSO WHERE ANOPROCESS=" & Val(Right$(txtNumProc.text, 4)) & " AND NUMEROPROC=" & Val(Left$(txtNumProc.text, Len(txtNumProc.text) - 5))
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
        
'        nCodCidadao = !CODCIDAPRO
'       .Close
'    End With
continua:
    If nCodCidadao = 0 Then GoTo fim
 '   Sql = "SELECT cidadao.codcidadao, cidadao.nomecidadao, cidadao.cpf, cidadao.cnpj, cidadao.codlogradouro, vwLOGRADOURO.ABREVTIPOLOG, "
 '   Sql = Sql & "vwLOGRADOURO.ABREVTITLOG, vwLOGRADOURO.NOMELOGRADOURO, cidadao.numimovel, cidadao.complemento,"
 '   Sql = Sql & "cidadao.nomelogradouro AS nomerua FROM  cidadao LEFT OUTER JOIN  vwLOGRADOURO ON cidadao.codlogradouro = vwLOGRADOURO.CODLOGRADOURO "
    Sql = "select * from vwfullcidadao WHERE CODCIDADAO=" & nCodCidadao
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        lblRequerente.Caption = SubNull(!nomecidadao)
        If Not IsNull(!CPF) And SubNull(!CPF) <> "" Then
            
            lblCPF.Caption = !CPF
        Else
            If Not IsNull(!Cnpj) Then
                lblCPF.Caption = Format(RdoAux!Cnpj, "00\.000\.000/0000-00")
            Else
                lblCPF.Caption = ""
            End If
        End If
        If Val(SubNull(!CodLogradouro)) > 0 Then
            lblEndCor.Caption = Trim$(SubNull(!Endereco)) & ", " & Val(SubNull(!NUMIMOVEL)) & " " & SubNull(!Complemento)
            'lblEndCor.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " Nº " & !NUMIMOVEL & " " & SubNull(!Complemento)
        Else
            lblEndCor.Caption = Trim$(SubNull(!Enderecoc)) & ", " & Val(SubNull(!NUMIMOVEL2)) & " " & SubNull(!Complemento2)
       '     lblEndCor.Caption = Trim$(SubNull(!nomerua)) & " Nº " & !NUMIMOVEL & " " & SubNull(!Complemento)
        End If
       .Close
    End With
'End If
fim:
CarregaProcesso

End Sub

Public Sub CarregaProcesso()
Dim nCodReduz As Long, nSoma As Double, RdoAux2 As rdoResultset, nNumproc As Long, nAno As Integer, sNumProc As String, nCont As Integer


nCont = 0
Inicio:
nSoma = 0
If bDV Then
    Sql = "SELECT * FROM vwPROCESSOPARCELA WHERE NUMPROCESSO='" & txtNumProc.Text & "' and codtributo>0"
Else
    nNumproc = Val(Left$(txtNumProc.Text, Len(txtNumProc.Text) - 5))
    sNumProc = CStr(nNumproc) & RetornaDVProcesso(nNumproc)
    nAno = Val(Right$(txtNumProc.Text, 4))
    sNumProc = sNumProc & "/" & CStr(nAno)
    Sql = "SELECT * FROM vwPROCESSOPARCELA WHERE NUMPROCESSO='" & sNumProc & "' and codtributo>0"
End If
nCont = nCont + 1

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        bDV = Not bDV
        If nCont < 2 Then GoTo Inicio
    End If
    If .RowCount > 0 Then
        nCodReduz = !CODIGORESP
        lblCod.Caption = Format(nCodReduz, "000000")
        Sql = "SELECT DISTINCT ANOEXERCICIO From ORIGEMREPARC WHERE NUMPROCESSO = '" & txtNumProc.Text & "' ORDER BY ANOEXERCICIO"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            lblAno.Caption = ""
            Do Until .EOF
                 lblAno.Caption = lblAno.Caption & !AnoExercicio & ", "
                .MoveNext
            Loop
           .Close
        End With
        lblAno.Caption = Chomp(lblAno.Caption, chomp_righT, 1)
'        Sql = "SELECT MIN(ANOEXERCICIO) AS MINIMO, MAX(ANOEXERCICIO) AS MAXIMO "
'        Sql = Sql & "From ORIGEMREPARC WHERE NUMPROCESSO = '" & txtNumProc.text & "'"
'        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        With RdoAux2
'            lblAno.Caption = !MINIMO & " - " & !MAXIMO
'           .Close
'        End With
        lblQtdeParc.Caption = !qtdeparcela
        Do Until .EOF
           If !NumParcela = 1 Then
              lblVenc.Caption = Format(!DataVencimento, "dd/mm/yyyy")
           End If
           nSoma = nSoma + !ValorTributo
          .MoveNext
        Loop
        lblValor.Caption = FormatNumber(nSoma, 2)
        
        
        If nCodReduz < 100000 Then
            Set xImovel = New clsImovel
            With xImovel
               .CarregaImovel nCodReduz
                lblProp.Caption = .NomePropPrincipal
                lblEnd.Caption = .EnderecoCompleto
            End With
        ElseIf nCodReduz > 100000 And nCodReduz < 500000 Then
            Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodReduz
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    lblProp.Caption = !razaosocial
                    lblEnd.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro & " nº " & SubNull(!Numero)
                End If
            End With
        ElseIf nCodReduz > 500000 Then
            Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If .RowCount > 0 Then
                    lblProp.Caption = !nomecidadao
                End If
                
            End With
        
        End If
        
    Else
        MsgBox "Processo de Reparcelamento não encontrado.", vbExclamation, "Atenção"
    End If
   .Close
End With

End Sub

Private Sub Limpa()

lblCod.Caption = ""
lblProp.Caption = ""
lblEnd.Caption = ""
lblRequerente.Caption = ""
lblEndCor.Caption = ""
lblCPF.Caption = ""
lblValor.Caption = ""
lblVenc.Caption = ""
lblAno.Caption = ""
lblQtdeParc.Caption = ""

End Sub
