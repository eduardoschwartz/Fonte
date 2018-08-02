VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGare 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guia de Arrecadação - GARE"
   ClientHeight    =   2820
   ClientLeft      =   2325
   ClientTop       =   2370
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2820
   ScaleWidth      =   9450
   Begin VB.TextBox txtExecutado 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   6090
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7920
      MaxLength       =   15
      TabIndex        =   5
      Top             =   1800
      Width           =   1410
   End
   Begin VB.TextBox txtVencto 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5175
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtNumExec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1800
      Width           =   1410
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   6
      TabIndex        =   1
      Top             =   1440
      Width           =   1005
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   8820
      TabIndex        =   0
      ToolTipText     =   "Consulta Cidadão"
      Top             =   45
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
      MICON           =   "frmGare.frx":0000
      PICN            =   "frmGare.frx":001C
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
      Height          =   315
      Left            =   7200
      TabIndex        =   6
      ToolTipText     =   "Imprimir Requerimento"
      Top             =   2385
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
      MICON           =   "frmGare.frx":0176
      PICN            =   "frmGare.frx":0192
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
      Height          =   315
      Left            =   8280
      TabIndex        =   7
      ToolTipText     =   "Sair da Tela"
      Top             =   2385
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmGare.frx":02EC
      PICN            =   "frmGare.frx":0308
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
      Caption         =   "Valor Total:"
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
      Left            =   6570
      TabIndex        =   24
      Top             =   1845
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Vencto:"
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
      Left            =   3870
      TabIndex        =   23
      Top             =   1845
      Width           =   1290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No Execução Fiscal:"
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
      Left            =   90
      TabIndex        =   22
      Top             =   1845
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código executado..:"
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
      Left            =   90
      TabIndex        =   21
      Top             =   1485
      Width           =   2055
   End
   Begin VB.Label lblUF 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   7470
      TabIndex        =   20
      Top             =   1080
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sigla UF..:"
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
      Left            =   6210
      TabIndex        =   19
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label lblCidade 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1350
      TabIndex        =   18
      Top             =   1080
      Width           =   4755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Município..:"
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
      Left            =   90
      TabIndex        =   17
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label lblRequerente 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   3420
      TabIndex        =   16
      Top             =   135
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Left            =   90
      TabIndex        =   15
      Top             =   135
      Width           =   2505
   End
   Begin VB.Label lblCodRequerente 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   2655
      TabIndex        =   14
      Top             =   135
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº DE RG..:"
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
      Left            =   90
      TabIndex        =   13
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label lblRG 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1350
      TabIndex        =   12
      Top             =   450
      Width           =   3405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ.:"
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
      Left            =   4950
      TabIndex        =   11
      Top             =   450
      Width           =   1110
   End
   Begin VB.Label lblCPF 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   6120
      TabIndex        =   10
      Top             =   450
      Width           =   3630
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço..:"
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
      Left            =   90
      TabIndex        =   9
      Top             =   765
      Width           =   1200
   End
   Begin VB.Label lblEndereco 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   225
      Left            =   1350
      TabIndex        =   8
      Top             =   765
      Width           =   8400
   End
End
Attribute VB_Name = "frmGare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
frmReport.ShowReport "GARE", frmMdi.hwnd, Me.hwnd
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
txtVencto.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Limpa()
lblCodRequerente.Caption = ""
lblRequerente.Caption = ""
lblRG.Caption = ""
lblCPF.Caption = ""
lblEndereco.Caption = ""
End Sub

Private Sub cmdCnsImovel_Click()
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frm.show
frm.ZOrder 0
End Sub

Private Sub Form_Activate()
If Val(lblCodRequerente.Caption) = 0 Then Exit Sub
If Val(lblCodRequerente.Caption) > 500000 Then
    Le
Else
    MsgBox "Código de cidadão inválido.", vbExclamation, "Atenção"
    Limpa
End If

End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset
Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & Val(lblCodRequerente.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
        lblRG.Caption = SubNull(!rg) & " " & SubNull(!ORGAO)
        lblCPF.Caption = SubNull(!CPF)
        If lblCPF.Caption <> "" Then
            lblCPF.Caption = Format(RdoAux!CPF, "00#\.###\.###-##")
        End If
        If lblCPF.Caption = "" Then
            If Not IsNull(!Cnpj) Then
                lblCPF.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
            End If
        End If
        lblEndereco.Caption = SubNull(!Endereco) & " ," & SubNull(!NUMIMOVEL)
        If Not IsNull(!DescBairro) And SubNull(!DescBairro) <> "" Then
            lblEndereco.Caption = lblEndereco.Caption & " - " & !DescBairro
     '   Else
     '       If Not IsNull(!NOMEBairro) Then
     '           lblEndereco.Caption = lblEndereco.Caption & " - " & !NOMEBairro
     '       End If
        End If
        If Not IsNull(!Complemento) Or !Complemento <> "" Then
            lblEndereco.Caption = lblEndereco.Caption & " " & !Complemento
        End If
        
'        If Not IsNull(!NomeCidade) Then
'            lblCidade.Caption = !NomeCidade
'            lblUF.Caption = SubNull(!NOMEUF)
'        Else
            If Not IsNull(!descCidade) Then
                lblCidade.Caption = !descCidade
                lblUF.Caption = SubNull(!siglaUF)
  '          Else
   '             lblCidade.Caption = !DEscCidade
    '            lblUF.Caption = SubNull(!siglaUF)
            
            End If
 '       End If
       
       .MoveNext
    Loop
   .Close
End With
CodCidadao = 0
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    txtCod_LostFocus
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If
End Sub

Private Sub txtCod_LostFocus()
Dim Sql As String, RdoAux As rdoResultset, sNome As String, nCodReduz As Long

nCodReduz = Val(txtCod.Text)
txtExecutado.Text = "": sNome = ""
If Val(nCodReduz) > 0 And Val(nCodReduz) < 100000 Then
    Sql = "SELECT NOMECIDADAO FROM vwFULLIMOVEL2 WHERE CODREDUZIDO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            sNome = !nomecidadao
        Else
            MsgBox "Código não cadastrado", vbCritical, "Atenção"
        End If
       .Close
    End With
ElseIf Val(nCodReduz) >= 100000 And Val(nCodReduz) < 500000 Then
    Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            sNome = !razaosocial
        Else
            MsgBox "Código não cadastrado", vbCritical, "Atenção"
        End If
       .Close
    End With
ElseIf Val(nCodReduz) >= 500000 And Val(nCodReduz) < 700000 Then
    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            sNome = !nomecidadao
        Else
            MsgBox "Código não cadastrado", vbCritical, "Atenção"
        End If
       .Close
    End With
End If

txtExecutado.Text = sNome

End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
End Sub
