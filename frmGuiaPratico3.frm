VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGuiaPratico3 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimento p/autorização especial de estacionamento"
   ClientHeight    =   5070
   ClientLeft      =   4290
   ClientTop       =   3660
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   9420
   Begin VB.CheckBox chK 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Morador, decretos Municipais: 5192 de 18 de novembro de 2008 e 5299 de junho de 2009."
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   3105
      Width           =   7260
   End
   Begin VB.CheckBox chK 
      BackColor       =   &H00EEEEEE&
      Caption         =   "304 deficiente, de 18 de dezembro de 2008 do CONTRAM;"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   2835
      Width           =   7260
   End
   Begin VB.CheckBox chK 
      BackColor       =   &H00EEEEEE&
      Caption         =   "303 idoso, de 18 de dezembro de 2008 do CONTRAM;"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   2565
      Width           =   7260
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Observação"
      Height          =   960
      Index           =   2
      Left            =   135
      TabIndex        =   33
      Top             =   3600
      Width           =   9015
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   645
         Left            =   45
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   225
         Width           =   8880
      End
   End
   Begin VB.TextBox txtPlaca 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6480
      TabIndex        =   5
      Top             =   1935
      Width           =   1455
   End
   Begin VB.TextBox txtRenavam 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3870
      TabIndex        =   4
      Top             =   1980
      Width           =   1455
   End
   Begin VB.TextBox txtCor 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   900
      TabIndex        =   3
      Top             =   1980
      Width           =   1590
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8055
      TabIndex        =   2
      Top             =   1485
      Width           =   870
   End
   Begin VB.TextBox txtModelo 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   5175
      TabIndex        =   1
      Top             =   1485
      Width           =   1905
   End
   Begin VB.TextBox txtMarca 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1935
      TabIndex        =   0
      Top             =   1485
      Width           =   1950
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   315
      Left            =   8820
      TabIndex        =   11
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
      MICON           =   "frmGuiaPratico3.frx":0000
      PICN            =   "frmGuiaPratico3.frx":001C
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
      Left            =   8235
      TabIndex        =   10
      ToolTipText     =   "Imprimir Requerimento"
      Top             =   4635
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
      MICON           =   "frmGuiaPratico3.frx":0176
      PICN            =   "frmGuiaPratico3.frx":0192
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
      Caption         =   "Placa..:"
      Enabled         =   0   'False
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
      Index           =   12
      Left            =   5535
      TabIndex        =   32
      Top             =   2025
      Width           =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "RENAVAM..:"
      Enabled         =   0   'False
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
      Index           =   11
      Left            =   2700
      TabIndex        =   31
      Top             =   2025
      Width           =   1110
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cor..:"
      Enabled         =   0   'False
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
      Left            =   180
      TabIndex        =   30
      Top             =   2025
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Enabled         =   0   'False
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
      Left            =   7290
      TabIndex        =   29
      Top             =   1530
      Width           =   660
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo..:"
      Enabled         =   0   'False
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
      Left            =   4095
      TabIndex        =   28
      Top             =   1530
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Veículo-Marca..:"
      Enabled         =   0   'False
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
      Left            =   180
      TabIndex        =   27
      Top             =   1530
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   9270
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label lblNum 
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
      Left            =   7830
      TabIndex        =   26
      Top             =   765
      Width           =   1020
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº..:"
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
      Left            =   7155
      TabIndex        =   25
      Top             =   765
      Width           =   570
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
      Left            =   1395
      TabIndex        =   24
      Top             =   765
      Width           =   5520
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
      Left            =   180
      TabIndex        =   23
      Top             =   765
      Width           =   1200
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
      Left            =   6165
      TabIndex        =   22
      Top             =   450
      Width           =   3630
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
      Left            =   4995
      TabIndex        =   21
      Top             =   450
      Width           =   1110
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
      Left            =   1395
      TabIndex        =   20
      Top             =   450
      Width           =   3405
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
      Left            =   180
      TabIndex        =   19
      Top             =   450
      Width           =   1200
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
      Left            =   2700
      TabIndex        =   18
      Top             =   135
      Width           =   705
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
      Left            =   180
      TabIndex        =   17
      Top             =   135
      Width           =   2505
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
      Left            =   3465
      TabIndex        =   16
      Top             =   135
      Width           =   5385
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro....:"
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
      Left            =   180
      TabIndex        =   15
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label lblBairro 
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
      Left            =   1395
      TabIndex        =   14
      Top             =   1080
      Width           =   3540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone..:"
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
      Left            =   4995
      TabIndex        =   13
      Top             =   1080
      Width           =   1200
   End
   Begin VB.Label lblFone 
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
      Left            =   6345
      TabIndex        =   12
      Top             =   1080
      Width           =   2685
   End
End
Attribute VB_Name = "frmGuiaPratico3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrint_Click()
Dim Sql As String

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "INSERT REPORTTMP(USUARIO,MEMO1) VALUES('" & NomeDeLogin & "','" & Mask(txtObs.Text) & "')"
cn.Execute Sql, rdExecDirect

frmReport.ShowReport2 "GUIAPRATICO3", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
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
        lblEndereco.Caption = SubNull(!Endereco)
        lblNum.Caption = SubNull(!NUMIMOVEL)
        If Not IsNull(!DescBairro) Then
            lblBairro.Caption = !DescBairro
'        Else
'            If Not IsNull(!NOMEBairro) Then
'                lblBairro.Caption = !NOMEBairro
'            End If
        End If
        
  '      If Not IsNull(!NomeCidade) Then
 '           lblCidade.Caption = !NomeCidade
    '    Else
   '         If Not IsNull(!DESCCidade) Then
'                lblCidade.Caption = !DESCCidade
            'End If
     '   End If
        lblFone.Caption = SubNull(!TELEFONE)
       .MoveNext
    Loop
   .Close
End With

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


End Sub


