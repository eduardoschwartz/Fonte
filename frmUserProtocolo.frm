VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmUsuario 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   6075
   ClientLeft      =   4320
   ClientTop       =   1605
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   5625
   Begin VB.Frame pn3 
      BackColor       =   &H00EEEEEE&
      Height          =   5325
      Left            =   6615
      TabIndex        =   8
      Top             =   45
      Width           =   5625
      Begin VB.ListBox lstGrupo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4755
         ItemData        =   "frmUserProtocolo.frx":0000
         Left            =   150
         List            =   "frmUserProtocolo.frx":0022
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   450
         Width           =   5295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Grupos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   17
         Top             =   180
         Width           =   2535
      End
   End
   Begin VB.Frame pn4 
      BackColor       =   &H00EEEEEE&
      Height          =   5325
      Left            =   6255
      TabIndex        =   24
      Top             =   45
      Width           =   5625
      Begin VB.ListBox lstSetor2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4755
         Left            =   90
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   450
         Width           =   5385
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Centro de Custos (Compras):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   26
         Top             =   150
         Width           =   3705
      End
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4470
      TabIndex        =   23
      ToolTipText     =   "Sair da Tela"
      Top             =   5685
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmUserProtocolo.frx":00A0
      PICN            =   "frmUserProtocolo.frx":00BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4470
      TabIndex        =   18
      ToolTipText     =   "Cancelar Edição"
      Top             =   5685
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmUserProtocolo.frx":012A
      PICN            =   "frmUserProtocolo.frx":0146
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   90
      TabIndex        =   19
      ToolTipText     =   "Novo Registro"
      Top             =   5685
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmUserProtocolo.frx":02A0
      PICN            =   "frmUserProtocolo.frx":02BC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1140
      TabIndex        =   20
      ToolTipText     =   "Editar Registro"
      Top             =   5685
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmUserProtocolo.frx":0416
      PICN            =   "frmUserProtocolo.frx":0432
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   2190
      TabIndex        =   21
      ToolTipText     =   "Excluir Registro"
      Top             =   5685
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmUserProtocolo.frx":058C
      PICN            =   "frmUserProtocolo.frx":05A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   3420
      TabIndex        =   22
      ToolTipText     =   "Gravar os Dados"
      Top             =   5685
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmUserProtocolo.frx":064A
      PICN            =   "frmUserProtocolo.frx":0666
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame pn2 
      BackColor       =   &H00EEEEEE&
      Height          =   5325
      Left            =   5730
      TabIndex        =   6
      Top             =   45
      Width           =   5625
      Begin VB.ListBox lstSetor 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4755
         Left            =   90
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   450
         Width           =   5385
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Centro de Custos:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   150
         Width           =   2535
      End
   End
   Begin VB.Frame pn1 
      BackColor       =   &H00EEEEEE&
      Height          =   5250
      Left            =   30
      TabIndex        =   7
      Top             =   330
      Width           =   5595
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1185
         TabIndex        =   28
         Top             =   4140
         Width           =   4215
      End
      Begin VB.CheckBox chkFiscal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Fiscal"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4680
         TabIndex        =   27
         Top             =   3420
         Width           =   750
      End
      Begin VB.ListBox lstUser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   2370
         Left            =   90
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   5325
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1185
         TabIndex        =   1
         Top             =   3060
         Width           =   4215
      End
      Begin VB.TextBox txtLogin 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1185
         TabIndex        =   2
         Top             =   3420
         Width           =   3270
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1185
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3780
         Width           =   1485
      End
      Begin VB.TextBox txtSenha2 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   3930
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3780
         Width           =   1485
      End
      Begin prjChameleon.chameleonButton cmdFoto 
         Height          =   270
         Left            =   585
         TabIndex        =   31
         ToolTipText     =   "Carregar imagem da assinatura"
         Top             =   4815
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "..."
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
         MICON           =   "frmUserProtocolo.frx":0A0B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCod 
         Caption         =   "0000"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4080
         TabIndex        =   34
         Top             =   4710
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "Id..:"
         Height          =   195
         Left            =   3720
         TabIndex        =   33
         Top             =   4710
         Width           =   375
      End
      Begin VB.Label lblAss 
         Caption         =   ".."
         Height          =   240
         Left            =   135
         TabIndex        =   32
         Top             =   4815
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Image Image1 
         Height          =   555
         Left            =   1170
         Stretch         =   -1  'True
         Top             =   4545
         Width           =   2085
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Assinatura..:"
         Height          =   255
         Index           =   3
         Left            =   135
         TabIndex        =   30
         Top             =   4545
         Width           =   945
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo.........:"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   29
         Top             =   4185
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Usuários:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         X1              =   5400
         X2              =   90
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome........:"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   3090
         Width           =   945
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Login.........:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   3465
         Width           =   945
      End
      Begin VB.Label lblSenha 
         BackStyle       =   0  'Transparent
         Caption         =   "Senha.......:"
         Height          =   255
         Left            =   150
         TabIndex        =   10
         Top             =   3825
         Width           =   945
      End
      Begin VB.Label lblSenha2 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmação.:"
         Height          =   255
         Left            =   2835
         TabIndex        =   9
         Top             =   3825
         Width           =   975
      End
   End
   Begin vbalDTab6.vbalDTabControl TabUser 
      Height          =   5565
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   9816
      AllowScroll     =   0   'False
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14737632
   End
End
Attribute VB_Name = "frmUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String
Dim Evento As String, sLogin As String, bAdmin As Boolean

Private Sub cmdAlterar_Click()
Evento = "Alterar"
Eventos "INCLUIR"
txtSenha.BackColor = Kde: txtSenha2.BackColor = Kde
txtSenha.Locked = True: txtSenha2.Locked = True
txtLogin.BackColor = Kde
txtLogin.Locked = True
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
Evento = ""
lstUser.ListIndex = 0
lstUser_Click
End Sub

Private Sub cmdExcluir_Click()

If lstUser.ListIndex = -1 Then Exit Sub

If MsgBox("Excluir o usuário " & lstUser.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub
sLogin = Trim$(txtLogin.Text)
'apaga o centro de custos
Sql = "DELETE FROM USUARIOCC WHERE NOME='" & sLogin & "'"
'cn.Execute Sql, rdExecDirect
'apaga o centro de custos
Sql = "DELETE FROM CPUSUARIOCC WHERE NOME='" & sLogin & "'"
'cn.Execute Sql, rdExecDirect
'remove acesso dos grupos
'For j = 0 To lstGrupo.ListCount - 1
'   Sql = "sp_droprolemember '" & lstGrupo.List(j) & "','" & sLogin & "'"
'   cn.Execute Sql, rdExecDirect
'Next
'apaga o usuário
Sql = "DELETE FROM USUARIO WHERE NOMELOGIN='" & sLogin & "'"
'cn.Execute Sql, rdExecDirect
'remove da tela
lstUser.RemoveItem (lstUser.ListIndex)
lstUser.ListIndex = 0: lstUser_Click

End Sub

Private Sub cmdFoto_Click()
Dim fName As String, cc As cCommonDlg

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Imagem da assinatura|*.jpg;*.png", , "C:\trabalho\gti\diversos", "Selecione uma assinatura", , Me.HWND, OFN_HIDEREADONLY, False
Image1.Picture = LoadPicture(fName)
lblAss.Caption = fName

End Sub

Private Sub cmdGravar_Click()
Dim j As Integer, bAchou As Boolean
Ocupado

txtNome.Text = UCase$(txtNome.Text)
txtLogin.Text = UCase$(txtLogin.Text)
txtSenha.Text = UCase$(txtSenha.Text)
txtSenha2.Text = UCase$(txtSenha2.Text)

If Trim$(txtNome.Text) = "" Then
    Liberado
    MsgBox "Digite o nome do usuário.", vbExclamation, "Atenção"
    txtNome.SetFocus
    Exit Sub
End If
If Trim$(txtLogin.Text) = "" Then
    Liberado
    MsgBox "Digite o nome de Login do usuário.", vbExclamation, "Atenção"
    txtLogin.SetFocus
    Exit Sub
End If
If Trim$(txtSenha.Text) = "" Or Trim$(txtSenha2.Text) = "" Then
    Liberado
    MsgBox "Digite a senha e confirmação de senha.", vbExclamation, "Atenção"
    txtSenha.SetFocus
    Exit Sub
End If
If Len(Trim$(txtSenha.Text)) < 6 Then
    Liberado
    MsgBox "Senha no mínimo 6 caracteres.", vbExclamation, "Atenção"
    Exit Sub
End If
If Trim$(txtSenha.Text) <> Trim$(txtSenha2.Text) Then
    Liberado
    MsgBox "Confirmação de senha não confere com a senha informada.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False
For j = 0 To lstSetor.ListCount - 1
    If lstSetor.Selected(j) = True Then
        bAchou = True
    End If
Next
If Not bAchou Then
   Liberado
   MsgBox "Selecione ao menos um centro de custos.", vbExclamation, "Atenção"
   Exit Sub
Else
    bAchou = False
    For j = 0 To lstGrupo.ListCount - 1
        If lstGrupo.Selected(j) = True Then
            bAchou = True
        End If
    Next
    If Not bAchou And bAdmin Then
       MsgBox "Selecione ao menos um grupo.", vbExclamation, "Atenção"
       Exit Sub
    End If
End If

If Not Grava() Then
    Liberado
    Exit Sub
End If
Liberado
Evento = ""
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
Evento = "Novo"
Eventos "INCLUIR"
Limpa
txtSenha.Enabled = True: txtSenha2.Enabled = True
txtSenha.BackColor = Branco: txtSenha2.BackColor = Branco
lblSenha.Enabled = True: lblSenha2.Enabled = True
lblSenha.Enabled = True: lblSenha2.Enabled = True
On Error Resume Next
txtNome.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim c As cTab, sUser As String

Centraliza Me

'sUser = Mid(frmMdi.Sbar.Panels(2).text, 10, Len(frmMdi.Sbar.Panels(2).text) - 8)
'If sUser = "SCHWARTZ" Or sUser = "SERGIO" Then
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "SERGIO" Then
    bAdmin = True
Else
    bAdmin = False
End If

CarregaUsuario
CarregaSetor
CarregaGrupo

Eventos "INICIAR"
With TabUser
    .ShowCloseButton = False
    Set c = .Tabs.Add("Tab1", , "Usuários")
    c.Panel = pn1
    Set c = .Tabs.Add("Tab2", , "Centro de Custos")
    c.Panel = pn2
    Set c = .Tabs.Add("Tab4", , "Centro de Custos (Compras)")
    c.Panel = pn4
    Set c = .Tabs.Add("Tab3", , "Grupos")
    c.Panel = pn3
End With

lstUser.ListIndex = 0

End Sub

Private Sub CarregaUsuario()
lstUser.Clear
'Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE SUBSTRING(USUARIO.NOMELOGIN,1,2)<>'F0' AND ATIVO=1"
Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE externo=0 AND ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstUser.AddItem !NomeCompleto
       .MoveNext
    Loop
End With

End Sub

Private Sub CarregaSetor()

Sql = "SELECT CODIGO,DESCRICAO FROM CENTROCUSTO WHERE ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstSetor.AddItem !descricao
        lstSetor.ItemData(lstSetor.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

'Sql = "SELECT CODIGO,DESCRICAO FROM CPCENTROCUSTO WHERE ATIVO=1"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
 '   Do Until .EOF
 '       lstSetor2.AddItem !descricao
 '       lstSetor2.ItemData(lstSetor2.NewIndex) = !Codigo
 '      .MoveNext
 '   Loop
 '  .Close
'End With

End Sub

Private Sub CarregaGrupo()

Sql = "SELECT * FROM VWGRUPOS"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If UCase$(Left$(!Name, 2)) <> "DB" Then
            lstGrupo.AddItem !Name
            lstGrupo.ItemData(lstGrupo.NewIndex) = !uID
        End If
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdFoto.Enabled = False
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmUsuario
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Locked = True
       End If
   Next
   chkFiscal.Enabled = False
   lstSetor.Enabled = False
   lblSenha.Enabled = False: lblSenha2.Enabled = False
   lstUser.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdFoto.Enabled = True
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmUsuario
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = vbWhite
          Ct.Locked = False
       End If
   Next
   chkFiscal.Enabled = True
   lstUser.Enabled = False
   lstSetor.Enabled = True
End If

If Not bAdmin Then
    cmdNovo.Enabled = False
    cmdExcluir.Enabled = False
    txtNome.Enabled = False
    txtSenha.Enabled = False
    txtSenha2.Enabled = False
    lstGrupo.Enabled = False
End If

End Sub

Private Sub lstUser_Click()
If lstUser.ListIndex > -1 Then
    Limpa
    Le
End If
End Sub

Private Sub Limpa()
Dim j As Integer
lblAss.Caption = ""

Set Image1.DataSource = Nothing
Image1.Picture = Nothing
lblCod.Caption = "0000"
txtNome.Text = ""
txtLogin = ""
txtSenha.Text = ""
txtSenha2.Text = ""
Text1.Text = ""
chkFiscal.value = vbUnchecked
For j = 0 To lstSetor.ListCount - 1
    lstSetor.Selected(j) = False
Next
For j = 0 To lstSetor2.ListCount - 1
    lstSetor2.Selected(j) = False
Next
For j = 0 To lstGrupo.ListCount - 1
    lstGrupo.Selected(j) = False
Next

End Sub

Private Sub Le()
Dim nCod As Integer, j As Integer

txtSenha.Text = "************"
txtSenha2.Text = "************"

Sql = "SELECT id,NOMELOGIN,FISCAL FROM USUARIO WHERE NOMECOMPLETO='" & Mask(lstUser.Text) & "' and ativo=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblCod.Caption = Format(!Id, "0000")
    txtNome.Text = lstUser.Text
    txtLogin.Text = !NomeLogin
    If Not IsNull(!fiscal) Then
        chkFiscal = IIf(!fiscal, 1, 0)
    End If
   .Close
End With

'Sql = "SELECT CODIGOCC FROM USUARIOCC WHERE NOME='" & txtLogin.Text & "'"
Sql = "SELECT CODIGOCC FROM USUARIOCC WHERE userid=" & RetornaUsuarioID(txtLogin.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nCod = !codigocc
        For j = 0 To lstSetor.ListCount - 1
            If lstSetor.ItemData(j) = nCod Then
                lstSetor.Selected(j) = True
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

'Sql = "SELECT CODIGOCC FROM CPUSUARIOCC WHERE NOME='" & txtLogin.Text & "'"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        nCod = !CODIGOCC
'        For j = 0 To lstSetor2.ListCount - 1
'            If lstSetor2.ItemData(j) = nCod Then
'                lstSetor2.Selected(j) = True
'            End If
'        Next
'       .MoveNext
'    Loop
'   .Close
'End With

Sql = "SELECT UID FROM VWUSUARIO WHERE NAME='" & txtLogin.Text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        nCod = !uID
    End If
   .Close
End With

'Sql = "SELECT * from usuario WHERE nomelogin='" & txtLogin.Text & "'"
Sql = "SELECT * from usuario WHERE id=" & RetornaUsuarioID(txtLogin.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For j = 0 To lstGrupo.ListCount - 1
            If lstGrupo.List(j) = !grupo Then
                lstGrupo.Selected(j) = True
            End If
        Next
       .MoveNext
    Loop
   .Close
End With


'le foto

Sql = "select * from assinatura where usuario='" & txtLogin.Text & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then

    Dim mStream As New ADODB.Stream
    Dim rst As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    
    adoConn.CursorLocation = adUseClient
    'adoConn.Open "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=schwartz;Password=kobudera;Initial Catalog=Tributacao;Data Source=" & IPServer
    adoConn.Open cn.Connect
    
    rst.Open "Select * from assinatura where usuario='" & txtLogin.Text & "'", adoConn, adOpenKeyset, adLockOptimistic
        
        
    With mStream
        .Type = adTypeBinary
        .Open
        If Not IsNull(rst("fotoass")) Then
            .Write rst("fotoass")
            Image1.DataField = "fotoass"
            Set Image1.DataSource = rst
        End If
    End With
    Set mStream = Nothing

End If
RdoAux.Close

End Sub

Private Function Grava() As Boolean
Dim j As Integer, qd As New rdoQuery, bAchou As Boolean, nCodigo As Integer
Grava = False
sLogin = UCase$(Trim$(txtLogin.Text))
Set qd.ActiveConnection = cn

If Evento = "Novo" Then
        Sql = "select max(id)as maximo from usuario"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nCodigo = RdoAux!maximo + 1
        RdoAux.Close
        lblCod.Caption = Format(nCodigo, "0000")
       'adiciona o usuário na tabela de usuários
        Sql = "INSERT USUARIO (id,NOMELOGIN,NOMECOMPLETO,ATIVO,FISCAL,SENHA,externo) VALUES(" & nCodigo & ",'"
        Sql = Sql & sLogin & "','" & Trim$(Mask(txtNome.Text)) & "',1," & IIf(chkFiscal.value = vbChecked, 1, 0) & ",'" & Encrypt128(Mask(txtSenha.Text), "everest") & "',0)"
        cn.Execute Sql, rdExecDirect
       'adiciona os centro de custos
        For j = 0 To lstSetor.ListCount - 1
            If lstSetor.Selected(j) = True Then
                Sql = "INSERT USUARIOCC (USERID,CODIGOCC) VALUES("
                Sql = Sql & RetornaUsuarioID(NomeDeLogin) & "," & lstSetor.ItemData(j) & ")"
                cn.Execute Sql, rdExecDirect
           End If
        Next
       'adiciona os centro de custos
'        For j = 0 To lstSetor2.ListCount - 1
'            If lstSetor2.Selected(j) = True Then
'                Sql = "INSERT CPUSUARIOCC (NOME,CODIGOCC) VALUES('"
 '               Sql = Sql & sLogin & "'," & lstSetor2.ItemData(j) & ")"
 '               cn.Execute Sql, rdExecDirect
 '          End If
 ''       Next
       'adiciona usuário a lista na tela
        lstUser.AddItem txtNome.Text
'       .Close
'    End With
Else
    nCodigo = Val(lblCod.Caption)
   'altera nome do usuário
    Sql = "UPDATE USUARIO SET NOMECOMPLETO='" & Trim$(txtNome.Text) & "',FISCAL=" & IIf(chkFiscal.value = vbChecked, 1, 0) & " WHERE "
    Sql = Sql & "NOMELOGIN='" & sLogin & "'"
    cn.Execute Sql, rdExecDirect
    'apaga o centro de custos
    Sql = "DELETE FROM USUARIOCC WHERE USERID=" & Val(lblCod.Caption)
    cn.Execute Sql, rdExecDirect
   'adiciona os centro de custos
    For j = 0 To lstSetor.ListCount - 1
        If lstSetor.Selected(j) = True Then
            Sql = "INSERT USUARIOCC (USERID,CODIGOCC) VALUES("
            Sql = Sql & Val(lblCod.Caption) & "," & lstSetor.ItemData(j) & ")"
            cn.Execute Sql, rdExecDirect
       End If
    Next
    
    'apaga o centro de custos
'    Sql = "DELETE FROM CPUSUARIOCC WHERE NOME='" & sLogin & "'"
'    cn.Execute Sql, rdExecDirect
   'adiciona os centro de custos
'    For j = 0 To lstSetor2.ListCount - 1
'        If lstSetor2.Selected(j) = True Then
'            Sql = "INSERT CPUSUARIOCC (NOME,CODIGOCC) VALUES('"
'            Sql = Sql & sLogin & "'," & lstSetor2.ItemData(j) & ")"
'            cn.Execute Sql, rdExecDirect
'       End If
'    Next
    
    
    If bAdmin Then
        'remove acesso dos grupos
 '        For j = 0 To lstGrupo.ListCount - 1
  '          Sql = "sp_droprolemember '" & lstGrupo.List(j) & "','" & sLogin & "'"
'            cn.Execute Sql, rdExecDirect
   '      Next
        'fornece acesso aos grupos
    '     For j = 0 To lstGrupo.ListCount - 1
         '   If lstGrupo.Selected(j) = True Then
     '           Sql = "sp_addrolemember '" & lstGrupo.List(j) & "','" & sLogin & "'"
      '          cn.Execute Sql, rdExecDirect
       '     End If
        ' Next
        'atualiza usuário na lista da tela
         lstUser.List(lstUser.ListIndex) = txtNome.Text
      End If
End If

Grava = True

'acesso padrão
If bAdmin Then
    DefaultAccess sLogin
End If

If lblAss.Caption <> "" Then
    'grava foto
    Dim mStream As New ADODB.Stream
    Dim rst As New ADODB.Recordset
    Dim adoConn As New ADODB.Connection
    
    adoConn.CursorLocation = adUseClient
    'adoConn.Open "Provider=SQLOLEDB.1;Persist Security Info=True;User ID=" & NomeDeLogin & ";Password=" & UserPwd & ";Initial Catalog=Tributacao;Data Source=" & IPServer
    adoConn.Open cn.Connect
        
    rst.Open "Select * from assinatura where usuario='" & txtLogin.Text & "'", adoConn, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        With mStream
            .Type = adTypeBinary
            .Open
            .LoadFromFile lblAss.Caption
            rst("fotoass").value = .Read
            rst.Update
        End With
        Set mStream = Nothing
    Else
        MsgBox "Sem assinatura cadastrada.", vbCritical, "Erro"
    End If
End If

End Function

