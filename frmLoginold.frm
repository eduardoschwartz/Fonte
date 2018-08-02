VERSION 5.00
Object = "{B60B1875-E5CA-11D2-BC3D-78A407C10000}#1.0#0"; "ksdpanel.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login do Sistema"
   ClientHeight    =   2205
   ClientLeft      =   3720
   ClientTop       =   3195
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4410
   Begin VB.CheckBox chkTeste 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Utilizar Base de Testes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   1950
      TabIndex        =   2
      Top             =   1710
      Width           =   2355
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   615
      Left            =   750
      TabIndex        =   11
      Top             =   1050
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Entrar"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   15658734
      MPTR            =   1
      MICON           =   "frmLogin.frx":030A
      PICN            =   "frmLogin.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KSDPanel.Panel Panel1 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   2490
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   53
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPwd1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1740
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2670
      Width           =   2265
   End
   Begin VB.TextBox txtPwd2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1740
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3060
      Width           =   2265
   End
   Begin VB.TextBox txtPwd 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2865
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1260
      TabIndex        =   0
      Top             =   210
      Width           =   2865
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   615
      Left            =   1950
      TabIndex        =   12
      Top             =   1050
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&Sair"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":0C00
      PICN            =   "frmLogin.frx":0C1C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPwd 
      Height          =   615
      Left            =   3150
      TabIndex        =   13
      Top             =   1050
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Se&nha"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":0F36
      PICN            =   "frmLogin.frx":0F52
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
      Left            =   1950
      TabIndex        =   14
      ToolTipText     =   "Cancelar Edição"
      Top             =   3480
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":126C
      PICN            =   "frmLogin.frx":1288
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
      Left            =   3030
      TabIndex        =   15
      ToolTipText     =   "Gravar os Dados"
      Top             =   3480
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":13E2
      PICN            =   "frmLogin.frx":13FE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblSeg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando Segurança. Aguarde..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   870
      TabIndex        =   10
      Top             =   1980
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nova Senha.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2730
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmação.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3090
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   210
      TabIndex        =   6
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   210
      TabIndex        =   5
      Top             =   270
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmLogin.frx":17A3
      Top             =   1380
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "frmLogin.frx":1AAD
      Top             =   1650
      Width           =   480
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bSair As Boolean

Private Sub cmdCancel_Click()
cmdPwd_Click
End Sub

Private Sub cmdGravar_Click()
Dim RdoAux As rdoResultset
Dim qd As New rdoQuery

If Trim$(txtUser.text) = "" Then
     MsgBox "Digite o Nome do Usuário.", vbExclamation, "Atenção"
     txtUser.SetFocus
     Exit Sub
End If

If Trim$(txtPwd.text) = "" Then
     MsgBox "Digite a Senha atual.", vbExclamation, "Atenção"
     txtPwd.SetFocus
     Exit Sub
End If

If Not Conecta(txtUser.text, txtPwd.text) Then
    Screen.MousePointer = vbdefualt
     MsgBox "Usuário e/ou Senha inválido(s)." & vbCrLf & "Verifique e tente novamente.", vbCritical, "Falha na Autenticação."
     Exit Sub
Else
     cn.Close
End If

If Len(txtPwd1.text) < 6 Then
     MsgBox "A Senha nova deve ter no mínimo 6 caracteres.", vbExclamation, "Atenção"
     txtPwd1.SetFocus
     Exit Sub
End If

If Trim$(txtPwd1.text) <> Trim$(txtPwd2.text) Then
     MsgBox "A Confirmação não é igual a nova senha.", vbExclamation, "Atenção"
     txtPwd1.SetFocus
     Exit Sub
End If

Conecta txtUser.text, txtPwd.text
Sql = "sp_password '" & txtPwd.text & "','" & txtPwd1.text & "'"
cn.Execute Sql, rdExecDirect
cn.Close

MsgBox "Sua Senha foi alterada com sucesso.", vbInformation, "SQL Server"
txtPwd.text = txtPwd1.text
cmdPwd_Click

End Sub

Private Sub cmdOK_Click()
Dim sUs As String
Dim Sql As String, RdoAux As rdoResultset, sDataBase As String, nUses As Integer
Dim cOS As New clsOS, sParam As String
    
If Me.Height = 4380 Then
     cmdGravar_Click
     Exit Sub
End If

If UCase$(txtPwd.text) = "NEWUSER" Then
   MsgBox "Entre em contato com o Setor de Informática para alterar sua senha !", vbExclamation, "Atenção"
'   cmdPwd_Click
Exit Sub
End If

Ocupado

'sParam = UCase(Command())
If chkTeste.Value = 0 Then
    sParam = ""
Else
    sParam = "-T"
End If

If Conecta(txtUser.text, txtPwd.text, sParam) Then
    Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='DATABASE'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount = 0 Then
           Sql = "INSERT PARAMETROS(NOMEPARAM,VALPARAM) VALUES('DATABASE'" & ",'" & CStr(Format(Now, "dd/mm/yyyy")) & "')"
           cn.Execute Sql, rdExecDirect
           sDataBase = CStr(Format(Now, "dd/mm/yyyy"))
        Else
           sDataBase = !VALPARAM
        End If
       .Close
    End With
    
    frmMdi.Sbar.Panels(6).text = "Data Base: " & sDataBase
    lblSeg.Visible = True
    lblSeg.Refresh
    bSair = True
    frmMdi.Sbar.Panels(1).text = cOS.OS_Name & " " & cOS.OS_ProductType & " " & cOS.OS_Version & " " & "Build:" & cOS.OS_Build & " " & cOS.OS_ServicePack
    Set cOS = Nothing
    frmMdi.Sbar.Panels(1).Enabled = True
    frmMdi.Sbar.Panels(2).text = "Usuario: " & txtUser.text
    frmMdi.Sbar.Panels(2).Enabled = True
    frmMdi.Sbar.Refresh
    Log Logon, Me.Name, Nenhum, "Logon no Sistema"
    LastUser = txtUser.text
    UserPwd = txtPwd.text
    nCodLastUser = RetornaIdUsuario(LastUser)
    sUs = ChrW$(83) + ChrW$(85) + ChrW$(80) + ChrW$(69) + ChrW$(82) + ChrW$(86) + ChrW$(73) + ChrW$(83) + ChrW$(79) + ChrW$(82)
'    If InStr(1, UCase$(cn.Connect), "DEVELOPER", vbBinaryCompare) > 0 Then
'        HabilitaMenu
'    Else
        If UCase(txtUser.text) <> sUs Then
           BoneHagana
        Else
           HabilitaMenu
        End If
'    End If
    lblSeg.Visible = False
    lblSeg.Refresh
    frmMdi.Timer1.Interval = 60000
    SaveSetting "GTI", "GERAL", "LASTUSER", txtUser.text
    CarregaDicionario
    If UCase(Command()) = "-L" Then
        frmMdi.frTeste.Visible = True
        frmMdi.frTeste.Caption = "ACESSANDO OS DADOS DAQUI MEMO"
    End If
    If InStr(1, cn.Connect, "TributacaoTeste", vbBinaryCompare) > 0 Then
        frmMdi.frTeste.Visible = True
    Else
        frmMdi.frTeste.Visible = False
    End If
    
    Unload Me
INICIO:
    If Not Security() Then
       If MsgBox("A Chave cadastrada não é válida ou expirou." & vbCrLf & vbCrLf & "Deseja cadastrar a chave agora ?", vbCritical + vbYesNo, "Acesso Negado ao Sistema !!!") = vbYes Then
          frmChave.show vbModal, frmMdi
          GoTo INICIO
       Else
          MsgBox "Você foi Desconectado do Sistema.", vbCritical, "INSTALAÇÃO NÃO AUTORIZADA"
          Unload frmMdi
       End If
    End If
Else
    Screen.MousePointer = vbdefualt
    MsgBox "Usuário e/ou Senha inválido(s)." & vbCrLf & "Verifique e tente se logar novamente.", vbCritical, "Falha na Autenticação."
End If
Liberado

End Sub

Private Sub BoneHagana()
Dim RdoAux As rdoResultset, Sql As String
Dim nCodUser As Integer

nCodUser = nCodLastUser

'***********Menus**************

For x = 1 To frmMdi.m_cMenuTabela.Count
     frmMdi.m_cMenuTabela.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuImobiliario.Count
     frmMdi.m_cMenuImobiliario.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuCadastro.Count
     frmMdi.m_cMenuCadastro.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuOpcoes.Count
     frmMdi.m_cMenuOpcoes.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuMobiliario.Count
     frmMdi.m_cMenuMobiliario.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuTributario.Count
     frmMdi.m_cMenuTributario.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuAvancado.Count
     frmMdi.m_cMenuAvancado.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuRelatorio.Count
     frmMdi.m_cMenuRelatorio.Enabled(x) = False
Next


HabilitaMenuDefault

Sql = "SELECT DISTINCT SEG_MENUACESSO.NOMEMENU FROM SEG_USERACESS INNER JOIN "
Sql = Sql & "SEG_MENUACESSO ON SEG_USERACESS.CODTELA = SEG_MENUACESSO.CODTELA "
Sql = Sql & "WHERE SEG_USERACESS.nomeUSUARIO = '" & txtUser.text & "' AND CODEVENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
      Do Until .EOF
            For x = 1 To frmMdi.m_cMenuTabela.Count
                 If frmMdi.m_cMenuTabela.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuTabela.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuTabela.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuCadastro.Count
                 If frmMdi.m_cMenuCadastro.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuCadastro.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuCadastro.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuImobiliario.Count
                 If frmMdi.m_cMenuImobiliario.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuImobiliario.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuImobiliario.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuMobiliario.Count
                 If frmMdi.m_cMenuMobiliario.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuMobiliario.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuMobiliario.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuOpcoes.Count
                 If frmMdi.m_cMenuOpcoes.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuOpcoes.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuOpcoes.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuTributario.Count
                 If frmMdi.m_cMenuTributario.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuTributario.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuTributario.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuAvancado.Count
                 If frmMdi.m_cMenuAvancado.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuAvancado.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuAvancado.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
            For x = 1 To frmMdi.m_cMenuRelatorio.Count
                 If frmMdi.m_cMenuRelatorio.ItemKey(x) <> "" Then
                    If frmMdi.m_cMenuRelatorio.ItemKey(x) = !NomeMenu Then
                       frmMdi.m_cMenuRelatorio.Enabled(x) = True
                       GoTo PROXIMO
                    End If
                 End If
            Next
PROXIMO:
          .MoveNext
      Loop
End With

End Sub

Private Sub HabilitaMenuDefault()
Dim nIndex As Long


nIndex = frmMdi.m_cMenuTabela.IndexForKey("mnuChangeUser")
frmMdi.m_cMenuTabela.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuTabela.IndexForKey("mnuClose")
frmMdi.m_cMenuTabela.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuConsulta.IndexForKey("mnuCnsProcesso")
frmMdi.m_cMenuConsulta.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuRDIV")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuCertidoes")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuTBAS")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuTFAT")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuTPMO")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuRelAtivTL")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuRelAtivISS")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuRelAtivISSFixo")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuRelDevedorVariavel")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuRelatorio.IndexForKey("mnuRelCartaCobrança")
frmMdi.m_cMenuRelatorio.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuTabela.IndexForKey("mnuTabSistemaTabelasBásicas")
frmMdi.m_cMenuTabela.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuTabela.IndexForKey("mnuFatores")
frmMdi.m_cMenuTabela.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuTributario.IndexForKey("mnuCalcTit")
frmMdi.m_cMenuTributario.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuTributario.IndexForKey("mnuDivAtiv")
frmMdi.m_cMenuTributario.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuOpcoes.IndexForKey("mnuSegurança")
frmMdi.m_cMenuOpcoes.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuTributario.IndexForKey("mnuCnsNumDV")
frmMdi.m_cMenuTributario.Enabled(nIndex) = True
nIndex = frmMdi.m_cMenuImobiliario.IndexForKey("mnuDetImovel")
frmMdi.m_cMenuImobiliario.Enabled(nIndex) = True


End Sub
 
Private Sub cmdPwd_Click()
If Me.Height = 4380 Then
     Me.Height = 2745
     cmdOK.Enabled = True
     cmdSair.Enabled = True
Else
     Me.Height = 4380
     cmdOK.Enabled = False
     cmdSair.Enabled = False
End If
End Sub

Private Sub cmdSair_Click()

Unload frmMdi
End
End Sub

Private Sub Form_Activate()
On Error Resume Next
txtPwd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
     cmdOK_Click
ElseIf KeyAscii = vbKeyEscape Then
     cmdSair_Click
End If

End Sub

Private Sub Form_Load()

Ocupado

lblSeg.Visible = False
Me.Height = 2745
LastUser = GetSetting("GTI", "GERAL", "LASTUSER")
frmMdi.Timer1.Interval = 0
txtUser.text = LastUser

bSair = False
Liberado

End Sub

Private Sub txtPwd_GotFocus()
txtPwd.SelStart = 0
txtPwd.SelLength = Len(txtPwd.text)
End Sub

Private Sub txtUser_GotFocus()
txtUser.SelStart = 0
txtUser.SelLength = Len(txtUser.text)
End Sub

Private Sub txtUser_LostFocus()
txtUser.text = UCase$(txtUser.text)

End Sub

Private Sub HabilitaMenu()

For x = 1 To frmMdi.m_cMenuTabela.Count
     frmMdi.m_cMenuTabela.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuImobiliario.Count
     frmMdi.m_cMenuImobiliario.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuCadastro.Count
     frmMdi.m_cMenuCadastro.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuOpcoes.Count
     frmMdi.m_cMenuOpcoes.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuMobiliario.Count
     frmMdi.m_cMenuMobiliario.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuTributario.Count
     frmMdi.m_cMenuTributario.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuAvancado.Count
     frmMdi.m_cMenuAvancado.Enabled(x) = True
Next
For x = 1 To frmMdi.m_cMenuRelatorio.Count
     frmMdi.m_cMenuRelatorio.Enabled(x) = True
Next

End Sub

Private Sub CarregaDicionario()
Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT ANOJUROS,PERCJUROS FROM JUROS ORDER BY ANOJUROS"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        dcJuros.Add !ANOJUROS, !PERCJUROS
       .MoveNext
    Loop
   .Close
End With

Sql = "SELECT ANOUFIR,VALORUFIR FROM UFIR ORDER BY ANOUFIR"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        dcUfir.Add !ANOUFIR, Virg2Ponto(!VALORUFIR)
       .MoveNext
    Loop
   .Close
End With



ReDim aMulta(0)

Sql = "SELECT ANOMULTA,MINDIA,MAXDIA,PERCDIA FROM MULTA ORDER BY ANOMULTA,MINDIA,MAXDIA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         ReDim Preserve aMulta(UBound(aMulta) + 1)
         aMulta(UBound(aMulta)).nAno = !ANOMULTA
         aMulta(UBound(aMulta)).nMin = !MINDIA
         aMulta(UBound(aMulta)).nMax = !MAXDIA
         aMulta(UBound(aMulta)).nValor = !PERCDIA
        .MoveNext
    Loop
End With

End Sub
