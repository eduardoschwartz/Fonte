VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login do Sistema"
   ClientHeight    =   2235
   ClientLeft      =   4860
   ClientTop       =   3555
   ClientWidth     =   4410
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGravar 
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Gravar"
      Height          =   285
      Left            =   1935
      TabIndex        =   14
      Top             =   3510
      Width           =   915
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   2925
      TabIndex        =   15
      Top             =   3510
      Width           =   915
   End
   Begin VB.CheckBox chkLocal 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Local"
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
      Left            =   900
      TabIndex        =   13
      Top             =   1710
      Visible         =   0   'False
      Width           =   855
   End
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
      TabIndex        =   5
      Top             =   1710
      Width           =   2355
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
      TabIndex        =   6
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
      TabIndex        =   7
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
      Top             =   555
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
      Top             =   180
      Width           =   2865
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   630
      Left            =   1935
      TabIndex        =   3
      Top             =   990
      Width           =   1170
      _ExtentX        =   2064
      _ExtentY        =   1111
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
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
   Begin prjChameleon.chameleonButton cmdOK 
      Height          =   630
      Left            =   630
      TabIndex        =   2
      Top             =   990
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1111
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmLogin.frx":0640
      PICN            =   "frmLogin.frx":065C
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
      Height          =   630
      Left            =   3195
      TabIndex        =   4
      Top             =   990
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1111
      BTYPE           =   3
      TX              =   "S&enha"
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
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
      ForeColor       =   &H00400000&
      Height          =   195
      Left            =   870
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
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
      TabIndex        =   9
      Top             =   615
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
      TabIndex        =   8
      Top             =   225
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   90
      Picture         =   "frmLogin.frx":182C
      Top             =   1380
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   300
      Picture         =   "frmLogin.frx":1B36
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

If Trim$(txtUser.Text) = "" Then
     MsgBox "Digite o Nome do Usuário.", vbExclamation, "Atenção"
     txtUser.SetFocus
     Exit Sub
End If

If Trim$(txtPwd.Text) = "" Then
     MsgBox "Digite a Senha atual.", vbExclamation, "Atenção"
     txtPwd.SetFocus
     Exit Sub
End If

If Len(txtPwd1.Text) < 6 Then
     MsgBox "A Senha nova deve ter no mínimo 6 caracteres.", vbExclamation, "Atenção"
     txtPwd1.SetFocus
     Exit Sub
End If

If Trim$(txtPwd1.Text) <> Trim$(txtPwd2.Text) Then
     MsgBox "A Confirmação não é igual a nova senha.", vbExclamation, "Atenção"
     txtPwd1.SetFocus
     Exit Sub
End If

Conecta UL, UP

Sql = "update usuario set senha='" & Encrypt128(Mask(txtPwd1.Text), "everest") & "' where nomelogin='" & txtUser.Text & "'"
cn.Execute Sql, rdExecDirect

MsgBox "Sua Senha foi alterada com sucesso.", vbInformation, "SQL Server"
txtPwd.Text = txtPwd1.Text
cmdPwd_Click

End Sub

Private Sub cmdOK_Click()
Dim sUs As String, RdoAux2 As rdoResultset
Dim Sql As String, RdoAux As rdoResultset, sDataBase As String
Dim cOS As New clsOS, sParam As String


If Me.Height = 4380 Then
     cmdGravar_Click
     Exit Sub
End If

Ocupado
bLocal = False
If chkTeste.value = 0 Then
    If chkLocal.value = 1 Then
        bLocal = True
        sParam = "-L"
        ConectaDBBKP
    Else
        sParam = ""
    End If
Else
    sParam = "-T"
End If


'If txtUser.Text <> "RENATA" And txtUser.Text <> "SOLANGE" Then
'    MsgBox "Sistema bloqueado!"
'    Exit Sub
'End If

If NomeDoComputador = "SKYNET" Then
    sPathAnexo = "D:\Trabalho\GTI\Documentos\"
    bFichaCompensacao = True
Else
    sPathAnexo = "\\192.168.200.130\atualizagti\documentos\"
    bFichaCompensacao = False
End If



If Not Conecta(UL, UP, sParam) Then Exit Sub
GetLanguage
CarregaDicionario
Sql = "select nomelogin,nomecompleto,senha,ativo from usuario where nomelogin='" & txtUser.Text & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    If .RowCount > 0 Then
        If Val(SubNull(!Ativo)) <> 1 Then
            Liberado
            MsgBox "A conta do usuário foi desativada no GTI." & vbCrLf & " Favor entrar em contato com o administrador do sistema." & vbCrLf & "(gti@jaboticabal.sp.gov.br).", vbCritical, "Atenção"
            Exit Sub
        End If
       
        
        If Len(txtPwd.Text) < 6 And txtUser.Text <> "SCHWARTZ" Then
            MsgBox "A senha deve ter no mínimo 6 caracteres.", vbCritical, "Atenção"
            Exit Sub
        End If
        
        If IsNull(!SENHA) Then
            Sql = "update usuario set senha='" & Encrypt128(Mask(txtPwd.Text), UP) & "' where nomelogin='" & txtUser.Text & "'"
            cn.Execute Sql, rdExecDirect
        Else
            If sParam <> "-L" Then
                If Decrypt128(!SENHA, UP) <> txtPwd.Text Then
                    Screen.MousePointer = vbdefualt
                    MsgBox "Usuário e/ou Senha inválido(s)." & vbCrLf & "Verifique e tente se logar novamente.", vbCritical, "Falha na Autenticação."
                    txtPwd.Text = "": txtPwd.SetFocus
                    Exit Sub
                End If
            End If
        End If
      
        '****
        'Só uma instância por usuário
        
'        Sql = "SELECT machines.usuario, machines.ip, machines.computer, machines.nome, usuario.ativo, usuario.logon, usuario.datalogon "
'        Sql = Sql & "FROM machines INNER JOIN usuario ON machines.usuario = usuario.nomelogin where usuario='" & txtUser.Text & "' and "
'        Sql = Sql & "datalogon >= '" & Format(Now, "mm/dd/yyyy") & "' and computer <> '" & NomeDoComputador & "' and logon=1"
'        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'        If RdoAux.RowCount > 0 Then
'            MsgBox "O seu login já está sendo utilizado em outro computador:" & vbCrLf & vbCrLf & "Computador: " & RdoAux!computer & vbCrLf & "Endereço IP: " & RdoAux!ip & vbCrLf & "Usuário windows: " & RdoAux!nome, vbCritical, "ACESSO NEGADO!"
'            RdoAux.Close
'            Exit Sub
'        End If
'        RdoAux.Close
        
        
        '***
      
        Sql = "update usuario set logon=1,datalogon='" & Format(Now, sDataFormat) & "' where nomelogin='" & txtUser.Text & "'"
        cn.Execute Sql, rdExecDirect
        
        If InStr(1, cn.Connect, "TributacaoTeste", vbBinaryCompare) > 0 Then
            Sql = "USE TributacaoTeste"
        ElseIf InStr(1, cn.Connect, "tributacaoBKP", vbBinaryCompare) > 0 Then
            Sql = "USE TributacaoBKP"
        Else
            Sql = "USE Tributacao"
        End If
        
        cn.Execute Sql
        bCloseChat = False
        If txtUser.Text = "SCHWARTZ" Then GoTo P1
        If InStr(1, UCase(Command$), "-NOVERSION", vbBinaryCompare) = 0 Then
            Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='VS'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If App.Minor >= Left(!valparam, 1) Then
                    If App.Revision < Right(!valparam, 3) Then
                        MsgBox "Atualize a versão do GTI para 2." & Left(!valparam, 1) & "." & Right(!valparam, 3) & " para poder acessar.", vbOKOnly + vbCritical, "Desconectado do Sistema"
                        Exit Sub
                    End If
                Else
                    MsgBox "Atualize a versão do GTI para 2." & Left(!valparam, 1) & "." & Right(!valparam, 3) & " para poder acessar.", vbOKOnly + vbCritical, "Desconectado do Sistema"
                    Exit Sub
                End If
               .Close
            End With
        End If
P1:
        Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='DATABASE'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
               Sql = "INSERT PARAMETROS(NOMEPARAM,VALPARAM) VALUES('DATABASE'" & ",'" & CStr(Format(Now, "dd/mm/yyyy")) & "')"
               cn.Execute Sql, rdExecDirect
               sDataBase = CStr(Format(Now, "dd/mm/yyyy"))
            Else
               sDataBase = !valparam
            End If
           .Close
        End With
        
        Sql = "select * from machines2 where computer='" & NomeDoComputador & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount = 0 Then
                Sql = "insert machines2(computer,margin_top,margin_bottom,margin_left,margin_right) values('" & NomeDoComputador & "',0,0,0,0)"
                cn.Execute Sql, rdExecDirect
                nMargem_Top = 0
                nMargem_Left = 0
                nMargem_Right = 0
                nMargem_Bottom = 0
            Else
                nMargem_Top = !margin_top
                nMargem_Left = !Margin_left
                nMargem_Right = !margin_right
                nMargem_Bottom = !margin_bottom
            End If
           .Close
        End With
        
        
        
        
        If frmMdi.frTeste.Visible = False Then
            NomeBaseDados = "TRIBUTACAO"
        Else
            If sParam = "-T" Then
                NomeBaseDados = "TRIBUTACAOTESTE"
                frmMdi.frTeste.Caption = "BASE DE TESTE"
            Else
                NomeBaseDados = "TRIBUTACAOBKP"
                frmMdi.frTeste.Caption = "BASE PARALELA - BACKUP DE 02/09/2013"
            End If
        End If
        
        Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='CIDADE'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
           NomeCidade = RdoAux!valparam
           RdoAux.Close
        End If
        Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='ANISTIA'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
        With RdoAux
            If .RowCount > 0 Then
                bAnistia = IIf(!valparam = 1, True, False)
            End If
           .Close
        End With
        
   
        
        frmMdi.Sbar.Panels(6).Text = "Data Base: " & sDataBase
        lblSeg.Visible = True
        lblSeg.Refresh
        bSair = True
        frmMdi.Sbar.Panels(1).Text = cOS.OS_Name & " " & cOS.OS_ProductType & " " & cOS.OS_Version & " " & "Build:" & cOS.OS_Build & " " & cOS.OS_ServicePack
        Set cOS = Nothing
        NomeDeLogin = Trim$(txtUser.Text)
        If UCase(NomeDeLogin) = "LEONE" Then NomeDeLogin = "ELAINE"
        frmMdi.Sbar.Panels(1).Enabled = True
        frmMdi.Sbar.Panels(2).Text = "Usuario: " & NomeDeLogin
        frmMdi.Sbar.Panels(2).Enabled = True
        frmMdi.Sbar.Refresh
        sWd = txtPwd.Text
        LastUser = txtUser.Text
        UserPwd = txtPwd.Text
        nCodLastUser = RetornaUsuarioID(LastUser)
        
        bComercioEletronico = False
'        If NomeDeLogin = "SCHWARTZ" Or IsAtendente Then
            bComercioEletronico = True
 '       End If
        
        CloseApplication
        
        If IsVBIDE Then
            NewSec = True
            NewSec = False
        Else
            NewSec = False
        End If
        
        
        bNovoBoleto = True
        
        
        If NewSec = False Then
            If UCase(NomeDeLogin) <> "SCHWARTZ" Then
               BoneHagana
            Else
               HabilitaMenu
            End If
        Else
            FillSec
        End If
        
        lblSeg.Visible = False
        lblSeg.Refresh
        'frmMdi.Timer2.Interval = 60000
        SaveSetting "GTI", "GERAL", "LASTUSER", txtUser.Text
        
     '   Sql = "UPDATE USUARIO SET LOGON=1,DATALOGON='" & Format(Now, "mm/dd/yyyy") & "' WHERE NOMELOGIN='" & NomeDeLogin & "'"
    '    cn.Execute Sql, rdExecDirect
        
        
        sPrintBottom = GetSetting("GTI", "PRINT", "BOTTOM")
        If sPrintBottom = "S" Then
            frmMdi.m_cMenuPrincipal.Checked(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True
        End If
        
        On Error Resume Next
        If FileDateTime(App.Path & "\TRIBUTACAO.EXE") < CDate("19/06/2015") Then
            FileCopy "\\192.168.200.130\ATUALIZAGTI\TRIBUTACAO.EXE", App.Path & "\TRIBUTACAO.EXE"
        End If
        
        
      
        
        
HERE2:
        CarregaDicionario
        If chkLocal.value = 1 Then
            frmMdi.Picture4.Visible = True
            frmMdi.frTeste.Visible = True
            frmMdi.frTeste.Caption = "BASE PARALELA - BACKUP DE 02/09/2013"
            NomeBaseDados = "TRIBUTACAOBKP"
            GoTo HERE
        End If
        If InStr(1, cn.Connect, "TributacaoTeste", vbBinaryCompare) > 0 Then
            frmMdi.Picture4.Visible = True
            frmMdi.frTeste.Visible = True
        Else
            frmMdi.Picture4.Visible = False
            frmMdi.frTeste.Visible = False
        End If
        modLg "Acesso ao Sistema"
        frmChat.Visible = False
        
        
        'PARCELAMENTO NOVO
'        If frmMdi.frTeste.Visible = False Then
'            frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamento")) = False
'        Else
'            If NomeDeLogin = "SCHWARTZ" Then
'                frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamento")) = True
'            End If
'        End If

        
HERE:
        Unload Me
Inicio:
    
    
    
    Else
        Screen.MousePointer = vbDefault
        MsgBox "Usuário e/ou Senha inválido(s)." & vbCrLf & "Verifique e tente se logar novamente.", vbCritical, "Falha na Autenticação."
        txtPwd.Text = "": txtPwd.SetFocus
        Exit Sub
    End If
End With


Liberado

End Sub

Private Sub BoneHagana()
Dim RdoAux As rdoResultset, Sql As String
Dim nCodUser As Integer, o As Object, x As Integer

nCodUser = nCodLastUser

For x = 1 To frmMdi.m_cMenuPrincipal.Count
    frmMdi.m_cMenuPrincipal.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuParam.Count
    frmMdi.m_cMenuParam.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuImob.Count
    frmMdi.m_cMenuImob.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuMob.Count
    frmMdi.m_cMenuMob.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuAtende.Count
    frmMdi.m_cMenuAtende.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuTrib.Count
    frmMdi.m_cMenuTrib.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuProt.Count
    frmMdi.m_cMenuProt.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuOutro.Count
    frmMdi.m_cMenuOutro.Enabled(x) = False
Next

Sql = "SELECT DISTINCT SEG_MENUACESSO.NOMEMENU FROM SEG_USERACESS INNER JOIN "
Sql = Sql & "SEG_MENUACESSO ON SEG_USERACESS.CODTELA = SEG_MENUACESSO.CODTELA "
Sql = Sql & "WHERE SEG_USERACESS.nomeUSUARIO = '" & txtUser.Text & "' AND CODEVENTO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
        Do Until .EOF
            'If UCase(!nomemenu) = "MNUCIDADAO" Then MsgBox "TESTE"
            For x = 1 To frmMdi.m_cMenuPrincipal.Count
                If frmMdi.m_cMenuPrincipal.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuPrincipal.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuParam.Count
                If frmMdi.m_cMenuParam.ItemKey(x) = !NOMEMENU Then
                    
                   frmMdi.m_cMenuParam.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuImob.Count
                If frmMdi.m_cMenuImob.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuImob.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuMob.Count
                If frmMdi.m_cMenuMob.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuMob.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuAtende.Count
                If frmMdi.m_cMenuAtende.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuAtende.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuTrib.Count
                If frmMdi.m_cMenuTrib.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuTrib.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuOutro.Count
                If frmMdi.m_cMenuOutro.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuOutro.Enabled(x) = True
                End If
            Next
            For x = 1 To frmMdi.m_cMenuProt.Count
                If frmMdi.m_cMenuProt.ItemKey(x) = !NOMEMENU Then
                   frmMdi.m_cMenuProt.Enabled(x) = True
                End If
            Next
                      
           .MoveNext
        Loop
End With
HabilitaMenuDefault
End Sub

Private Sub HabilitaMenuDefault()
Dim nIndex As Long
NomeDeLogin = UCase(NomeDeLogin)
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrinter")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuSelectPrinter")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuChangeUser")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuClose")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTerritorial")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTributário")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTTRI")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuOutrosParam")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCadastro")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuConsultaImovel")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuAtividadeImovel")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuRelatorioImob")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuDevedorIPTU")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuRural")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuGTIRural")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuLancRocada")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnu2ViaLaser")) '156->72
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuMalaDiretaRocada")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnu2ViaLaser")) '156->71
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuMalaDiretaISSCCivil")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnu2ViaLaser")) '156->71
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuMalaDiretaRural")) = frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCadRural")) '156->300
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuDevedorIPTU")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuOutroImob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelatorioMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuSN")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuDevedores")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAtividadeMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadastroMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAtividadeMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAtividadeCadMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuConsultaMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuProdutividade")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaRua")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaCNPJ")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaAtividade")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaContador")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaSocio")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaAtiv")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnurelParcNPago")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelMEI")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelEstimado")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvaras")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuDevedores")) = True
'frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvaraEmitido")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelatorioMob1")) = frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaPagSN")) '132->133
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuSenha")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuRelatorioAte")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCobrancaJudicial")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuConsultaAte")) = True
'frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuSenhaISS")) = True
'frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuRenovaAlvara")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCidadao")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuMalaDiretaCidadao")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuDepositoCRI")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->161
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuDeclaraIsentoIPTU")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->166
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuRequerIsentoIPTU")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->167
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuGuiaPratico1")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->162
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuGuiaPratico2")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->163
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuGuiaPratico5")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->164
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuGuiaPratico3")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) '169->165

frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuOutrosT")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuSituacaoTributaria")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuCalcTit")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuAtivBanco")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuAlugueis")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDividaAtivaT")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuRelatorioTrib")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuBuscaArq")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDebitoAjPago")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuComplementoPagto")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuSimples")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuOptanteDARel")) = frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuOptanteDA"))
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuCobranca")) = frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuOptanteDA"))

frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuParametroProt")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuTramiteAberto")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuTramiteEnviado")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoAssunto")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoCC")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoAno")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuAssuntoDoc")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuAnexos")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuAdmin")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuSeguranca")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuSeguranca")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuSecretariaObra")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuParamObra")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuNovaGIA")) = frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadMobiliario"))
'frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuSilDeca")) = frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadMobiliario"))
'frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnu2vianotificacao")) = frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadMobiliario"))

If frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadMobiliario")) = False And frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCadImob")) = False Then
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRepParcNPIPTU")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRepParcNPISS")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaAtividade")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaCNPJ")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaContador")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaRua")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelEstimado")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelMEI")) = False
   ' frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvaraEmitido")) = False
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuDevedorIPTU")) = False
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTLAN")) = False
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCobrancaJudicial")) = False
    'frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuSenhaISS")) = False
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuBuscaArq")) = False
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuComplementoPagto")) = False
'    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDebitoAjPago")) = False
  '  frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("CnsDebitoImob")) = False
End If

frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuITBIObs")) = frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuITBI"))
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuSenhaResumo")) = False
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCancelReparc")) = False
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCancelParcelamentoAuto")) = False
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuMalaDiretaParc")) = False
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuPagamentoMensalParc")) = False
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuBuscaArq")) = False


If NomeDeLogin = "LUIZH" Or NomeDeLogin = "LEANDRO" Or NomeDeLogin = "ROSANGELA" Or NomeDeLogin = "NOELI" Or NomeDeLogin = "EDUARDO" Or NomeDeLogin = "ROSE" Or NomeDeLogin = "CARMELINO" Then
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuITBIRel")) = True '239
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDebitoAjPago")) = True '236
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuSNCnpjReceita")) = True '217
End If

If NomeDeLogin = "RENATA" Or NomeDeLogin = "SOLANGE" Then
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnu2ViaEspecial")) = True '176
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuSenhaResumo")) = True '151
End If

If NomeDeLogin = "RENATA" Or NomeDeLogin = "ROSE" Or NomeDeLogin = "SOLANGE" Or NomeDeLogin = "ANA" Or NomeDeLogin = "GLEISE" Or NomeDeLogin = "LUIZH" Or NomeDeLogin = "JOSIANE" Or NomeDeLogin = "GLEISE" Or _
    NomeDeLogin = "RITA" Or NomeDeLogin = "DANIELAR" Or NomeDeLogin = "DANIELAT" Or NomeDeLogin = "TALITA" Or NomeDeLogin = "SIMONE" Or NomeDeLogin = "EDUARDO" Or NomeDeLogin = "AMFNFONSECA" Then
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuRelRefis")) = True '180
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuEmiteDoc")) = True '180
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuRelRefisParc")) = True '181
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDocEmitido")) = True '240
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuPagoTributo")) = True '240
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCancelReparc")) = True '168
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuCancelParcelamentoAuto")) = True '172
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuMalaDiretaParc")) = True '173
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuPagamentoMensalParc")) = True '174
    frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuTramiteAtraso")) = True
End If


 frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuEspolio")) = False '67
If NomeDeLogin = "SCHWARTZ" Or NomeDeLogin = "RENATA" Or NomeDeLogin = "IORIO" Or NomeDeLogin = "ROSE" Or NomeDeLogin = "JOSEANE" Or NomeDeLogin = "JOSIANE" Or NomeDeLogin = "HELOISA" Or _
     NomeDeLogin = "FERNANDA.SIMOLIN" Or NomeDeLogin = "SOLANGE" Or _
    NomeDeLogin = "FACTORE" Or NomeDeLogin = "MARIELA" Or NomeDeLogin = "TICYANNE.OKIMASU" Or NomeDeLogin = "MARIELA.CUSTODIO" Or NomeDeLogin = "JOAOF" Or IsAtendente Then
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuEspolio")) = True '67
End If

If NomeDeLogin = "LORAINE" Then
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuSimulaRural")) = True '86
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuRelCadRuralFull")) = True '83
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuEventoRural")) = True '298
End If

If frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCnsImovel")) = True Then
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCnsAvancadaImob")) = True
End If

If NomeDeLogin = "ROSE" Or NomeDeLogin = "JOESANE" Or NomeDeLogin = "DANIELE.SILVA" Then
    frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuPagamentoMensalParc")) = True
    frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuIntegrativa")) = True '289
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuComplementoPagto")) = False
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuPagamentoCC")) = True
End If

If frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadMobiliario")) = True Then
'    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvaraProvisorio")) = True
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnae")) = True
     frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuNFEmitida")) = True
'    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnaeVS")) = True
    
End If

If NomeDeLogin = "LUIZH" Or NomeDeLogin = "DANIELAR" Or NomeDeLogin = "LEANDRO" Or NomeDeLogin = "LEONARDODINIZ" Or NomeDeLogin = "ROSANGELA" Or NomeDeLogin = "PAULA" Or NomeDeLogin = "NOELI" Or NomeDeLogin = "MAURICIOJ" Or NomeDeLogin = "RODRIGOC" Or NomeDeLogin = "PAULO" Or NomeDeLogin = "MARILIA" Or NomeDeLogin = "SIMONE" Then
   ' frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuExporta")) = True '288
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuNotificacao2")) = True '234
End If

frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuCorrecaoCPF")) = True
If NomeDeLogin <> "RENATA" And NomeDeLogin <> "SOLANGE" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "ANA" And NomeDeLogin <> "ROSE" And NomeDeLogin <> "MARIELA.CUSTODIO" And NomeDeLogin <> "RITA" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "SIMONE" And NomeDeLogin <> "DANIELAT" And NomeDeLogin <> "HELOISA" And NomeDeLogin <> "FACTORE" And NomeDeLogin <> "LUIZH" Then
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuCorrecaoCPF")) = False
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuLogr")) = False '51
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuBairro")) = False '5
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuCidade")) = False '1
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTipoLog")) = False '10
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTitLog")) = False '14
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTLAN")) = False '18
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabTributoAliq")) = False '33
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTTRI")) = False '22
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTTLA")) = False '23
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuArtigoTributo")) = False '299
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaPPARC")) = False '29
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaUfir")) = False
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuFeriado")) = False '37
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuBanco")) = False 'ELIMINADO
    frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoAno")) = False '280
    frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuAssuntoDoc")) = False '280
Else
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuLogr")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuBairro")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuCidade")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTipoLog")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTitLog")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTLAN")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabTributoAliq")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTTRI")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaTTLA")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuArtigoTributo")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaPPARC")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTabSistemaUfir")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuFeriado")) = True
    frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuBanco")) = True
End If
   

If NomeDeLogin <> "ANA" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "CARLOS.SANTOS" And NomeDeLogin <> "MARIELA.CUSTODIO" And NomeDeLogin <> "HELOISA" And NomeDeLogin <> "JOAOF" Then
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuAverbacao")) = False '68
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuRolImovel")) = False '69
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuDevedorIPTU")) = False '70
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuMalaDiretaRocada")) = False '71
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuLancRocada")) = False '72
   
End If

If NomeDeLogin = "NOELI" Then
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuImportarSN")) = True
End If

If NomeDeLogin <> "HELOISA" Then
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCorrigeBairro")) = False '68
Else
    frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCorrigeBairro")) = True '68
End If

If NomeDeLogin <> "LUIZH" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "LEANDRO" And NomeDeLogin <> "EDUARDO" And NomeDeLogin <> "ROSANGELA" And NomeDeLogin <> "PAULA" And NomeDeLogin <> "RITA" And NomeDeLogin <> "NOELI" And NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "PAULO" And NomeDeLogin <> "MARILIA" And NomeDeLogin <> "VANESSA" And NomeDeLogin <> "DIONE" And NomeDeLogin <> "RENILDA" And NomeDeLogin <> "ALESSANDRA" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "ROSE" Then
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEscContab")) = False '92
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuTabAtivTL")) = False '96
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuTabAtivISS")) = False '100
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuVigSan")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnsEmpresaAvancada")) = False '105
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnsNF")) = False '106
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnsNFDoc")) = False '107
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnsISSVarPago")) = False '108
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCnae")) = False '109
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuSuspende")) = False '110
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuNotificaISS")) = False '114
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuSalaEmp")) = False '115
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuNovaGIA")) = False '116
'    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuSilDeca")) = False '117
'    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuProdutEvento")) = False
'    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuProdutTarefa")) = False
    'frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvara")) = False '122
'    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvaraEmitido")) = False '123
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuDevIssVar")) = False '124
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuDevIssEst")) = False '125
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelDevTaxaLic")) = False '126
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelDevTaxaLicAuto")) = False '127
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelDevTaxaLicAlvara")) = False '128
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelDevVigSanit")) = False '129
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelDevedorGeral")) = False '130
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelatorioMob1")) = False '133
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaSN")) = False '131
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaPagSN")) = False '132
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelatorioMob1")) = False 'REPETIDO
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaCnae")) = False 'NOVOVOVOVOVOVO
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaTL")) = False '134
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaTL3")) = False '135
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaIssVE")) = False '136
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaIssFixo")) = False '137
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaVS")) = False '138
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRepParcNPIPTU")) = False '139
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRepParcNPISS")) = False '140
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuMalaDireta")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuISSMensal")) = False '141
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaRua")) = False '142
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaCNPJ")) = False '143
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaContador")) = False '145
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelEstimado")) = False '146
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelMEI")) = False '147
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuIssPagoAtividade")) = False '148
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuResumoIssCCivil")) = False '149
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuNFEmitida")) = False
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuEmpresaAtividade")) = False '144
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuSituacaoTributaria")) = False '229
    'frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDebitoAjPago")) = False
    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuComplementoPagto")) = False '237
    
Else
    frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaCnae")) = True
End If

If NomeDeLogin <> "LUIZH" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "LEANDRO" And NomeDeLogin <> "EDUARDO" And NomeDeLogin <> "ROSANGELA" And NomeDeLogin <> "PAULA" And NomeDeLogin <> "RITA" And NomeDeLogin <> "NOELI" And NomeDeLogin <> "RODRIGOC" And NomeDeLogin <> "PAULO" And NomeDeLogin <> "MARILIA" And NomeDeLogin <> "VANESSA" And NomeDeLogin <> "DIONE" And NomeDeLogin <> "RENILDA" And NomeDeLogin <> "ALESSANDRA" And NomeDeLogin <> "GLEISE" And NomeDeLogin <> "ROSE" And Not IsAtendente Then
'    frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnu2vianotificacao")) = False '241
End If

'If Not IsAtendente Then
'        frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuCalcGeral")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoArquivado")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuEtiquetaProt")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuResumoDiario")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuPublicacao")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuTramiteEnviado")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoAssunto")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoCC")) = False
'        frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuProcessoAno")) = False
 '       frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuTramiteAberto")) = False
'End If


   
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
modLg000
lblSeg.Visible = False
Me.Height = 2745
LastUser = GetSetting("GTI", "GERAL", "LASTUSER")
txtUser.Text = LastUser
chkLocal.Visible = True
bSair = False
Liberado

End Sub

Private Sub txtPwd_GotFocus()
txtPwd.SelStart = 0
txtPwd.SelLength = Len(txtPwd.Text)
End Sub

Private Sub txtUser_GotFocus()
txtUser.SelStart = 0
txtUser.SelLength = Len(txtUser.Text)
End Sub

Private Sub txtUser_LostFocus()
txtUser.Text = UCase$(txtUser.Text)

End Sub

Private Sub HabilitaMenu()
Dim o As Object

For Each o In frmMdi.Controls
    If Left(o.Name, 3) = "mnu" Then
        o.Enabled = True
    End If
Next

End Sub

Private Sub CarregaDicionario()
Dim RdoAux As rdoResultset, Sql As String, sData As String

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

dcFeriado.RemoveAll
Sql = "SELECT * FROM FERIADODEF ORDER BY ANO,MES,DIA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sData = Format(!DIA, "00") & "/" & Format(!Mes, "00") & "/" & Format(!Ano, "0000")
'        dcFeriado.Add sData, !CODFERIADO
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

If NomeDeLogin <> "SCHWARTZ" Then
Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='CETLAN'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
   nSeqFator = Int(!valparam)
   .Close
End With
End If
nSeqFator2 = CDate(Format(Now, "dd/mm/yyyy")) - CDate(MIB_DB)

End Sub

Private Sub ConectaDBBKP()
Dim DataSourceName As String
Dim DatabaseName As String
Dim Description As String
Dim DriverPath As String
Dim DriverName As String
Dim LastUser As String
Dim Regional As String
Dim Server As String

Dim lResult As Long
Dim hKeyHandle As Long

DataSourceName = "odbcTribLocal"
DatabaseName = "TributacaoBKP"
Description = "Base Local do GTI"
DriverPath = "<path to your SQL Server driver>"
LastUser = ""
Server = IPServer
DriverName = "SQL Server"

'Create the new DSN key.

lResult = RegCreateKey(HKEY_LOCAL_MACHINE, "SOFTWARE\ODBC\ODBC.INI\" & _
     DataSourceName, hKeyHandle)

'Set the values of the new DSN key.

lResult = RegSetValueEx(hKeyHandle, "Database", 0&, REG_SZ, _
   ByVal DatabaseName, Len(DatabaseName))
lResult = RegSetValueEx(hKeyHandle, "Description", 0&, REG_SZ, _
   ByVal Description, Len(Description))
lResult = RegSetValueEx(hKeyHandle, "Driver", 0&, REG_SZ, _
   ByVal DriverPath, Len(DriverPath))
lResult = RegSetValueEx(hKeyHandle, "LastUser", 0&, REG_SZ, _
   ByVal LastUser, Len(LastUser))
lResult = RegSetValueEx(hKeyHandle, "Server", 0&, REG_SZ, _
   ByVal Server, Len(Server))

'Close the new DSN key.

lResult = RegCloseKey(hKeyHandle)

'Open ODBC Data Sources key to list the new DSN in the ODBC Manager.
'Specify the new value.
'Close the key.

lResult = RegCreateKey(HKEY_LOCAL_MACHINE, _
   "SOFTWARE\ODBC\ODBC.INI\ODBC Data Sources", hKeyHandle)
lResult = RegSetValueEx(hKeyHandle, DataSourceName, 0&, REG_SZ, _
   ByVal DriverName, Len(DriverName))
lResult = RegCloseKey(hKeyHandle)

End Sub

Function IsVBIDE() As Boolean
    Static stBeenHere As Boolean
    Static stResult As Boolean
    
    If stBeenHere Then
        IsVBIDE = stResult
    Else
        stBeenHere = True
        On Error GoTo Trap
        Debug.Print 1 / 0
    End If
    Exit Function
Trap:
    stResult = True
    IsVBIDE = True
    Exit Function
End Function

Private Sub FillSec()
Dim Sql As String, RdoAux As rdoResultset, x As Integer, aMenu() As String, o As Object, nPos As Integer

Sql = "SELECT count(*) as contador From sec_item"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
ReDim aMenu(RdoAux!contador)
RdoAux.Close

Sql = "select id from sec_user_item where usuario='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
SecId = RdoAux!Id
RdoAux.Close

Enable_Basic_Menu

Sql = "SELECT id, MenuName From sec_item Where (MenuName Is Not Null)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        aMenu(!Id) = !MenuName
       .MoveNext
    Loop
   .Close
End With

For x = 1 To Len(SecId)
    If BitId(x) = 0 Then
        aMenu(x) = ""
    End If
Next

For x = 1 To frmMdi.m_cMenuParam.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuParam.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuParam.Enabled(x) = True
           Exit For
        End If
    Next
Next
For x = 1 To frmMdi.m_cMenuImob.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuImob.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuImob.Enabled(x) = True
           Exit For
        End If
    Next
Next
For x = 1 To frmMdi.m_cMenuMob.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuMob.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuMob.Enabled(x) = True
           Exit For
        End If
    Next
Next
For x = 1 To frmMdi.m_cMenuAtende.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuAtende.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuAtende.Enabled(x) = True
           Exit For
        End If
    Next
Next
For x = 1 To frmMdi.m_cMenuTrib.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuTrib.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuTrib.Enabled(x) = True
           Exit For
        End If
    Next
Next
For x = 1 To frmMdi.m_cMenuProt.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuProt.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuProt.Enabled(x) = True
           Exit For
        End If
    Next
Next
For x = 1 To frmMdi.m_cMenuOutro.Count
    For nPos = 1 To UBound(aMenu)
        If frmMdi.m_cMenuOutro.ItemKey(x) = aMenu(nPos) Then
           frmMdi.m_cMenuOutro.Enabled(x) = True
           Exit For
        End If
    Next
Next

End Sub

Private Sub Enable_Basic_Menu()
Dim x As Integer

For x = 1 To frmMdi.m_cMenuPrincipal.Count
    frmMdi.m_cMenuPrincipal.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuParam.Count
    frmMdi.m_cMenuParam.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuImob.Count
    frmMdi.m_cMenuImob.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuMob.Count
    frmMdi.m_cMenuMob.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuAtende.Count
    frmMdi.m_cMenuAtende.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuTrib.Count
    frmMdi.m_cMenuTrib.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuProt.Count
    frmMdi.m_cMenuProt.Enabled(x) = False
Next
For x = 1 To frmMdi.m_cMenuOutro.Count
    frmMdi.m_cMenuOutro.Enabled(x) = False
Next

frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrinter")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuSelectPrinter")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuChangeUser")) = True
frmMdi.m_cMenuPrincipal.Enabled(frmMdi.m_cMenuPrincipal.IndexForKey("mnuClose")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTerritorial")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuTributário")) = True
frmMdi.m_cMenuParam.Enabled(frmMdi.m_cMenuParam.IndexForKey("mnuOutrosParam")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuCadastro")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuConsultaImovel")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuAtividadeImovel")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuRelatorioImob")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuGTIRural")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuRural")) = True
frmMdi.m_cMenuImob.Enabled(frmMdi.m_cMenuImob.IndexForKey("mnuOutroImob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadastroMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAtividadeMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuCadastroMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAtividadeMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAtividadeCadMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuConsultaMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuProdutividade")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuRelatorioMob")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuAlvaras")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuDevedores")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuSN")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnuListaAtiv")) = True
frmMdi.m_cMenuMob.Enabled(frmMdi.m_cMenuMob.IndexForKey("mnurelParcNPago")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuSenha")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuParcelamentoDivida")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuRelatorioAte")) = True
frmMdi.m_cMenuAtende.Enabled(frmMdi.m_cMenuAtende.IndexForKey("mnuConsultaAte")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuCalcTit")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuAtivBanco")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuSimples")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuAlugueis")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuDividaAtivaT")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuOutrosT")) = True
frmMdi.m_cMenuTrib.Enabled(frmMdi.m_cMenuTrib.IndexForKey("mnuRelatorioTrib")) = True
frmMdi.m_cMenuProt.Enabled(frmMdi.m_cMenuProt.IndexForKey("mnuParametroProt")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuAdmin")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuSecretariaObra")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuParamObra")) = True
frmMdi.m_cMenuOutro.Enabled(frmMdi.m_cMenuOutro.IndexForKey("mnuSeguranca")) = True

End Sub

Private Sub GetLanguage()
Dim Sql As String, RdoAux As rdoResultset

Sql = "select @@language as idioma"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux!idioma = "us_english" Then
    sDataFormat = "mm/dd/yyyy"
Else
    sDataFormat = "dd/mm/yyyy"
End If
RdoAux.Close


End Sub
