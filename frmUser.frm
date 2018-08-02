VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmUser 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Usuários do Sistema"
   ClientHeight    =   3705
   ClientLeft      =   2595
   ClientTop       =   2625
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3705
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbGrupo 
      Height          =   315
      ItemData        =   "frmUser.frx":0000
      Left            =   900
      List            =   "frmUser.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   2895
   End
   Begin VB.ListBox lstRole 
      Height          =   1035
      Left            =   2010
      TabIndex        =   2
      Top             =   1620
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.ListBox lstUser 
      BackColor       =   &H00C0E0FF&
      Height          =   3180
      Left            =   0
      TabIndex        =   1
      Top             =   510
      Width           =   4005
   End
   Begin MSComctlLib.ImageList imlSecurity 
      Left            =   5910
      Top             =   4050
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":0062
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   3705
      Left            =   4050
      TabIndex        =   0
      Top             =   0
      Width           =   3915
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         BorderStyle     =   0  'None
         Height          =   1125
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2070
         Width           =   3435
      End
      Begin VB.TextBox txtUser 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   570
         Width           =   3615
      End
      Begin prjChameleon.chameleonButton cmdAddUser 
         Height          =   315
         Left            =   510
         TabIndex        =   9
         ToolTipText     =   "Adiciona Usuário"
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUser.frx":00C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdHelp 
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         ToolTipText     =   "Ajuda desta Tela"
         Top             =   960
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUser.frx":00DC
         PICN            =   "frmUser.frx":00F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRemoveUser 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   930
         TabIndex        =   11
         ToolTipText     =   "Remove Usuário"
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUser.frx":0252
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
         Left            =   1770
         TabIndex        =   8
         ToolTipText     =   "Sair da Tela"
         Top             =   960
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmUser.frx":026E
         PICN            =   "frmUser.frx":028A
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
         Caption         =   "Nome do Usuário"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   2115
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Grupos:"
      Height          =   225
      Left            =   90
      TabIndex        =   6
      Top             =   150
      Width           =   675
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbGrupo_Click()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset

Me.MousePointer = vbHourglass
lstUser.Clear

Sql = "SELECT MEMBERUID FROM SYSMEMBERS WHERE GROUPUID=" & cmbGrupo.ItemData(cmbGrupo.ListIndex)
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
     Do Until .EOF
          Sql = "SELECT NAME FROM VWUSUARIO WHERE UID=" & !MEMBERUID
          Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          With RdoAux3
                 lstUser.AddItem !Name
                .Close
          End With
         .MoveNext
     Loop
    .Close
End With

Me.MousePointer = vbDefault
If lstUser.ListCount > 0 Then lstUser.ListIndex = 0

End Sub

Private Sub cmdAddUser_Click()

If InStr(1, UCase$(cn.Connect), "DEVELOPER", vbBinaryCompare) > 0 Then
    MsgBox "Não é possivel criar usuários porque a ODBC esta configurada para acessar a base de Testes TributacaoDeveloper", vbCritical, "atenção"
    Exit Sub
End If

If Trim$(txtUser.text) = "" Then
     MsgBox "Digite o nome do novo usuário.", vbExclamation, "Atenção"
     Exit Sub
End If

For X = 0 To lstUser.ListCount - 1
    lstUser.ListIndex = X
    If UCase$(lstUser.text) = UCase$(Trim$(txtUser.text)) Then
         MsgBox "Este usuário já pertence a este grupo.", vbExclamation, "Atenção"
         Exit Sub
    End If
Next

If MsgBox("Adicionar o usuário " & txtUser.text & " ao grupo " & cmbGrupo.text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
     If Not oSQLServer.IsLogin(txtUser.text) Then
        'adiciona o usuario a base de dados tributacao
         Sql = "sp_addlogin '" & txtUser.text & "','NEWUSER','Tributacao'"
         cn.Execute Sql, rdExecDirect
     End If
     If Not oSQLServer.Databases("tributacao").IsUser(txtUser.text) Then
        'garante acesso do usuario a base de dados
         Sql = "sp_adduser '" & txtUser.text & "','" & txtUser.text & "','" & cmbGrupo.text & "'"
         cn.Execute Sql, rdExecDirect
     End If
     'adiciona o usuário ao grupo
     Sql = "sp_addrolemember '" & cmbGrupo.text & "','" & txtUser.text & "'"
     cn.Execute Sql, rdExecDirect
     
     DefaultAccess txtUser.text
     'adiciona a Lista
     txtUser.text = ""
     cmbGrupo_Click
End If

End Sub

Private Sub cmdRemoveUser_Click()

If InStr(1, UCase$(cn.Connect), "DEVELOPER", vbBinaryCompare) > 0 Then
    MsgBox "Não é possivel Remover usuários porque a ODBC esta configurada para acessar a base de Testes TributacaoDeveloper", vbCritical, "atenção"
    Exit Sub
End If

If lstUser.ListIndex = -1 Then
     MsgBox "Selecione o usuário a ser excluido.", vbExclamation, "Atenção"
     Exit Sub
End If

If MsgBox("Remover o usuário " & lstUser.text & " do Grupo " & cmbGrupo.text & "?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    'remove o usuario do grupo
    Sql = "sp_droprolemember '" & cmbGrupo.text & "','" & lstUser.text & "'"
    cn.Execute Sql, rdExecDirect
    cmbGrupo_Click
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim sTabela As String, aTabela() As String, nNumCampo As Integer

Ocupado
Screen.MousePointer = vbHourglass

Centraliza Me
'On Error GoTo Erro
Set oSQLServer = New SQLDMO.SQLServer
oSQLServer.LoginTimeout = 20

With oSQLServer
    .Disconnect
    .LoginTimeout = 20
    .Connect "DBSERVER", LastUser, UserPwd
    ReDim aTabela(oSQLServer.Databases("Tributacao").Tables.Count)
    For x22 = 1 To oSQLServer.Databases("Tributacao").Tables.Count
        aTabela(x22) = oSQLServer.Databases("Tributacao").Tables(x22).Name
    Next


    For x22 = 1 To UBound(aTabela)
        sTabela = aTabela(x22)
        If Left$(sTabela, 3) <> "sys" And Left$(sTabela, 2) <> "dt" Then
        oSQLServer.Databases("Tributacao").Tables(sTabela).Name = LCase$(oSQLServer.Databases("Tributacao").Tables(sTabela).Name)
        nNumCampo = oSQLServer.Databases("Tributacao").Tables(sTabela).Columns.Count
        For z = 1 To nNumCampo
'            DoEvents
            oSQLServer.Databases("Tributacao").Tables(sTabela).Columns(z).Name = LCase$(oSQLServer.Databases("Tributacao").Tables(sTabela).Columns(z).Name) & "@"
            oSQLServer.Databases("Tributacao").Tables(sTabela).Columns(z).Name = Left(oSQLServer.Databases("Tributacao").Tables(sTabela).Columns(z).Name, Len(oSQLServer.Databases("Tributacao").Tables(sTabela).Columns(z).Name) - 1)
'            MsgBox oSQLServer.Databases("Tributacao").Tables(x).Name
        Next
        End If
    Next

End With
Carrega
Screen.MousePointer = vbDefault
Liberado

Exit Sub
Erro:

MsgBox Err.Description
Screen.MousePointer = vbDefault
'Resume Next
Liberado
End Sub

Private Sub Carrega()

Dim RdoAux As rdoResultset
   
Sql = "SELECT * FROM VWGRUPOS"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If UCase$(Left$(!Name, 2)) <> "DB" Then
               lstRole.AddItem !Name
               cmbGrupo.AddItem !Name
               cmbGrupo.ItemData(cmbGrupo.NewIndex) = !UID
        End If
       .MoveNext
    Loop
   .Close
End With
   
If cmbGrupo.ListCount > 0 Then cmbGrupo.ListIndex = 0
   
End Sub

Private Sub txtUser_LostFocus()
txtUser.text = UCase$(txtUser.text)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oSQLServer = Nothing
End Sub

Private Sub txtUser_GotFocus()
txtUser.SelStart = 0
txtUser.SelLength = Len(txtUser.text)
End Sub

