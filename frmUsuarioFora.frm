VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmUsuarioFora 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuários de Fora"
   ClientHeight    =   3630
   ClientLeft      =   1080
   ClientTop       =   3180
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7875
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6720
      TabIndex        =   4
      ToolTipText     =   "Cancelar Edição"
      Top             =   3240
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
      MICON           =   "frmUsuarioFora.frx":0000
      PICN            =   "frmUsuarioFora.frx":001C
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
      Left            =   1170
      TabIndex        =   5
      ToolTipText     =   "Atribuir funcionários a um usuário"
      Top             =   3240
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Atribuir"
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
      MICON           =   "frmUsuarioFora.frx":0176
      PICN            =   "frmUsuarioFora.frx":0192
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
      Left            =   2250
      TabIndex        =   6
      ToolTipText     =   "Remover um funcionário"
      Top             =   3240
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
      MICON           =   "frmUsuarioFora.frx":02EC
      PICN            =   "frmUsuarioFora.frx":0308
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
      Left            =   5640
      TabIndex        =   7
      ToolTipText     =   "Gravar os Dados"
      Top             =   3240
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
      MICON           =   "frmUsuarioFora.frx":03AA
      PICN            =   "frmUsuarioFora.frx":03C6
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
      TabIndex        =   8
      ToolTipText     =   "Incluir um novo funcionário"
      Top             =   3240
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
      MICON           =   "frmUsuarioFora.frx":076B
      PICN            =   "frmUsuarioFora.frx":0787
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
      Left            =   6690
      TabIndex        =   9
      ToolTipText     =   "Sair da Tela"
      Top             =   3240
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
      MICON           =   "frmUsuarioFora.frx":08E1
      PICN            =   "frmUsuarioFora.frx":08FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lst2 
      Appearance      =   0  'Flat
      Height          =   2730
      Left            =   3990
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.ListBox lst1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   30
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   330
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione os Funcionários de Fora"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4020
      TabIndex        =   3
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione um Usuário do Sistema"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   2895
   End
End
Attribute VB_Name = "frmUsuarioFora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tpUsuario
    id As Integer
    NomeLogin As String
    NomeCompleto As String
End Type
Dim RdoAux As rdoResultset, Sql As String, bExec As Boolean
Dim Evento As String, aUser() As tpUsuario, aFunc() As tpUsuario

Private Sub cmdAlterar_Click()
Eventos "INCLUIR"
Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
Evento = ""
End Sub

Private Sub cmdExcluir_Click()
Dim sLogin As String
If lst2.ListIndex = -1 Then
    MsgBox "Selecione um funcionário.", vbExclamation, "Atenção"
    Exit Sub
End If
sLogin = RetornaLogin(lst2.Text, "F")

Sql = "SELECT ANO,NUMERO FROM PROCESSOGTI WHERE RESPONSAVEL='" & sLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Não é possível excluir este funcionário pois existem processos atribuidos a ele.", vbExclamation, "Atenção"
       .Close
        Exit Sub
    Else
        If MsgBox("Excluir o funcionário " & lst2.Text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            Sql = "DELETE FROM USUARIOFUNC WHERE FUNCLOGIN='" & sLogin & "'"
            cn.Execute Sql, rdExecDirect
            Sql = "DELETE FROM USUARIO WHERE NOMELOGIN='" & sLogin & "'"
            cn.Execute Sql, rdExecDirect
            lst2.RemoveItem (lst2.ListIndex)
        End If
       .Close
    End If
End With

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer
'Sql = "DELETE FROM USUARIOFUNC WHERE USERLOGIN='" & RetornaLogin(lst1.Text, "U") & "'"
Sql = "DELETE FROM USUARIOFUNC WHERE USERid=" & lst1.ItemData(lst1.ListIndex)
cn.Execute Sql, rdExecDirect


For x = 0 To lst2.ListCount - 1
    If lst2.Selected(x) = True Then
'        Sql = "INSERT USUARIOFUNC(userid,USERLOGIN,FUNCLOGIN) VALUES(" & lst1.ItemData(lst1.ListIndex) & ",'"
'        Sql = Sql & RetornaLogin(lst1.Text, "U") & "','" & RetornaLogin(lst2.List(x), "F") & "')"
        Sql = "INSERT USUARIOFUNC(userid,FUNCLOGIN) VALUES(" & lst1.ItemData(lst1.ListIndex) & ",'" & RetornaLogin(lst2.List(x), "F") & "')"
        cn.Execute Sql, rdExecDirect
    End If
Next

Eventos "INICIAR"
Evento = ""
End Sub

Private Sub cmdNovo_Click()
Dim z As Variant, n As Integer, nCodigo As Integer



z = InputBox("Digite o nome do funcionário que reberá o processo.", "Incluir Novo Funcionário")
If Trim$(z) = "" Then Exit Sub

Sql = "SELECT NOMECOMPLETO FROM USUARIO WHERE NOMECOMPLETO='" & UCase$(z) & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Funcionário/Usuário já cadastrado.", vbCritical, "Atenção"
        .Close
        Exit Sub
    End If
End With

Sql = "SELECT NOMELOGIN From Usuario WHERE ((SUBSTRING(NOMELOGIN,1,2)='F0') OR (SUBSTRING(nomelogin, 1, 2) = 'F1' ) ) ORDER BY NOMELOGIN DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        n = Val(Right$(!NomeLogin, 3)) + 1
    Else
        n = 1
    End If
   .Close
End With

Sql = "select max(id)as maximo from usuario"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nCodigo = RdoAux!maximo + 1
RdoAux.Close


Sql = "INSERT USUARIO (id,NOMELOGIN,NOMECOMPLETO,externo) VALUES(" & nCodigo & ",'"
Sql = Sql & "F" & Format(n, "000") & "','" & Mask(CStr(UCase$(z))) & "',1)"
cn.Execute Sql, rdExecDirect
lst2.AddItem z
ReDim Preserve aFunc(UBound(aFunc) + 1)
aFunc(UBound(aFunc)).NomeLogin = "F" & Format(n, "000")
aFunc(UBound(aFunc)).NomeCompleto = z

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
bExec = True
Centraliza Me
CarregaUsuario
Evento = ""
Eventos "INICIAR"
End Sub

Private Sub CarregaUsuario()
ReDim aUser(0): ReDim aFunc(0)

'Sql = "SELECT NOMELOGIN,NOMECOMPLETO FROM USUARIO WHERE SUBSTRING(NOMELOGIN,1,2) <> 'F0' and SUBSTRING(NOMELOGIN,1,2) <> 'F1'"
Sql = "SELECT ID,NOMELOGIN,NOMECOMPLETO FROM USUARIO WHERE externo=0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aUser(UBound(aUser) + 1)
        aUser(UBound(aUser)).id = !id
        aUser(UBound(aUser)).NomeLogin = !NomeLogin
        aUser(UBound(aUser)).NomeCompleto = !NomeCompleto
        lst1.AddItem !NomeCompleto
        lst1.ItemData(lst1.NewIndex) = !id
       .MoveNext
    Loop
   .Close
End With

'Sql = "SELECT NOMELOGIN,NOMECOMPLETO FROM USUARIO WHERE SUBSTRING(NOMELOGIN,1,2) = 'F0' or SUBSTRING(NOMELOGIN,1,2) = 'F1'"
Sql = "SELECT id,NOMELOGIN,NOMECOMPLETO FROM USUARIO WHERE externo=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aFunc(UBound(aFunc) + 1)
        aFunc(UBound(aFunc)).id = !id
        aFunc(UBound(aFunc)).NomeLogin = !NomeLogin
        aFunc(UBound(aFunc)).NomeCompleto = !NomeCompleto
        lst2.AddItem !NomeCompleto
        lst2.ItemData(lst2.NewIndex) = !id
       .MoveNext
    Loop
   .Close
End With
If lst1.ListCount > 0 Then lst1.ListIndex = 0

End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   
End If

End Sub

Private Sub lst1_Click()
Dim sNome As String, x As Integer
bExec = False
For x = 0 To lst2.ListCount - 1
    lst2.Selected(x) = False
Next

If Evento = "" Then
    Sql = "SELECT USUARIO.NOMECOMPLETO FROM usuariofunc INNER JOIN "
    Sql = Sql & "USUARIO ON usuariofunc.funclogin = USUARIO.NOMELOGIN "
    Sql = Sql & "WHERE usuariofunc.userid = " & lst1.ItemData(lst1.ListIndex)
    'Sql = Sql & "WHERE usuariofunc.userlogin = '" & RetornaLogin(lst1.Text, "U") & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            sNome = !NomeCompleto
            For x = 0 To lst2.ListCount
                If lst2.List(x) = sNome Then
                    lst2.Selected(x) = True
                    Exit For
                End If
            Next
           .MoveNext
        Loop
    End With
    
End If
bExec = True

End Sub

Private Sub lst2_ItemCheck(Item As Integer)
If Evento = "" And bExec Then
    lst2.Selected(Item) = Not lst2.Selected(Item)
End If
End Sub




Private Function RetornaLogin(sNomeCompleto As String, sTipo As String) As String
Dim x As Integer

If sTipo = "U" Then
    For x = 1 To UBound(aUser)
        If aUser(x).NomeCompleto = sNomeCompleto Then
            RetornaLogin = aUser(x).NomeLogin
            Exit Function
        End If
    Next
Else
    For x = 1 To UBound(aFunc)
        If aFunc(x).NomeCompleto = sNomeCompleto Then
            RetornaLogin = aFunc(x).NomeLogin
            Exit Function
        End If
    Next
End If

End Function

