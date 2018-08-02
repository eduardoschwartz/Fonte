VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerenciador de grupos de chat"
   ClientHeight    =   4725
   ClientLeft      =   5895
   ClientTop       =   5190
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   6570
   Begin VB.ComboBox cmbGrupo 
      Height          =   315
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   5235
   End
   Begin VB.ListBox lstUser 
      Height          =   3660
      Left            =   45
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   6465
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5355
      TabIndex        =   2
      ToolTipText     =   "Sair da Tela"
      Top             =   4290
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
      MICON           =   "frmGrupo.frx":0000
      PICN            =   "frmGrupo.frx":001C
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
      Left            =   4275
      TabIndex        =   3
      ToolTipText     =   "Gravar os Dados"
      Top             =   4290
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
      MICON           =   "frmGrupo.frx":008A
      PICN            =   "frmGrupo.frx":00A6
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
      Caption         =   "Grupo.....:"
      Height          =   195
      Left            =   405
      TabIndex        =   4
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim sRet As String

Private Sub cmbGRUPO_Click()
CarregaUser
End Sub

Private Sub cmdGravar_Click()
Dim x As Integer

Sql = "DELETE FROM CHATGRUPOUSUARIO WHERE GRUPO=" & cmbGrupo.ItemData(cmbGrupo.ListIndex)
cn.Execute Sql, rdExecDirect

For x = 0 To lstUser.ListCount - 1
    If lstUser.Selected(x) = True Then
        lstUser.ListIndex = x
        Sql = "INSERT CHATGRUPOUSUARIO(GRUPO,NOME) VALUES("
        Sql = Sql & cmbGrupo.ItemData(cmbGrupo.ListIndex) & ",'" & lstUser.Text & "')"
        cn.Execute Sql, rdExecDirect
    End If
Next

MsgBox "Usuários gravados com sucesso.", vbInformation, "Informação"
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()
Centraliza Me
sRet = RetEventUserForm(Me.Name)
CarregaGrupo

End Sub

Private Sub CarregaGrupo()

Sql = "Select CODIGO,NOME From chatGRUPO ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

cmbGrupo.Clear
With RdoAux
   .MoveFirst
    Do Until .EOF
       cmbGrupo.AddItem !NOME
       cmbGrupo.ItemData(cmbGrupo.NewIndex) = !Codigo
      .MoveNext
    Loop
   .Close
End With
cmbGrupo.ListIndex = 0

Sql = "Select NOMELOGIN From USUARIO WHERE ATIVO=1 ORDER BY nomeLOGIN"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

lstUser.Clear
With RdoAux
   .MoveFirst
    Do Until .EOF
       lstUser.AddItem !NomeLogin
      .MoveNext
    Loop
   .Close
End With
cmbGRUPO_Click
End Sub

Private Sub CarregaUser()
Dim x As Integer

Sql = "SELECT USUARIO.NOMELOGIN FROM USUARIO INNER JOIN "
Sql = Sql & "CHATGRUPOUSUARIO ON USUARIO.NOMELOGIN = CHATGRUPOUSUARIO.NOME "
Sql = Sql & "WHERE CHATGRUPOUSUARIO.GRUPO =" & cmbGrupo.ItemData(cmbGrupo.ListIndex)
Sql = Sql & " ORDER BY NOMELOGIN"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

For x = 0 To lstUser.ListCount - 1
   lstUser.Selected(x) = False
Next

With RdoAux
    Do Until .EOF
       For x = 0 To lstUser.ListCount - 1
           If lstUser.List(x) = !NomeLogin Then
              lstUser.Selected(x) = True
              Exit For
           End If
       Next
      .MoveNext
    Loop
   .Close
End With

End Sub

