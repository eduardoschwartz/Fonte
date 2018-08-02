VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSecurityByUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segurança por Usuário"
   ClientHeight    =   3030
   ClientLeft      =   5775
   ClientTop       =   2700
   ClientWidth     =   10620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   10620
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Importar"
      Height          =   345
      Left            =   90
      TabIndex        =   5
      Top             =   2550
      Width           =   1095
   End
   Begin VB.ListBox lstDetail 
      Height          =   1860
      Left            =   5370
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   600
      Width           =   5205
   End
   Begin VB.ListBox lstMaster 
      Height          =   1815
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   5145
   End
   Begin VB.ComboBox cmbUser 
      Height          =   315
      ItemData        =   "frmSecurityByUser.frx":0000
      Left            =   870
      List            =   "frmSecurityByUser.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   5655
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   5460
      TabIndex        =   4
      ToolTipText     =   "Gravar os Dados"
      Top             =   2610
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
      MICON           =   "frmSecurityByUser.frx":0004
      PICN            =   "frmSecurityByUser.frx":0020
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
      Caption         =   "Usuário.:"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "frmSecurityByUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Acesso
    id As Integer
    idMaster As Integer
    descricao As String
    Valor As Boolean
End Type
Private Type aUser
    Usuario As String
    Acesso As String
End Type

Private aUser() As aUser
Private aAcesso() As Acesso
Dim bExec As Boolean

Private Sub cmbUser_Click()
Dim sLoginName As String

If cmbUser.ListIndex = -1 Then Exit Sub
If Not bExec Then Exit Sub

sLoginName = RetrieveLoginName(cmbUser.Text)
LoadSecurity sLoginName

End Sub

Private Sub cmdGravar_Click()
Dim sLoginName As String
sLoginName = RetrieveLoginName(cmbUser.Text)
SaveSecurity sLoginName
MsgBox "gravado"
End Sub

Private Sub Command1_Click()
Dim RdoAux As rdoResultset, Sql As String, nPos As Integer, RdoAux2 As rdoResultset, sID As String
Dim sNome As String, ax As String, ax2 As String, x As Integer, bFind As Boolean, nId As Integer, y As Integer
ReDim aUser(0)

Sql = "select count(*) as contador from sec_item"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nPos = RdoAux!contador
RdoAux.Close

ax = String$(nPos, "0")

Sql = "select nomelogin from usuario where ativo=1 order by nomelogin"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aUser(UBound(aUser) + 1)
        aUser(UBound(aUser)).Usuario = !NomeLogin
        aUser(UBound(aUser)).Acesso = ax
       .MoveNext
    Loop
   .Close
End With

Sql = "delete from sec_user_item"
cn.Execute Sql, rdExecDirect

Sql = "select * from sec_item order by id"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nId = !id
        If IsNull(!codtela) Then
        Else
            Sql = "select nomeusuario FROM seg_useracess INNER JOIN usuario ON seg_useracess.nomeusuario = usuario.nomelogin "
            Sql = Sql & " WHERE usuario.ativo = 1 and codtela=" & !codtela & " and codevento=" & !codevento & " order by nomeusuario"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                Do Until .EOF
                    sNome = UCase(!nomeusuario)
                    bFind = False
                    For y = 1 To UBound(aUser)
                        If aUser(y).Usuario = sNome Then
                            x = y
                            bFind = True
                            Exit For
                        End If
                    Next
                    If bFind Then
                        ax = aUser(x).Acesso
                        ax2 = SetBit(ax, nId, 1)
                    End If
                    aUser(y).Acesso = ax2
                   .MoveNext
                Loop
               .Close
            End With
        End If
       .MoveNext
    Loop
   .Close
End With



For x = 1 To UBound(aUser)
    Sql = "insert sec_user_item(usuario,id) values('" & aUser(x).Usuario & "','" & aUser(x).Acesso & "')"
    cn.Execute Sql, rdExecDirect
Next

ax = ""
For x = 1 To UBound(aAcesso)
    ax = ax & "1"
Next
Sql = "update sec_user_item set id='" & ax & "' where id='SCHWARTZ'"
cn.Execute Sql, rdExecDirect

Sql = "select * from sec_item_out order by id"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For y = 1 To UBound(aUser)
            If aUser(y).Usuario = !Usuario Then
                ax = aUser(y).Acesso
                ax2 = SetBit(ax, !id, 1)
                aUser(y).Acesso = ax2
                
                Sql = "update sec_user_item set id='" & aUser(y).Acesso & "' where usuario='" & !Usuario & "'"
                cn.Execute Sql, rdExecDirect
                
                Exit For
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

'***
Sql = "select * from sec_item_cod"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        Sql = "select * from sec_user_item"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                sID = !id
                If BitId(RdoAux!codold) = 1 Then
                    ax2 = SetBit(sID, RdoAux!codnew, 1)
                    Sql = "update sec_user_item set id='" & ax2 & "' where usuario='" & !Usuario & "'"
                    cn.Execute Sql, rdExecDirect
                End If
               .MoveNext
            Loop
           .Close
        End With
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"
End Sub

Public Function SetBit(OldString As String, nPos As Integer, nValor As Integer)
Dim NewPos As Integer, ax As String

For NewPos = 1 To Len(OldString)
    If NewPos = nPos Then
        ax = ax & nValor
    Else
        ax = ax & Mid(OldString, NewPos, 1)
    End If
Next
SetBit = ax

End Function


Private Sub Form_Load()
bExec = False
Centraliza Me
CarregaUsuario
LoadMatrix
'CarregaMaster
bExec = True
cmbUser_Click
End Sub

Private Sub LoadMatrix()
Dim RdoAux As rdoResultset, Sql As String, nPos As Integer
ReDim aAcesso(0)

Sql = "select id,idmaster,descricao from sec_item order by id"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aAcesso(UBound(aAcesso) + 1)
        nPos = UBound(aAcesso)
        aAcesso(nPos).id = !id
        aAcesso(nPos).idMaster = !idMaster
        aAcesso(nPos).descricao = !descricao
        aAcesso(nPos).Valor = False
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub CarregaUsuario()
Dim RdoAux As rdoResultset, Sql As String
Sql = "select nomelogin,nomecompleto from usuario where ativo=1 order by nomecompleto"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbUser.AddItem !NomeCompleto & " (" & !NomeLogin & ")"
       .MoveNext
    Loop
   .Close
End With
cmbUser.ListIndex = 0
End Sub

Private Sub CarregaMaster()
Dim x As Integer

For x = 1 To UBound(aAcesso)
    With aAcesso(x)
        If .id = .idMaster Then
            lstMaster.AddItem .descricao
            lstMaster.ItemData(lstMaster.NewIndex) = .id
        End If
    End With
Next
End Sub

Private Sub CarregaDetail(CodMaster As Integer)
Dim x As Integer
lstDetail.Clear
For x = 1 To UBound(aAcesso)
    With aAcesso(x)
        If .idMaster = CodMaster And .id <> CodMaster Then
            lstDetail.AddItem .descricao
            lstDetail.ItemData(lstDetail.NewIndex) = .id
            lstDetail.Selected(lstDetail.ListCount - 1) = .Valor
        ElseIf .idMaster = CodMaster And .id = CodMaster Then
            lstDetail.AddItem "Visualizar/Executar"
            lstDetail.ItemData(lstDetail.NewIndex) = .id
            lstDetail.Selected(lstDetail.ListCount - 1) = .Valor
        End If
    End With
Next


End Sub

Private Sub lstDetail_ItemCheck(Item As Integer)
Dim nId As Integer, x As Integer, bFind As Boolean

nId = lstDetail.ItemData(lstDetail.ListIndex)
For x = 1 To UBound(aAcesso)
    If aAcesso(x).id = nId Then
        aAcesso(x).Valor = lstDetail.Selected(Item)
        Exit For
    End If
Next

bFind = False
For x = 0 To lstDetail.ListCount - 1
    If lstDetail.Selected(x) = True Then
        bFind = True
        Exit For
    End If
Next
nId = lstMaster.ItemData(lstMaster.ListIndex)
For x = 1 To UBound(aAcesso)
    If aAcesso(x).id = nId Then
        aAcesso(x).Valor = bFind
        Exit For
    End If
Next

If bFind And lstDetail.Selected(0) = False Then
    lstDetail.Selected(0) = True
End If


End Sub

Private Sub lstMaster_Click()
CarregaDetail lstMaster.ItemData(lstMaster.ListIndex)
End Sub

Private Sub LoadSecurity(sUsuario As String)
Dim Sql As String, x As Integer, y As Integer, ax As String, RdoAux As rdoResultset

lstMaster.Clear: lstDetail.Clear
CarregaMaster

For x = 1 To UBound(aAcesso)
    aAcesso(x).Valor = False
Next

Sql = "select * from sec_user_item where usuario='" & sUsuario & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        ax = !id
    Else
        ax = "0"
    End If
   .Close
End With

For x = 1 To Len(ax)
    For y = 1 To UBound(aAcesso)
        If x = aAcesso(y).id Then
            aAcesso(y).Valor = Val(Mid(ax, x, 1))
            Exit For
        End If
    Next
Next

lstMaster.ListIndex = 0
End Sub

Private Sub SaveSecurity(sUsuario As String)
Dim Sql As String, x As Integer, ax As String, RdoAux As rdoResultset

ax = ""
For x = 1 To UBound(aAcesso)
    If aAcesso(x).Valor = True Then
        ax = ax & "1"
    Else
        ax = ax & "0"
    End If
Next

Sql = "select * from sec_user_item where usuario='" & sUsuario & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount > 0 Then
    Sql = "update sec_user_item set id='" & ax & "' where usuario='" & sUsuario & "'"
Else
    Sql = "insert sec_user_item(usuario,id) values('" & sUsuario & "','" & ax & "')"
End If
cn.Execute Sql, rdExecDirect

RdoAux.Close

End Sub

Private Function RetrieveLoginName(Usuario As String)
Dim nPosIni As Integer

nPosIni = InStr(1, Usuario, "(", vbTextCompare)
nPosFim = InStr(1, Usuario, ")", vbTextCompare)
RetrieveLoginName = Mid(Usuario, nPosIni + 1, nPosFim - nPosIni - 1)

End Function

