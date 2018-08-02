VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmParamObra 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   7005
   ClientTop       =   5370
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   6045
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   5955
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   2745
      TabIndex        =   1
      ToolTipText     =   "Novo Registro"
      Top             =   2565
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
      MICON           =   "frmParamObra.frx":0000
      PICN            =   "frmParamObra.frx":001C
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
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   2565
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
      MICON           =   "frmParamObra.frx":0176
      PICN            =   "frmParamObra.frx":0192
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
      Left            =   4950
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   2565
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&xcluir"
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
      MICON           =   "frmParamObra.frx":02EC
      PICN            =   "frmParamObra.frx":0308
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmParamObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sigla As String, sCaption As String

Public Property Let sSigla(sSiglaForm As String)
    Sigla = sSiglaForm
End Property

Private Sub cmdAlterar_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCod As Integer

If List1.ListIndex = -1 Then
    MsgBox "Selecione um Registro!", vbCritical, "Atenção"
    Exit Sub
End If

Inicio:
z = InputBox("Digite o novo nome.", "Alteração de Registro", List1.Text)
If z <> "" Then
    z = Left(z, 100)
    Sql = "SELECT * FROM PARAMOBRA WHERE SIGLA='" & Sigla & "' AND NOME='" & UCase(z) & "' AND CODIGO<>" & List1.ItemData(List1.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Nome já cadastrado!", vbCritical, "Atenção"
           .Close
            GoTo Inicio
        End If
       .Close
    End With
    Sql = "UPDATE PARAMOBRA SET NOME='" & UCase(Mask(CStr(z))) & "' WHERE SIGLA='" & Sigla & "' AND CODIGO=" & List1.ItemData(List1.ListIndex)
    cn.Execute Sql, rdExecDirect
    CarregaLista
End If

End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String
On Error GoTo Erro
If List1.ListIndex = -1 Then
    MsgBox "Selecione um Registro!", vbCritical, "Atenção"
Else
    If MsgBox("Deseja excluir este Registro?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "DELETE FROM PARAMOBRA WHERE SIGLA='" & Sigla & "' AND CODIGO=" & List1.ItemData(List1.ListIndex)
        cn.Execute Sql, rdExecDirect
        CarregaLista
    End If
End If

Exit Sub
Erro:
MsgBox "Não é possível excluir este registro, pois seu nome esta ligado a  outras partes do sistema.", vbCritical, "Atenção"

End Sub

Private Sub cmdNovo_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCod As Integer
Inicio:
z = InputBox("Digite o nome do novo registro.", "Inclusão de Registro", z)
If z <> "" Then
    z = Left(z, 100)
    Sql = "SELECT * FROM PARAMOBRA WHERE SIGLA='" & Sigla & "' AND NOME='" & UCase(z) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Nome já cadastrado!", vbCritical, "Atenção"
           .Close
            GoTo Inicio
        End If
       .Close
    End With
    Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM PARAMOBRA WHERE SIGLA='" & Sigla & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!MAXIMO) Then
            nCod = 1
        Else
            nCod = !MAXIMO + 1
        End If
        Sql = "INSERT PARAMOBRA(SIGLA,CODIGO,NOME) VALUES('" & Sigla & "'," & nCod & ",'" & UCase(Mask(CStr(z))) & "')"
        cn.Execute Sql, rdExecDirect
        List1.AddItem UCase(z)
        List1.ItemData(List1.NewIndex) = nCod
    End With
End If

End Sub

Private Sub Form_Load()
Select Case Sigla
    Case "EQ"
        Me.Caption = "Tabela de Equipes"
    Case "FC"
        Me.Caption = "Tabela de Chefes de Equipe"
    Case "AS"
        Me.Caption = "Tabela de Assuntos"
    Case "AT"
        Me.Caption = "Tabela de Atendentes"
    Case "TA"
        Me.Caption = "Tabela de Tipo de Atendimento"
End Select
Centraliza Me
CarregaLista

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset
List1.Clear

Sql = "SELECT * FROM PARAMOBRA WHERE SIGLA='" & Sigla & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        List1.AddItem !NOME
        List1.ItemData(List1.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

If List1.ListCount > 0 Then List1.ListIndex = 0

End Sub
