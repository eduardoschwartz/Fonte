VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTabelaAssunto 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de Assuntos"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   6030
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   15
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   5955
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   2715
      TabIndex        =   1
      ToolTipText     =   "Novo Registro"
      Top             =   2550
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
      MICON           =   "frmTabelaAssunto.frx":0000
      PICN            =   "frmTabelaAssunto.frx":001C
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
      Left            =   3810
      TabIndex        =   2
      ToolTipText     =   "Editar Registro"
      Top             =   2550
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
      MICON           =   "frmTabelaAssunto.frx":0176
      PICN            =   "frmTabelaAssunto.frx":0192
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
      Left            =   4920
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   2550
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
      MICON           =   "frmTabelaAssunto.frx":02EC
      PICN            =   "frmTabelaAssunto.frx":0308
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
Attribute VB_Name = "frmTabelaAssunto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlterar_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCod As Integer

If List1.ListIndex = -1 Then
    MsgBox "Selecione um Assunto!", vbCritical, "Atenção"
    Exit Sub
End If

Inicio:
z = InputBox("Digite o novo nome do Assunto.", "Alteração de Assunto", List1.text)
If z <> "" Then
    z = Left(z, 100)
    Sql = "SELECT * FROM ASSUNTOOBRA WHERE NOME='" & UCase(z) & "' AND CODIGO<>" & List1.ItemData(List1.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Assunto já cadastrado!", vbCritical, "Atenção"
           .Close
            GoTo Inicio
        End If
       .Close
    End With
    Sql = "UPDATE ASSUNTOOBRA SET NOME='" & UCase(Mask(CStr(z))) & "' WHERE CODIGO=" & List1.ItemData(List1.ListIndex)
    cn.Execute Sql, rdExecDirect
    CarregaLista
End If

End Sub

Private Sub cmdExcluir_Click()
Dim Sql As String
On Error GoTo Erro
If List1.ListIndex = -1 Then
    MsgBox "Selecione um Assunto!", vbCritical, "Atenção"
Else
    If MsgBox("Deseja excluir este Assunto?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        Sql = "DELETE FROM ASSUNTOOBRA WHERE CODIGO=" & List1.ItemData(List1.ListIndex)
        cn.Execute Sql, rdExecDirect
        CarregaLista
    End If
End If

Exit Sub
Erro:
MsgBox "Não é possível excluir este Assunto, pois seu nome esta ligado a  outras partes do sistema.", vbCritical, "Atenção"

End Sub

Private Sub cmdNovo_Click()
Dim z As Variant, Sql As String, RdoAux As rdoResultset, nCod As Integer
Inicio:
z = InputBox("Digite o nome do novo Assunto.", "Inclusão de Assunto", z)
If z <> "" Then
    z = Left(z, 100)
    Sql = "SELECT * FROM ASSUNTOOBRA WHERE NOME='" & UCase(z) & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Assunto já cadastrado!", vbCritical, "Atenção"
           .Close
            GoTo Inicio
        End If
       .Close
    End With
    Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM ASSUNTOOBRA"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If IsNull(!MAXIMO) Then
            nCod = 1
        Else
            nCod = !MAXIMO + 1
        End If
        Sql = "INSERT ASSUNTOOBRA(CODIGO,NOME) VALUES(" & nCod & ",'" & UCase(Mask(CStr(z))) & "')"
        cn.Execute Sql, rdExecDirect
        List1.AddItem UCase(z)
        List1.ItemData(List1.NewIndex) = nCod
    End With
End If

End Sub

Private Sub Form_Load()
Centraliza Me
CarregaLista

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset
List1.Clear
Sql = "SELECT * FROM ASSUNTOOBRA"
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


