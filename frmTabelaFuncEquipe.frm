VERSION 5.00
Begin VB.Form frmTabelaFuncEquipe 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabela de Funcionários por Equipe"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   8325
   Begin VB.ListBox lstFunc 
      Appearance      =   0  'Flat
      Height          =   2505
      Left            =   90
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   945
      Width           =   8115
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7965
      TabIndex        =   2
      ToolTipText     =   "Atualizar Lista"
      Top             =   180
      Width           =   240
   End
   Begin VB.ComboBox cmbEq 
      Height          =   315
      Left            =   1485
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   180
      Width           =   6405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione os funcionários que participam da Equipe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   4
      Top             =   675
      Width           =   4515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome da Equipe:"
      Height          =   240
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   1275
   End
End
Attribute VB_Name = "frmTabelaFuncEquipe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bExec As Boolean, nCodEquipe As Integer

Private Sub cmbEq_Click()
Dim x As Integer, Sql As String, RdoAux As rdoResultset

Grava
nCodEquipe = cmbEq.ItemData(cmbEq.ListIndex)
For x = 0 To lstFunc.ListCount - 1
    lstFunc.Selected(x) = False
Next

Sql = "SELECT paramobra.nome, paramobra.codigo, equipefuncionario.codequipe "
Sql = Sql & "FROM equipefuncionario INNER JOIN paramobra ON equipefuncionario.codfunc = paramobra.codigo "
Sql = Sql & " where codequipe=" & cmbEq.ItemData(cmbEq.ListIndex) & " and sigla='FC'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        For x = 0 To lstFunc.ListCount - 1
            If !Codigo = lstFunc.ItemData(x) Then
                lstFunc.Selected(x) = True
                Exit For
            End If
        Next
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdRefresh_Click()
CarregaEquipe
End Sub

Private Sub Form_Load()
Centraliza Me
bExec = True
nCodEquipe = 0
CarregaEquipe
CarregaLista
End Sub

Private Sub CarregaEquipe()
Dim Sql As String, RdoAux As rdoResultset

bExec = False
cmbEq.Clear

Sql = "SELECT CODIGO,NOME FROM PARAMOBRA WHERE SIGLA='EQ'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbEq.AddItem !NOME
        cmbEq.ItemData(cmbEq.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
bExec = True

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset

If Not bExec Then Exit Sub
lstFunc.Clear
Sql = "SELECT nome, codigo FROM PARAMOBRA WHERE SIGLA='FC'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        lstFunc.AddItem !NOME
        lstFunc.ItemData(lstFunc.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Grava()
Dim Sql As String, x As Integer
If nCodEquipe = 0 Then Exit Sub

Sql = "DELETE FROM EQUIPEFUNCIONARIO WHERE CODEQUIPE=" & nCodEquipe
cn.Execute Sql, rdExecDirect

With lstFunc
    For x = 0 To .ListCount - 1
        If .Selected(x) Then
            Sql = "INSERT EQUIPEFUNCIONARIO (CODEQUIPE,CODFUNC) VALUES(" & nCodEquipe & "," & .ItemData(x) & ")"
            cn.Execute Sql, rdExecDirect
        End If
    Next
End With


End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Grava
End Sub
