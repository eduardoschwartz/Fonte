VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEspolio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de Proprietário"
   ClientHeight    =   1305
   ClientLeft      =   5280
   ClientTop       =   4800
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4890
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   1620
      TabIndex        =   4
      Text            =   "cmbTipo"
      Top             =   720
      Width           =   1860
   End
   Begin VB.TextBox txtCod 
      Height          =   285
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   0
      Top             =   315
      Width           =   1095
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   3645
      TabIndex        =   2
      ToolTipText     =   "Gravar os Dados"
      Top             =   720
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
      MICON           =   "frmEspolio.frx":0000
      PICN            =   "frmEspolio.frx":001C
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
      Left            =   3630
      TabIndex        =   5
      ToolTipText     =   "Remover o tipo de proprietário"
      Top             =   300
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
      MICON           =   "frmEspolio.frx":03C1
      PICN            =   "frmEspolio.frx":03DD
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
      Caption         =   "Tipo proprietário:"
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   3
      Top             =   720
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "Código cidadão:"
      Height          =   240
      Index           =   0
      Left            =   270
      TabIndex        =   1
      Top             =   315
      Width           =   1230
   End
End
Attribute VB_Name = "frmEspolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExcluir_Click()
Dim nCod As Long, Sql As String

nCod = Val(txtCod.Text)
If MsgBox("Remover o tipo de proprietário?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
    Sql = "delete from espolio where codigo=" & nCod
    cn.Execute Sql, rdExecDirect
'    Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " HH:MM:SS") & "','" & NomeDeLogin & "','"
'    Sql = Sql & "O tipo de cidadão foi removido.')"
    Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " HH:MM:SS") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
    Sql = Sql & "O tipo de cidadão foi removido.')"
    cn.Execute Sql, rdExecDirect
End If

End Sub

Private Sub cmdGravar_Click()
Dim nCod As Long, Sql As String, RdoAux As rdoResultset

nCod = Val(txtCod.Text)
Sql = "select * from cidadao where codcidadao=" & nCod
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    MsgBox "Código não cadastrado!", vbCritical, "Erro"
    RdoAux.Close
Else
    RdoAux.Close
    Sql = "select * from espolio where codigo=" & nCod
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        If MsgBox("Cidadão já possui um tipo cadastrado em " & Format(RdoAux!Data, "dd/mm/yyyy") & " por " & RdoAux!Usuario & vbCrLf & "Você deseja alterar?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "update espolio set tipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & ",data='" & Format(Now, sDataFormat) & "',usuario='" & RetornaUsuarioFullName & "' where codigo=" & nCod
            cn.Execute Sql, rdExecDirect
'            Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'            Sql = Sql & "O tipo de cidadão foi alterado para " & cmbTipo.Text & "')"
            Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
            Sql = Sql & "O tipo de cidadão foi alterado para " & cmbTipo.Text & "')"
            cn.Execute Sql, rdExecDirect
        End If
        RdoAux.Close
    Else
        RdoAux.Close
        Sql = "insert espolio (codigo,tipo,data,usuario) values(" & nCod & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & ",'" & Format(Now, "mm/dd/yyyy") & "','" & RetornaUsuarioFullName & "')"
        cn.Execute Sql, rdExecDirect
        
'        Sql = "insert historicocidadao(codigo,data,usuario,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & NomeDeLogin & "','"
'        Sql = Sql & "O tipo de cidadão foi alterado para " & cmbTipo.Text & "')"
        Sql = "insert historicocidadao(codigo,data,userid,obs) values(" & nCod & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'"
        Sql = Sql & "O tipo de cidadão foi alterado para " & cmbTipo.Text & "')"
        cn.Execute Sql, rdExecDirect
        
        MsgBox "Tipo de proprietário cadastrado com sucesso.", vbInformation, "Atenção"
    End If
End If


End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

Sql = "select codigo,nome from tipousuario order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !Nome
        cmbTipo.ItemData(cmbTipo.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
cmbTipo.ListIndex = 0

End Sub


Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub
