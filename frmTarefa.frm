VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTarefa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Adicionar/Editar tarefa"
   ClientHeight    =   3000
   ClientLeft      =   10380
   ClientTop       =   6540
   ClientWidth     =   6810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTarefa 
      Appearance      =   0  'Flat
      Height          =   1185
      Left            =   90
      MaxLength       =   400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1290
      Width           =   6615
   End
   Begin VB.ComboBox cmbUser 
      Height          =   315
      Left            =   1860
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   4845
   End
   Begin MSComCtl2.DTPicker dtData 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   90046465
      CurrentDate     =   42836
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   345
      Left            =   5580
      TabIndex        =   6
      ToolTipText     =   "Cancelar operação"
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmTarefa.frx":0000
      PICN            =   "frmTarefa.frx":001C
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
      Height          =   345
      Left            =   4380
      TabIndex        =   7
      ToolTipText     =   "Gravar a tarefa"
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmTarefa.frx":0176
      PICN            =   "frmTarefa.frx":0192
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
      Height          =   345
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Excluir tarefa"
      Top             =   2580
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   609
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
      MICON           =   "frmTarefa.frx":0537
      PICN            =   "frmTarefa.frx":0553
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
      Caption         =   "Descrição da tarefa"
      Height          =   225
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   990
      Width           =   3045
   End
   Begin VB.Label Label1 
      Caption         =   "Criar tarefa para o usuário"
      Height          =   225
      Index           =   1
      Left            =   1890
      TabIndex        =   2
      Top             =   120
      Width           =   2085
   End
   Begin VB.Label Label1 
      Caption         =   "Data de Conclusão"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1545
   End
End
Attribute VB_Name = "frmTarefa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGravar_Click()
Dim sql As String, sNome As String

If cmbUser.ListIndex = -1 Then
    MsgBox "Selecione um usuário para receber a tarefa", vbCritical, "Atenção"
    Exit Sub
End If

If Trim(txtTarefa.Text) = "" Then
    MsgBox "Digite a descrição da tarefa", vbCritical, "Atenção"
    Exit Sub
End If

If Format(dtData.value, "dd/mm/yyyy") < Format(Now, "dd/mm/yyyy") Then
    MsgBox "A data de conclusão da tarefa não pode ser inferior a data atual", vbCritical, "Atenção"
    Exit Sub
End If

sNome = RetornaUsuarioLoginName(cmbUser.Text)

sql = "insert agenda(user_send,user_receive,data_inclusao,data_previsao,compromisso) values('"
sql = sql & NomeDeLogin & "','" & sNome & "','" & Format(Now, "mm/dd/yyyy") & "','" & Format(dtData.value, "mm/dd/yyyy") & "','" & Mask(txtTarefa.Text) & "')"
cn.Execute sql, rdExecDirect
frmAgenda.LoadAgenda
Unload Me

End Sub

Private Sub Form_Load()
Dim sql As String, RdoAux As rdoResultset

sql = "select nomecompleto from usuario where ativo=1 order by nomecompleto"
Set RdoAux = cn.OpenResultset(sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbUser.AddItem !NomeCompleto
       .MoveNext
    Loop
   .Close
End With

End Sub
