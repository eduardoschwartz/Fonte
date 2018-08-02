VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmObsCidadao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observação de Cidadão"
   ClientHeight    =   7515
   ClientLeft      =   7110
   ClientTop       =   1755
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtHist 
      Height          =   6165
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   870
      Width           =   8625
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   7620
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   7140
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
      MICON           =   "frmObsCidadao.frx":0000
      PICN            =   "frmObsCidadao.frx":001C
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
      Left            =   6480
      TabIndex        =   3
      ToolTipText     =   "Gravar os Dados"
      Top             =   7140
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
      MICON           =   "frmObsCidadao.frx":008A
      PICN            =   "frmObsCidadao.frx":00A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   795
      Left            =   30
      Top             =   60
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   1402
      FrameColor      =   12829635
      Style           =   0
      Caption         =   "Histórico de Alterações"
      TextColor       =   13579779
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ComboBox cmbHist 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   7275
      End
      Begin prjChameleon.chameleonButton btLoad 
         Height          =   345
         Left            =   7530
         TabIndex        =   1
         ToolTipText     =   "Exibir histórico anterior"
         Top             =   330
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Exibir"
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
         MCOL            =   16711935
         MPTR            =   1
         MICON           =   "frmObsCidadao.frx":044B
         PICN            =   "frmObsCidadao.frx":0467
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmObsCidadao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nCodCidadao As Long, sOldHist As String

Public Property Let nContribuinte(nCodigoContribuinte As Long)
    nCodCidadao = nCodigoContribuinte
End Property

Private Sub cmbHist_Click()
Dim Sql As String, RdoAux As rdoResultset

If cmdGravar.Enabled = True Then
    If MsgBox("Deseja salvar as alterações na observação?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
        cmdGravar_Click
        Exit Sub
    End If
End If

If cmbHist.ListIndex = -1 Then
    Exit Sub
End If

On Error Resume Next
txtHist.Text = ""
Sql = "SELECT  * FROM obscidadao where id=" & cmbHist.ItemData(cmbHist.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
'    If .RowCount > 0 Then
        txtHist.Text = RdoAux!obs
        sOldHist = txtHist.Text
'    End If
   .Close
End With
cmdGravar.Enabled = False
End Sub

Private Sub cmdGravar_Click()
Dim Sql As String

'Sql = "insert obscidadao(codigo,timestamp,obs,usuario) values(" & nCodCidadao & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & Mask(txtHist.Text) & "','" & NomeDeLogin & "')"
Sql = "insert obscidadao(codigo,timestamp,obs,USERID) values(" & nCodCidadao & ",'" & Format(Now, sDataFormat & " hh:mm:ss") & "','" & Mask(txtHist.Text) & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
cn.Execute Sql, rdExecDirect
cmdGravar.Enabled = False
Carrega

End Sub

Private Sub cmdSair_Click()

If cmdGravar.Enabled = True Then
    If MsgBox("Deseja salvar as alterações na observação?", vbQuestion + vbYesNo, "ATENÇÃO") = vbYes Then
        cmdGravar_Click
    End If
End If

Unload Me
End Sub

Private Sub Form_Load()

Carrega

End Sub

Private Sub txtHist_Change()
    cmdGravar.Enabled = txtHist.Text <> sOldHist
    
End Sub

Private Sub Carrega()
Dim Sql As String, RdoAux As rdoResultset

cmbHist.Clear
Me.Caption = Me.Caption & " - Código: " & nCodCidadao
'Sql = "SELECT  id,codigo,timestamp,usuario,obs,nomecompleto FROM obscidadao INNER JOIN  usuario ON obscidadao.usuario = usuario.nomelogin where codigo=" & nCodCidadao & " order by timestamp desc"
Sql = "SELECT obscidadao.id, obscidadao.codigo, obscidadao.timestamp, obscidadao.obs, obscidadao.userid, usuario.nomecompleto FROM obscidadao INNER JOIN "
Sql = Sql & "usuario ON obscidadao.userid = usuario.Id where codigo=" & nCodCidadao & " order by timestamp desc"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbHist.AddItem Format(!TimeStamp, "dd/mm/yyyy hh:mm:ss") & " - " & !NomeCompleto
        cmbHist.ItemData(cmbHist.NewIndex) = !id
       .MoveNext
    Loop
   .Close
End With

cmdGravar.Enabled = False
If cmbHist.ListCount > 0 Then
    cmbHist.ListIndex = 0
End If

End Sub
