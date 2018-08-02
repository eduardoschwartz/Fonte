VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunicador interno do GTI"
   ClientHeight    =   4665
   ClientLeft      =   2250
   ClientTop       =   3555
   ClientWidth     =   10650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   10650
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5175
      Top             =   3690
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
            Picture         =   "frmChat.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":046A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdOpcoes 
      Height          =   315
      Left            =   10170
      TabIndex        =   6
      ToolTipText     =   "Outras opções"
      Top             =   4275
      Width           =   405
      _ExtentX        =   714
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
      MCOL            =   14737632
      MPTR            =   1
      MICON           =   "frmChat.frx":07BC
      PICN            =   "frmChat.frx":07D8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer Timer2 
      Left            =   765
      Top             =   4185
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   9585
      TabIndex        =   5
      ToolTipText     =   "Limpar tudo"
      Top             =   4275
      Width           =   405
      _ExtentX        =   714
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmChat.frx":0889
      PICN            =   "frmChat.frx":08A5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4125
      Left            =   8415
      TabIndex        =   4
      Top             =   90
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   7276
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Usuários conectados"
         Object.Width           =   3352
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   315
      Top             =   4185
   End
   Begin prjChameleon.chameleonButton cmdEnviar 
      Default         =   -1  'True
      Height          =   315
      Left            =   8415
      TabIndex        =   1
      ToolTipText     =   "Enviar mensagem"
      Top             =   4275
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Enviar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChat.frx":0947
      PICN            =   "frmChat.frx":0963
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1080
      MaxLength       =   2000
      TabIndex        =   0
      Top             =   4320
      Width           =   7305
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   4125
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   7276
      _Version        =   393217
      BackColor       =   -2147483633
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":09D1
   End
   Begin VB.Label Label1 
      Caption         =   "Mensagem..:"
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   4320
      Width           =   960
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuOpcoes As cPopupMenu
Attribute m_cMenuOpcoes.VB_VarHelpID = -1
Private Type Usuarios
    sNomeLogin As String
    sFullName As String
    bLogado As Boolean
End Type
Private aNome() As Usuarios

Private Sub cmdEnviar_Click()
Dim x As Integer, bAchou As Boolean, sDest As String, nCod As Integer
Dim RdoAux As rdoResultset, Sql As String

If Trim(txtMsg.Text) = "" Then
    MsgBox "Digite uma mensagem.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False: sDest = ""
For x = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(x).Checked Then
        bAchou = True
        sDest = sDest & lvMain.ListItems(x).Text & ", "
    End If
Next
If Not bAchou Then
    MsgBox "Nenhum usuário selecionado.", vbExclamation, "Atenção"
    Exit Sub
End If
sDest = Left(sDest, Len(sDest) - 2)

With Rtb
    .SelColor = &H8000&
    .SelText = "[" & Format(Now, "hh:mm:ss") & "] "
    .SelColor = vbBlue
    .SelText = NomeDeLogin & " para "
    .SelColor = vbRed
    .SelText = sDest & ": "
    .SelColor = vbBlack
    .SelText = txtMsg.Text & vbCrLf
End With

For x = 1 To lvMain.ListItems.Count
    If lvMain.ListItems(x).Checked Then
        Sql = "SELECT MAX(SEQ) AS MAXIMO FROM CHAT WHERE DATACHAT='" & Format(Now, "mm/dd/yyyy") & "'"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            nCod = 1
        Else
            nCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        Sql = "INSERT CHAT(DATACHAT,SEQ,HORA,USERSEND,USERRECEIVED,MSG,ATIVO) VALUES('" & Format(Now, "mm/dd/yyyy") & "'," & nCod & ",'" & Format(Now, "hh:mm:ss") & "','" & NomeDeLogin & "','"
        Sql = Sql & lvMain.ListItems(x).Text & "','" & Mask(txtMsg.Text) & "',1)"
        cn.Execute Sql, rdExecDirect
    End If
Next

txtMsg.Text = ""
txtMsg.SetFocus

End Sub

Private Sub cmdExcluir_Click()
Dim x As Integer

If MsgBox("Deseja limpar todas as mensagens?", vbQuestion + vbYesNo, "Conformação") = vbNo Then Exit Sub

With Rtb
    .Text = ""
    .SelBold = True
    .SelColor = &HC000C0
    .SelText = "Chat iniciado as: " & Format(Now, "hh:mm:ss") & " - logado como " & NomeDeLogin & vbCrLf & vbCrLf
    .SelBold = False
End With

For x = 1 To lvMain.ListItems.Count
    lvMain.ListItems(x).Checked = False
Next

End Sub

Private Sub cmdOpcoes_Click()
lIndex = m_cMenuOpcoes.ShowPopupMenu(cmdOpcoes.Left, cmdOpcoes.Top - 1450, cmdOpcoes.Left, cmdOpcoes.Top, Me.ScaleWidth - cmdOpcoes.Left - cmdOpcoes.Width, cmdOpcoes.Top + cmdOpcoes.Height, False)
End Sub

Private Sub Form_Load()

MontaMenu
CarregaUser
With Rtb
    .SelBold = True
    .SelColor = &HC000C0
    .SelText = "Chat iniciado as: " & Format(Now, "hh:mm:ss") & " - logado como " & NomeDeLogin & vbCrLf & vbCrLf
    .SelBold = False
End With

End Sub

Private Sub CarregaUser()
Dim Sql As String, RdoAux As rdoResultset, bAchou As Boolean
Dim itmX As ListItem, x As Integer, Y As Integer
Dim oListItem As ListItem
On Error GoTo Erro

ReDim aNome(0)
Sql = "SELECT * FROM USUARIO WHERE NOMELOGIN<>'" & NomeDeLogin & "' AND ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aNome(UBound(aNome) + 1)
        aNome(UBound(aNome)).sNomeLogin = !NomeLogin
        aNome(UBound(aNome)).sFullName = !NomeCompleto
        If IsNull(!logon) Then
            aNome(UBound(aNome)).bLogado = False
        Else
            aNome(UBound(aNome)).bLogado = !logon
        End If
            
       .MoveNext
    Loop
   .Close
End With

If lvMain.ListItems.Count = 0 Then GoTo Inclui
'Inicio:
For x = 1 To UBound(aNome)
    'tiramos quem não esta mais logado
    Set oListItem = ListViewFindItem(aNome(x).sNomeLogin, lvMain, elvSearchText)
    If Not aNome(x).bLogado And lvMain.ListItems(oListItem.Index).Bold = True Then
        Sql = "UPDATE USUARIO SET LOGON=0 WHERE NOMELOGIN='" & lvMain.ListItems(oListItem.Index).Text & "'"
        cn.Execute Sql, rdExecDirect
        lvMain.ListItems(oListItem.Index).ForeColor = &H808080
        lvMain.ListItems(oListItem.Index).Bold = False
        lvMain.ListItems(oListItem.Index).SmallIcon = 2
    
    End If
Next
Inclui:
'incluimos quem não esta
For x = 1 To UBound(aNome)
    bAchou = False
    For Y = 1 To lvMain.ListItems.Count
        If aNome(x).sNomeLogin = lvMain.ListItems(Y).Text Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        If aNome(x).bLogado Then
            Set itmX = lvMain.ListItems.Add(, , aNome(x).sNomeLogin, , 1)
            itmX.ForeColor = &H8000&
            itmX.Bold = True
        Else
            Set itmX = lvMain.ListItems.Add(, , aNome(x).sNomeLogin, , 2)
            itmX.ForeColor = &H808080
            itmX.Bold = False
        End If
    Else
        If aNome(x).bLogado Then
            lvMain.ListItems(Y).ForeColor = &H8000&
            lvMain.ListItems(Y).Bold = True
            lvMain.ListItems(Y).SmallIcon = 1
        Else
            lvMain.ListItems(Y).ForeColor = &H808080
            lvMain.ListItems(Y).Bold = False
            lvMain.ListItems(Y).SmallIcon = 2
        End If
    End If
Next
Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = Not bCloseChat
If Not bCloseChat Then
    Me.Hide
End If
End Sub

Private Sub Form_Resize()
frmMdi.btBar(11).BackColor = &HE7E3E7
Timer2.Interval = 0

End Sub

Private Sub m_cMenuOpcoes_Click(ItemNumber As Long)
Dim Sql As String, RdoAux As rdoResultset, sNomeArq As String, ax As String

Select Case m_cMenuOpcoes.ItemKey(ItemNumber)
    Case "mnuExibir"
        On Error Resume Next
        MsgBox RetornaUsuarioFullName2(lvMain.SelectedItem.Text), vbOKOnly, "Nome completo do usuário: " & lvMain.SelectedItem.Text
        On Error GoTo 0
    Case "mnuExcluir"
        If MsgBox("Deseja excluir todo o seu histórico de conversa?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "DELETE FROM CHAT WHERE USERSEND='" & NomeDeLogin & "' OR USERRECEIVED='" & NomeDeLogin & "' AND ATIVO=0"
            cn.Execute Sql, rdExecDirect
            With Rtb
                .Text = ""
                .SelBold = True
                .SelColor = &HC000C0
                .SelText = "Chat iniciado as: " & Format(Now, "hh:mm:ss") & " - logado como " & NomeDeLogin & vbCrLf & vbCrLf
                .SelBold = False
            End With
        End If
    Case "mnuHistorico"
        sNomeArq = App.Path & "\bin\" & NomeDeLogin & "Chat.txt"
        FF1 = FreeFile()
        Open sNomeArq For Output As FF1
        Print #FF1, "*************************************************"
        Print #FF1, "Histórico gerado em: " & Now
        Print #FF1, "***************************************************"
        Print #FF1, ""
        Sql = "SELECT * FROM CHAT WHERE USERSEND='" & NomeDeLogin & "' OR USERRECEIVED='" & NomeDeLogin & "' ORDER BY DATACHAT,HORA"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                ax = "[" & Format(!DATACHAT, "dd/mm") & " às " & Format(!HORA, "hh:mm") & "] " & !USERSEND & " para " & !USERRECEIVED & " : " & !Msg
                Print #FF1, ax
               .MoveNext
            Loop
           .Close
        End With
        Close #FF1
        ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)
    Case "mnuGrupo"
        If NomeDeLogin <> "SCHWARTZ" Then
            MsgBox "Em manutenção", vbExclamation, "Erro"
            Exit Sub
        End If
        frmGrupo.show
        frmGrupo.ZOrder (0)
End Select

End Sub

Private Sub Timer1_Timer()
Dim Sql As String, RdoAux As rdoResultset

CarregaUser
Exit Sub
Sql = "SELECT * FROM chat WHERE DATACHAT='" & Format(Now, "mm/dd/yyyy") & "' AND USERRECEIVED='" & NomeDeLogin & "' AND ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 And (Me.Visible = False Or Me.WindowState = vbMinimized) Then
        RdoAux.Close
        Timer2.Interval = 1000
        Exit Sub
    Else
        Timer2.Interval = 0
        frmMdi.btBar(11).BackColor = &HE7E3E7
    End If
    Do Until .EOF
        With Rtb
            .SelColor = &H8000&
            .SelText = "[" & Format(RdoAux!HORA, "hh:mm:ss") & "] "
            .SelColor = vbBlue
            .SelText = RdoAux!USERSEND & " para "
            .SelColor = vbRed
            .SelText = NomeDeLogin & ": "
            .SelColor = vbBlack
            .SelText = RdoAux!Msg & vbCrLf
        End With
        Sql = "UPDATE CHAT SET ATIVO=0  WHERE DATACHAT='" & Format(Now, "mm/dd/yyyy") & "' AND SEQ=" & RdoAux!Seq
        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Timer2_Timer()
If frmMdi.btBar(11).BackColor = &HE7E3E7 Then
    frmMdi.btBar(11).BackColor = vbRed
Else
    frmMdi.btBar(11).BackColor = &HE7E3E7
End If
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdEnviar_Click
End If
End Sub

Private Sub MontaMenu()
   Set m_cMenuOpcoes = New cPopupMenu
   With m_cMenuOpcoes
      .hwndOwner = Me.hwnd
      .GradientHighlight = True
      
      i = .AddItem("Exibir histórico do chat", "", 1, , , , , "mnuHistorico")
      .OwnerDraw(i) = True
      i = .AddItem("Excluir todo o histórico", "", 1, , , , , "mnuExcluir")
      .OwnerDraw(i) = True
      i = .AddItem("Exibir nome completo do usuário", "", 1, , , , , "mnuExibir")
      .OwnerDraw(i) = True
      i = .AddItem("Gerenciador de grupos", "", 1, , , , , "mnuGrupo")
      .OwnerDraw(i) = True
   End With

End Sub
