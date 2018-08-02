VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmChat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comunicador interno do GTI"
   ClientHeight    =   4980
   ClientLeft      =   2250
   ClientTop       =   3555
   ClientWidth     =   10650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   10650
   Begin VB.TextBox txtUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   1215
      TabIndex        =   12
      Top             =   100
      Width           =   9375
   End
   Begin VB.Frame frGrupo 
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   8415
      TabIndex        =   8
      Top             =   720
      Width           =   2190
      Begin VB.ComboBox cmbGrupo 
         Height          =   315
         Left            =   45
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   90
         Width           =   2085
      End
      Begin MSComctlLib.ListView lvMain2 
         Height          =   3255
         Left            =   0
         TabIndex        =   10
         Top             =   450
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5175
      Top             =   4005
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
            Picture         =   "frmChat.frx":0552
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdOpcoes 
      Height          =   315
      Left            =   10170
      TabIndex        =   6
      ToolTipText     =   "Outras opções"
      Top             =   4590
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
      MICON           =   "frmChat.frx":08A4
      PICN            =   "frmChat.frx":08C0
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
      Top             =   4500
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   9720
      TabIndex        =   5
      ToolTipText     =   "Limpar tudo"
      Top             =   4590
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
      MICON           =   "frmChat.frx":0971
      PICN            =   "frmChat.frx":098D
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
      Height          =   3615
      Left            =   8415
      TabIndex        =   4
      Top             =   810
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
      Top             =   4500
   End
   Begin prjChameleon.chameleonButton cmdEnviar 
      Default         =   -1  'True
      Height          =   315
      Left            =   8415
      TabIndex        =   1
      ToolTipText     =   "Enviar mensagem"
      Top             =   4590
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
      MICON           =   "frmChat.frx":0A2F
      PICN            =   "frmChat.frx":0A4B
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
      Top             =   4590
      Width           =   7305
   End
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   4125
      Left            =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   405
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   7276
      _Version        =   393217
      BackColor       =   -2147483633
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmChat.frx":0AB9
   End
   Begin vbalDTab6.vbalDTabControl vTab 
      Height          =   4155
      Left            =   8415
      TabIndex        =   7
      Top             =   405
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   7329
      AllowScroll     =   0   'False
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Enviando para:"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Mensagem..:"
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   4635
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
Dim soundfile As String
Dim returnval As Long
Private Type Usuarios
    sNomeLogin As String
    sFullName As String
    bLogado As Boolean
    sData As String
End Type
Dim bOnline As Boolean
Dim bBip As Boolean
Dim bRunOnce As Boolean
Dim bInvisivel As Boolean
Private aNome() As Usuarios, aNomeTmp() As Usuarios

Private Sub cmbGRUPO_Click()
Dim z As Long
z = SendMessage(lvMain2.hwnd, LVM_DELETEALLITEMS, 0, 0)

If cmbGrupo.ListIndex > -1 Then
    
    Sql = "SELECT USUARIO.NOMELOGIN,USUARIO.LOGON FROM USUARIO INNER JOIN "
    Sql = Sql & "CHATGRUPOUSUARIO ON USUARIO.NOMELOGIN = CHATGRUPOUSUARIO.NOME "
    Sql = Sql & "WHERE USUARIO.ATIVO=1 AND CHATGRUPOUSUARIO.GRUPO =" & cmbGrupo.ItemData(cmbGrupo.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            If IsNull(!logon) Then
                bLogado = False
            Else
                bLogado = !logon
            End If
            If bLogado Then
                Set itmX = lvMain2.ListItems.Add(, , !NomeLogin, , 1)
                itmX.ForeColor = &H8000&
                itmX.Bold = False
                itmX.Checked = True
            Else
                If Not bOnline Then
                    Set itmX = lvMain2.ListItems.Add(, , !NomeLogin, , 2)
                    itmX.ForeColor = &H808080
                    itmX.Bold = False
                    itmX.Checked = True
                End If
            End If
            
           .MoveNext
        Loop
       .Close
    End With
    SendToMsg
End If

End Sub

Private Sub cmdEnviar_Click()
Dim x As Integer, bAchou As Boolean, sDest As String, nCod As Integer
Dim RdoAux As rdoResultset, Sql As String

If Trim(txtMsg.Text) = "" Then
    MsgBox "Digite uma mensagem.", vbExclamation, "Atenção"
    Exit Sub
End If

bAchou = False: sDest = ""

If vTab.Tabs.Item(1).Selected = True Then
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
Else
    For x = 1 To lvMain2.ListItems.Count
        If lvMain2.ListItems(x).Checked Then
            bAchou = True
            sDest = sDest & lvMain2.ListItems(x).Text & ", "
        End If
    Next
    If Not bAchou Then
        MsgBox "Nenhum usuário selecionado.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If


sDest = Left(sDest, Len(sDest) - 2)
If NomeDeLogin = "SCHWARTZ" Then
    Rtb.SelFontName = "Comic Sans MS"
Else
    Rtb.SelFontName = "MS Sans Serif"
End If

With Rtb
    .SelStart = Len(.Text) + 1
    .SelColor = &H8000&
    .SelText = "[" & Format(Now, "hh:mm:ss") & "] "
    .SelColor = vbBlue
    .SelText = NomeDeLogin & " para "
    .SelColor = vbRed
    .SelText = sDest & ": "
    .SelColor = vbBlack
    .SelText = txtMsg.Text & vbCrLf
End With

If vTab.Tabs.Item(1).Selected = True Then
    For x = 1 To lvMain.ListItems.Count
        If lvMain.ListItems(x).Checked Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM CHAT WHERE DATACHAT='" & Format(Now, "mm/dd/yyyy") & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux!maximo) Then
                nCod = 1
            Else
                nCod = RdoAux!maximo + 1
            End If
            RdoAux.Close
            Sql = "INSERT CHAT(DATACHAT,SEQ,HORA,USERSEND,USERRECEIVED,MSG,ATIVO) VALUES('" & Format(Now, "mm/dd/yyyy") & "'," & nCod & ",'" & Format(Now, "hh:mm:ss") & "','" & NomeDeLogin & "','"
            Sql = Sql & lvMain.ListItems(x).Text & "','" & Left(Mask(Encrypt128(txtMsg.Text, "GTIchat")), 2000) & "',1)"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT C001(DATACHAT,SEQ,HORA,USERSEND,USERRECEIVED,MSG,ATIVO) VALUES('" & Format(Now, "mm/dd/yyyy") & "'," & nCod & ",'" & Format(Now, "hh:mm:ss") & "','" & NomeDeLogin & "','"
            Sql = Sql & lvMain.ListItems(x).Text & "','" & Left(Mask(txtMsg.Text), 2000) & "',1)"
            cn.Execute Sql, rdExecDirect
        End If
    Next
Else
    For x = 1 To lvMain2.ListItems.Count
        If lvMain2.ListItems(x).Checked Then
            Sql = "SELECT MAX(SEQ) AS MAXIMO FROM CHAT WHERE DATACHAT='" & Format(Now, "mm/dd/yyyy") & "'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            If IsNull(RdoAux!maximo) Then
                nCod = 1
            Else
                nCod = RdoAux!maximo + 1
            End If
            RdoAux.Close
            Sql = "INSERT CHAT(DATACHAT,SEQ,HORA,USERSEND,USERRECEIVED,MSG,ATIVO) VALUES('" & Format(Now, "mm/dd/yyyy") & "'," & nCod & ",'" & Format(Now, "hh:mm:ss") & "','" & NomeDeLogin & "','"
            Sql = Sql & lvMain2.ListItems(x).Text & "','" & Left(Mask(Encrypt128(txtMsg.Text, "GTIchat")), 2000) & "',1)"
            cn.Execute Sql, rdExecDirect
            Sql = "INSERT C001(DATACHAT,SEQ,HORA,USERSEND,USERRECEIVED,MSG,ATIVO) VALUES('" & Format(Now, "mm/dd/yyyy") & "'," & nCod & ",'" & Format(Now, "hh:mm:ss") & "','" & NomeDeLogin & "','"
            Sql = Sql & lvMain2.ListItems(x).Text & "','" & Left(Mask(txtMsg.Text), 2000) & "',1)"
            cn.Execute Sql, rdExecDirect
        End If
    Next
End If
txtMsg.Text = ""
txtMsg.SetFocus

End Sub

Private Sub cmdExcluir_Click()
Dim x As Integer

If MsgBox("Deseja limpar todas as mensagens?", vbQuestion + vbYesNo, "Conformação") = vbNo Then Exit Sub
If NomeDeLogin = "SCHWARTZ" Then
    Rtb.SelFontName = "Comic Sans MS"
Else
    Rtb.SelFontName = "MS Sans Serif"
End If

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

txtUsers.Text = ""

End Sub

Private Sub cmdOpcoes_Click()
lIndex = m_cMenuOpcoes.ShowPopupMenu(cmdOpcoes.Left, cmdOpcoes.Top - 2000, cmdOpcoes.Left, cmdOpcoes.Top, Me.ScaleWidth - cmdOpcoes.Left - cmdOpcoes.Width, cmdOpcoes.Top + cmdOpcoes.Height, False)
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset, bLogado As Boolean
soundfile = App.Path & "\bin\Monitor.wav"
Centraliza Me
bRunOnce = True
ReDim aNome(0)
On Error GoTo fim
Dim c As cTab
With vTab
    Set c = .Tabs.Add("Tab1", , "Usuários")
    c.Panel = lvMain
    Set c = .Tabs.Add("Tab2", , "Grupos")
    c.Panel = frGrupo
End With

If NomeDeLogin = "SCHWARTZ" Then
    Rtb.SelFontName = "Comic Sans MS"
Else
    Rtb.SelFontName = "MS Sans Serif"
End If

Sql = "select * from usuario where nomelogin='" & NomeDeLogin & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

RdoAux.Close

MontaMenu
bOnline = True
bBip = False
LoadList

CarregaUser
With Rtb
    .SelBold = True
    .SelColor = &HC000C0
    .SelText = "Chat iniciado as: " & Format(Now, "hh:mm:ss") & " - logado como " & NomeDeLogin & vbCrLf & vbCrLf
    .SelBold = False
End With

Sql = "SELECT CODIGO,NOME FROM CHATGRUPO ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbGrupo.AddItem !Nome
        cmbGrupo.ItemData(cmbGrupo.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With
fim:
End Sub

Private Sub CarregaUser()
Dim Sql As String, RdoAux As rdoResultset, bAchou As Boolean
Dim itmX As ListItem, x As Integer, Y As Integer, z As Long
Dim oListItem As ListItem
On Error GoTo Erro
Inicio:
ReDim aNomeTmp(0)
Sql = "SELECT * FROM USUARIO WHERE NOMELOGIN<>'" & NomeDeLogin & "' AND ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        
        ReDim Preserve aNomeTmp(UBound(aNomeTmp) + 1)
        aNomeTmp(UBound(aNomeTmp)).sNomeLogin = !NomeLogin
        aNomeTmp(UBound(aNomeTmp)).sFullName = !NomeCompleto
        If IsNull(!logon) Then
            aNomeTmp(UBound(aNomeTmp)).bLogado = False
        Else
            aNomeTmp(UBound(aNomeTmp)).bLogado = !logon
        End If
        If IsDate(!DATALOGON) Then
            aNomeTmp(UBound(aNomeTmp)).sData = Format(!DATALOGON, "dd/mm/yyyy")
        Else
            aNomeTmp(UBound(aNomeTmp)).sData = "01/01/1990"
        End If
        If DateDiff("d", CDate(aNomeTmp(UBound(aNomeTmp)).sData), Now) > 0 And aNomeTmp(UBound(aNomeTmp)).bLogado Then
            If NomeDeLogin <> "SCHWARTZ" Then
                Sql = "UPDATE USUARIO SET LOGON=0,DATALOGON=null WHERE NOMELOGIN='" & !NomeLogin & "'"
                cn.Execute Sql, rdExecDirect
            End If
        End If
proximo:
       .MoveNext
    Loop
   .Close
End With

If UBound(aNome) = 0 Then
    For x = 1 To UBound(aNomeTmp)
        ReDim Preserve aNome(UBound(aNome) + 1)
        aNome(x).bLogado = aNomeTmp(x).bLogado
        aNome(x).sFullName = aNomeTmp(x).sFullName
        aNome(x).sNomeLogin = aNomeTmp(x).sNomeLogin
    Next
End If

On Error GoTo fim
For x = 1 To UBound(aNomeTmp)
    If aNomeTmp(x).bLogado <> aNome(x).bLogado Then
        Set oListItem = ListViewFindItem(aNomeTmp(x).sNomeLogin, lvMain, elvSearchText)
        If Not oListItem Is Nothing Then
            If Not aNomeTmp(x).bLogado Then
                If Not bOnline Then
                    lvMain.ListItems(oListItem.Index).ForeColor = &H808080
                    lvMain.ListItems(oListItem.Index).Bold = False
                    lvMain.ListItems(oListItem.Index).SmallIcon = 2
                Else
                    lvMain.ListItems.Remove (oListItem.Index)
                End If
            Else
                lvMain.ListItems(oListItem.Index).ForeColor = &H8000&
                lvMain.ListItems(oListItem.Index).Bold = False
                lvMain.ListItems(oListItem.Index).SmallIcon = 1
            End If
        Else
            If aNomeTmp(x).bLogado = True Then
                Set itmX = lvMain.ListItems.Add(, , aNomeTmp(x).sNomeLogin, , 1)
                itmX.ForeColor = &H8000&
                itmX.Bold = False
            Else
                
            End If
        End If
    End If
Next

For x = 1 To UBound(aNomeTmp)
    If aNomeTmp(x).bLogado <> aNome(x).bLogado Then
        Set oListItem = ListViewFindItem(aNomeTmp(x).sNomeLogin, lvMain2, elvSearchText)
        If Not oListItem Is Nothing Then
            If Not aNomeTmp(x).bLogado Then
                lvMain2.ListItems(oListItem.Index).ForeColor = &H808080
                lvMain2.ListItems(oListItem.Index).Bold = False
                lvMain2.ListItems(oListItem.Index).SmallIcon = 2
            Else
                lvMain2.ListItems(oListItem.Index).ForeColor = &H8000&
                lvMain2.ListItems(oListItem.Index).Bold = False
                lvMain2.ListItems(oListItem.Index).SmallIcon = 1
            End If
        End If
    End If
Next

ReDim aNome(0)
For x = 1 To UBound(aNomeTmp)
    ReDim Preserve aNome(UBound(aNome) + 1)
    aNome(x).bLogado = aNomeTmp(x).bLogado
    aNome(x).sFullName = aNomeTmp(x).sFullName
    aNome(x).sNomeLogin = aNomeTmp(x).sNomeLogin
Next
Y = 0
For x = 1 To UBound(aNomeTmp)
    If aNomeTmp(x).bLogado Then
        Y = Y + 1
    End If
Next

Me.Caption = "Comunicador interno do GTI - " & Y + 1 & " usuários conectados."
frmMdi.Sbar.Panels(1).Text = Y + 1 & " usuários conectados."
If InStr(1, cn.Connect, "SERVER=192.") > 0 Then
    frmMdi.Sbar.Panels(1).Text = Y + 1 & " usuários conectados."
Else
    frmMdi.Sbar.Panels(1).Text = Y + 1 & " usuários conectados (Base Local)."
End If
        
fim:
Exit Sub
Erro:
MsgBox Err.Description
Conecta NomeDeLogin, UserPwd, ""
GoTo Inicio
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

Private Sub lvMain_ItemCheck(ByVal Item As MSComctlLib.ListItem)
SendToMsg
End Sub

Private Sub lvMain2_ItemCheck(ByVal Item As MSComctlLib.ListItem)
SendToMsg
End Sub

Private Sub m_cMenuOpcoes_Click(ItemNumber As Long)
Dim Sql As String, RdoAux As rdoResultset, sNomeArq As String, ax As String

Select Case m_cMenuOpcoes.ItemKey(ItemNumber)
    Case "mnuInvisivel"
        m_cMenuOpcoes.Checked(ItemNumber) = Not m_cMenuOpcoes.Checked(ItemNumber)
        If m_cMenuOpcoes.Checked(ItemNumber) = True Then
            bInvisivel = True
            Sql = "update usuario set logon=0 where nomelogin='" & NomeDeLogin & "'"
        Else
            bInvisivel = False
            Sql = "update usuario set logon=1 where nomelogin='" & NomeDeLogin & "'"
        End If
        cn.Execute Sql, rdExecDirect
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
                ax = "[" & Format(!DATACHAT, "dd/mm") & " às " & Format(!HORA, "hh:mm") & "] " & !USERSEND & " para " & !USERRECEIVED & " : " & Decrypt128(!Msg, "GTIchat")
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
    Case "mnuOnline"
        m_cMenuOpcoes.Checked(ItemNumber) = Not m_cMenuOpcoes.Checked(ItemNumber)
        bOnline = m_cMenuOpcoes.Checked(ItemNumber)
        LoadList
        cmbGRUPO_Click
    Case "mnuBip"
        m_cMenuOpcoes.Checked(ItemNumber) = Not m_cMenuOpcoes.Checked(ItemNumber)
        bBip = m_cMenuOpcoes.Checked(ItemNumber)
End Select

End Sub

Private Sub Timer1_Timer()
Dim Sql As String, RdoAux As rdoResultset
Dim retval As Integer, FWInfo As FLASHWINFO

With FWInfo
    .cbSize = 20
    .hwnd = frmMdi.hwnd
    .dwFlags = FLASHW_ALL
    .uCount = 5
    .dwTimeout = 0
End With

CarregaUser
''
If UCase(NomeDeLogin) = "SCHWARTZ" Then
    Exit Sub
End If

If bRunOnce Then
    Sql = "SELECT * FROM chat WHERE DATACHAT='" & Format(Now, sDataFormat) & "' AND USERRECEIVED='" & NomeDeLogin & "' AND ATIVO=1"
Else
    Sql = "SELECT * FROM chat WHERE USERRECEIVED='" & NomeDeLogin & "' AND ATIVO=1"
    bRunOnce = False
End If

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 And (Me.Visible = False Or Me.WindowState = vbMinimized) Then
        RdoAux.Close
        Timer2.Interval = 1000
        If frmMdi.WindowState = vbMinimized Then
            retval = FlashWindowEx(FWInfo)
            If bBip Then
                returnval = PlaySound(soundfile, 0, &H0)
            End If
        End If
    ElseIf .RowCount > 0 And (frmMdi.WindowState = vbMinimized) Then
        retval = FlashWindowEx(FWInfo)
        If bBip Then
            returnval = PlaySound(soundfile, 0, &H0)
        End If
    Else
        Timer2.Interval = 0
        frmMdi.btBar(11).BackColor = &HE7E3E7
        Do Until .EOF
            With Rtb
                If NomeDeLogin = "SCHWARTZ" Then
                    Rtb.SelFontName = "Comic Sans MS"
                Else
                    Rtb.SelFontName = "MS Sans Serif"
                End If
                .SelColor = &H8000&
                .SelText = "[" & Format(RdoAux!HORA, "hh:mm:ss") & "] "
                .SelColor = vbBlue
                .SelText = RdoAux!USERSEND & " para "
                .SelColor = vbRed
                .SelText = NomeDeLogin & ": "
                .SelColor = vbBlack
                .SelText = Decrypt128(RdoAux!Msg, "GTIchat") & vbCrLf
            End With
 '           If NomeDeLogin <> "SCHWARTZ" Then
                Sql = "UPDATE CHAT SET ATIVO=0  WHERE DATACHAT='" & Format(Now, "mm/dd/yyyy") & "' AND SEQ=" & RdoAux!Seq
                cn.Execute Sql, rdExecDirect
 '           End If
           .MoveNext
        Loop
        
    End If
    On Error Resume Next
   RdoAux.Close
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
      
      i = .AddItem("Ficar invisível", "", 1, , , bInvisivel, , "mnuInvisivel")
      .OwnerDraw(i) = True
      i = .AddItem("Exibir histórico do chat", "", 1, , , , , "mnuHistorico")
      .OwnerDraw(i) = True
      i = .AddItem("Excluir todo o histórico", "", 1, , , , , "mnuExcluir")
      .OwnerDraw(i) = True
      i = .AddItem("Exibir nome completo do usuário", "", 1, , , , , "mnuExibir")
      .OwnerDraw(i) = True
      i = .AddItem("Exibir apenas usuários Online", "", 1, , , True, , "mnuOnline")
      .OwnerDraw(i) = True
      i = .AddItem("Ativar aviso sonóro", "", 1, , , False, , "mnuBip")
      .OwnerDraw(i) = True
      i = .AddItem("Gerenciador de grupos", "", 1, , , , , "mnuGrupo")
      .OwnerDraw(i) = True
   End With

End Sub

Private Sub SendToMsg()
Dim x As Integer, sDest As String
sDest = ""
txtUsers.Text = ""
If vTab.Tabs.Item(1).Selected = True Then
    For x = 1 To lvMain.ListItems.Count
        If lvMain.ListItems(x).Checked Then
            sDest = sDest & lvMain.ListItems(x).Text & ","
        End If
    Next
Else
    For x = 1 To lvMain2.ListItems.Count
        If lvMain2.ListItems(x).Checked Then
            sDest = sDest & lvMain2.ListItems(x).Text & ","
        End If
    Next
End If
If sDest <> "" Then
    txtUsers.Text = Left(sDest, Len(sDest) - 1)
End If

End Sub

Private Sub vTab_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal x As Single, ByVal Y As Single)
SendToMsg
End Sub

Private Sub LoadList()
Dim Sql As String, RdoAux As rdoResultset, bLogado As Boolean, z As Long
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)
Sql = "SELECT * FROM USUARIO WHERE NOMELOGIN<>'" & NomeDeLogin & "' AND ATIVO=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If IsNull(!logon) Then
            bLogado = False
        Else
            bLogado = !logon
        End If
        If bLogado Then
            Set itmX = lvMain.ListItems.Add(, , !NomeLogin, , 1)
            itmX.ForeColor = &H8000&
            itmX.Bold = False
        Else
            If Not bOnline Then
                Set itmX = lvMain.ListItems.Add(, , !NomeLogin, , 2)
                itmX.ForeColor = &H808080
                itmX.Bold = False
            End If
        End If
        
       .MoveNext
    Loop
   .Close
End With

End Sub
