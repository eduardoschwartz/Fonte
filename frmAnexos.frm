VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAnexos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controle de anexos"
   ClientHeight    =   5955
   ClientLeft      =   7245
   ClientTop       =   4635
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5955
   ScaleWidth      =   9495
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   18
      Text            =   "frmAnexos.frx":0000
      Top             =   5400
      Width           =   9315
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4200
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Inclusão manual de anexos"
      ForeColor       =   &H00000080&
      Height          =   1875
      Left            =   90
      TabIndex        =   8
      Top             =   3480
      Width           =   9315
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1710
         MaxLength       =   100
         TabIndex        =   12
         Top             =   1020
         Width           =   7425
      End
      Begin VB.TextBox txtArquivo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   690
         Width           =   6855
      End
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   5295
      End
      Begin prjChameleon.chameleonButton btCancel 
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         ToolTipText     =   "Cancelar Edição"
         Top             =   1470
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
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
         MICON           =   "frmAnexos.frx":00B0
         PICN            =   "frmAnexos.frx":00CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btGravar 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Gravar os Dados"
         Top             =   1470
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
         MICON           =   "frmAnexos.frx":0226
         PICN            =   "frmAnexos.frx":0242
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton btArquivo 
         Height          =   315
         Left            =   1710
         TabIndex        =   11
         ToolTipText     =   "Clque para selecionar um arquivo"
         Top             =   660
         Width           =   375
         _ExtentX        =   661
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
         MCOL            =   16711935
         MPTR            =   1
         MICON           =   "frmAnexos.frx":05E7
         PICN            =   "frmAnexos.frx":0603
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Observação.............:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1050
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Selecione o arquivo.:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Anexo.........:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2265
      Left            =   90
      TabIndex        =   4
      Top             =   600
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   3995
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tipo"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Data"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nome do Arquivo"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Anexado por"
         Object.Width           =   2716
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Observação"
         Object.Width           =   5293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Mes"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Ano"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "newname"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "ext"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "tipo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "seq"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "userid"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "codigo"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtNome 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   7035
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin prjChameleon.chameleonButton btBusca 
      Height          =   315
      Left            =   1860
      TabIndex        =   2
      ToolTipText     =   "Busca documento dentro dos arquivos"
      Top             =   120
      Width           =   435
      _ExtentX        =   767
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
      MICON           =   "frmAnexos.frx":06BE
      PICN            =   "frmAnexos.frx":06DA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btAdd 
      Height          =   345
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Incluir Anexo"
      Top             =   2970
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Incluir"
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
      MICON           =   "frmAnexos.frx":0834
      PICN            =   "frmAnexos.frx":0850
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btAbrir 
      Height          =   345
      Left            =   7590
      TabIndex        =   6
      ToolTipText     =   "Abrir o anexo selecionado"
      Top             =   2970
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   " Abrir o arquivo"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16711935
      MPTR            =   1
      MICON           =   "frmAnexos.frx":09AA
      PICN            =   "frmAnexos.frx":09C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton btDel 
      Height          =   345
      Left            =   1200
      TabIndex        =   7
      ToolTipText     =   "Excluir anexo"
      Top             =   2970
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   16776960
      MPTR            =   1
      MICON           =   "frmAnexos.frx":0A81
      PICN            =   "frmAnexos.frx":0A9D
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
      Caption         =   "Código..:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   765
   End
End
Attribute VB_Name = "frmAnexos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btAbrir_Click()
Dim nCodigo As Long, nTipo As Integer, nSeq As Integer, nMes As Integer, nAno As Integer, sOldName As String, sNewName As String, sExt As String
Dim sPathOrigem As String, sPathSandBox As String, fso As New FileSystemObject

If lvMain.ListItems.Count = 0 Then
    MsgBox "Selecione um anexo.", vbCritical, "Atenção"
    Exit Sub
End If

Ocupado

nCodigo = Val(txtCod.Text)
With lvMain.SelectedItem
    sOldName = .SubItems(2)
    nMes = .SubItems(5)
    nAno = .SubItems(6)
    sNewName = .SubItems(7)
    sExt = .SubItems(8)
    nTipo = .SubItems(9)
    nSeq = .SubItems(10)
End With

sPathOrigem = sPathAnexo & Format(nTipo, "00") & "\" & Format(nAno, "0000") & "\" & Format(nMes, "00") & "\" & sNewName
sPathSandBox = App.Path & "\" & "Sandbox"

If fso.FolderExists(sPathSandBox) = False Then
    fso.CreateFolder (sPathSandBox)
End If

If fso.FileExists(sPathSandBox) Then
    fso.DeleteFile (sPathSandBox)
End If
If Not fso.FolderExists(sPathAnexo & Format(nTipo, "00") & "\" & Format(nAno, "0000") & "\" & Format(nMes, "00")) Then
    fso.CreateFolder (sPathAnexo & Format(nTipo, "00") & "\" & Format(nAno, "0000") & "\" & Format(nMes, "00"))
    MsgBox "Arquivo não localizado.", vbCritical, "Erro"
Else
    sPathSandBox = sPathSandBox & "\" & sOldName
    If fso.FileExists(sPathOrigem) Then
        fso.CopyFile sPathOrigem, sPathSandBox, True
        ShellExecute 0&, "open", sPathSandBox, vbNullString, vbNullString, conSwNormal
    Else
        MsgBox "Arquivo não localizado.", vbCritical, "Erro"
    End If
End If
Liberado

End Sub

Private Sub btAdd_Click()
'If NomeDeLogin <> "SCHWARTZ" Then
 '   MsgBox "em desenvolvimento"
  '  Exit Sub
'End If

If txtNome.Text = "" Then
    MsgBox "Digite um código válido e aperte o botão consultar.", vbCritical, "Erro"
    Exit Sub
End If

If txtCod.Enabled = True Then
    txtCod.Enabled = False
    btBusca.Enabled = False
    lvMain.Enabled = False
    btDel.Enabled = False
    btAdd.Enabled = False
    btAbrir.Enabled = False
    cmbTipo.Enabled = True
    cmbTipo.ListIndex = 0
    txtObs.Enabled = True
    btGravar.Enabled = True
    btCancel.Enabled = True
    btArquivo.Enabled = True
    cmbTipo.BackColor = Branco
    txtObs.BackColor = Branco
End If

cmbTipo.SetFocus
End Sub

Private Sub btArquivo_Click()
With CommonDialog1
    .DialogTitle = "Selecione um anexo"
    .CancelError = True
    .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST
    .InitDir = App.Path & "\bin"
    .Filter = "All Files (*.*)|*.*"
    .ShowOpen
    
    txtArquivo.Text = .FileName
    
End With

End Sub

Private Sub btBusca_Click()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, sNome As String, RdoAux2 As rdoResultset, itmX As ListItem

nCodReduz = Val(txtCod.Text)
If nCodReduz = 0 Then
    MsgBox "Digite um código válido.", vbCritical, "Erro"
    Exit Sub
End If
txtNome.Text = ""

Ocupado

sNome = RetornaNome(nCodReduz)
If sNome <> "" Then
    txtNome.Text = sNome
    CarregaLista
Else
    Liberado
    MsgBox "Código não cadastrado.", vbCritical, "Erro"
    Exit Sub
End If
Liberado

If lvMain.ListItems.Count = 0 Then
    MsgBox "Não existem anexos para este código.", vbCritical, "Atenção"
End If

End Sub

Private Sub CarregaLista()
Dim Sql As String, RdoAux As rdoResultset, nCodReduz As Long, sNome As String, RdoAux2 As rdoResultset, itmX As ListItem

lvMain.ListItems.Clear
nCodReduz = Val(txtCod.Text)
ConectaBinary
Sql = "select Anexos.codigo, Anexos.Tipo,tipodocumento.nome, anexos.seq,anexos.Ano,anexos.Mes,anexos.OldName,anexos.NewName,anexos.Ext,Anexos_controle.Data,Anexos_controle.UserId,Anexos_controle.Observacao "
Sql = Sql & "from anexos inner join Anexos_controle on Anexos.Codigo=Anexos_controle.Codigo and Anexos.Tipo=Anexos_controle.Tipo and Anexos.Seq=Anexos_controle.Seq "
Sql = Sql & "inner join tipodocumento on Anexos.Tipo=tipodocumento.codigo where Anexos_controle.Codigo=" & nCodReduz
Set RdoAux2 = cnBinary.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
        Set itmX = lvMain.ListItems.Add(, , !Nome)
        itmX.SubItems(1) = Format(!Data, "dd/mm/yyyy")
        itmX.SubItems(2) = !oldname
        itmX.SubItems(3) = RetornaUsuarioFullName3(!userid)
        itmX.SubItems(4) = SubNull(!OBSERVACAO)
        itmX.SubItems(5) = !Mes
        itmX.SubItems(6) = !Ano
        itmX.SubItems(7) = !newname
        itmX.SubItems(8) = !ext
        itmX.SubItems(9) = !Tipo
        itmX.SubItems(10) = !Seq
        itmX.SubItems(11) = !userid
        itmX.SubItems(12) = !Codigo
        
       .MoveNext
    Loop
   .Close
End With
If lvMain.ListItems.Count > 0 Then
    lvMain.ListItems(1).Selected = True
End If
cnBinary.Close
End Sub


Private Sub btCancel_Click()
TravaInclusao
End Sub

Private Sub btDel_Click()
Dim nId As Integer, nId2 As Integer, nCodigo As Long, nTipo As Integer, nSeq As Integer, Sql As String

nId = RetornaUsuarioID(NomeDeLogin)
If lvMain.ListItems.Count = 0 Then
    MsgBox "Selecione um anexo.", vbCritical, "Atenção"
    Exit Sub
Else
    nId2 = Val(lvMain.SelectedItem.SubItems(11))
    If nId <> nId2 Then
        MsgBox "Você pode excluir apenas seus anexos.", vbCritical, "Atenção"
        Exit Sub
    Else
        If MsgBox("Excluir este anexo?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
            ConectaBinary
            nCodigo = Val(lvMain.SelectedItem.SubItems(12))
            nTipo = Val(lvMain.SelectedItem.SubItems(9))
            nSeq = Val(lvMain.SelectedItem.SubItems(10))
            
            Sql = "delete from anexos where codigo=" & nCodigo & " and tipo=" & nTipo & " and seq=" & nSeq
            cnBinary.Execute Sql, rdExecDirect
            Sql = "delete from anexos_controle where codigo=" & nCodigo & " and tipo=" & nTipo & " and seq=" & nSeq
            cnBinary.Execute Sql, rdExecDirect
            cnBinary.Close
            CarregaLista
        End If
    End If
End If

End Sub

Private Sub btGravar_Click()
Dim nTipo As Integer, sArquivoOld As String, sArquivoNew As String, fso As New FileSystemObject, sArq As String, nCodReduz As Long
Dim Sql As String, RdoAux As rdoResultset, nSeq As Integer

nCodReduz = Val(txtCod.Text)
If txtArquivo.Text = "" Then
    MsgBox "Selecione um arquivo válido>", vbCritical, "Erro"
    Exit Sub
End If

nTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
sArquivoOld = ParsePath(txtArquivo.Text, vbNormal)


sPath = sPathAnexo & nTipo
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If
sPath = sPathAnexo & nTipo & "\" & Year(Now)
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If
sPath = sPathAnexo & nTipo & "\" & Year(Now) & "\" & Month(Now)
If fso.FolderExists(sPath) = False Then
    fso.CreateFolder (sPath)
End If


ConectaBinary

Sql = "select max(seq) as maximo from anexos_controle where codigo=" & nCodReduz & " and tipo=" & nTipo
Set RdoAux = cnBinary.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    nSeq = 0
Else
    nSeq = RdoAux!maximo + 1
End If
RdoAux.Close

sArquivoNew = Format(nCodReduz, "000000") & Format(nTipo, "00") & Format(nSeq, "0000")
fso.CopyFile txtArquivo.Text, sPath & "\" & sArquivoNew, True

Sql = "insert anexos(codigo,tipo,seq,ano,mes,oldname,newname,ext) values(" & nCodReduz & "," & nTipo & ","
Sql = Sql & nSeq & "," & Year(Now) & "," & Month(Now) & ",'" & Mask(sArquivoOld) & "','" & sArquivoNew & "','" & Right(sArquivoOld, 3) & "')"
cnBinary.Execute Sql, rdExecDirect
 
Sql = "insert anexos_controle(codigo,tipo,seq,data,userid,observacao) values(" & nCodReduz & "," & nTipo & ","
Sql = Sql & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "'," & RetornaUsuarioID(NomeDeLogin) & ",'" & Mask(txtObs.Text) & "')"
cnBinary.Execute Sql, rdExecDirect

cnBinary.Close

CarregaLista
TravaInclusao
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset

ConectaBinary

Sql = "select codigo,nome from tipodocumento order by nome"
Set RdoAux = cnBinary.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem (!Nome)
        cmbTipo.ItemData(cmbTipo.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

cnBinary.Close
TravaInclusao
Centraliza Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim fso As New FileSystemObject
fso.DeleteFolder App.Path & "\" & "Sandbox", True
End Sub

Private Sub txtCod_Change()
txtNome.Text = ""
lvMain.ListItems.Clear
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub TravaInclusao()
txtCod.Enabled = True
btBusca.Enabled = True
lvMain.Enabled = True
btDel.Enabled = True
btAdd.Enabled = True
btAbrir.Enabled = True
cmbTipo.ListIndex = -1
cmbTipo.Enabled = False
btGravar.Enabled = False
btCancel.Enabled = False
btArquivo.Enabled = False
txtObs.Enabled = False
txtArquivo.Text = ""
txtObs.Text = ""
cmbTipo.BackColor = Me.BackColor
txtObs.BackColor = Me.BackColor

End Sub

