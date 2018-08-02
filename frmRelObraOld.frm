VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmRelatObra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatórios de Atendimento"
   ClientHeight    =   2475
   ClientLeft      =   6510
   ClientTop       =   3105
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2475
   ScaleWidth      =   4710
   Begin VB.ComboBox cmbSit 
      Height          =   315
      ItemData        =   "frmRelObraOld.frx":0000
      Left            =   135
      List            =   "frmRelObraOld.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2025
      Width           =   1770
   End
   Begin VB.ComboBox cmbEquipe 
      Height          =   315
      ItemData        =   "frmRelObraOld.frx":003F
      Left            =   90
      List            =   "frmRelObraOld.frx":0041
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1350
      Visible         =   0   'False
      Width           =   4515
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmRelObraOld.frx":0043
      Left            =   90
      List            =   "frmRelObraOld.frx":0045
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   675
      Width           =   4515
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   180
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelObraOld.frx":0047
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   3570
      TabIndex        =   1
      Top             =   195
      Width           =   1020
      _ExtentX        =   1799
      _ExtentY        =   503
      MouseIcon       =   "frmRelObraOld.frx":0063
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   3555
      TabIndex        =   5
      ToolTipText     =   "Imprimir esta Tela"
      Top             =   2025
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmRelObraOld.frx":007F
      PICN            =   "frmRelObraOld.frx":009B
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
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº da Ordem"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Dt.Entrada"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Dt.Exec."
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Assunto"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Equipe"
         Object.Width           =   2542
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Situação"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label lblEquipe 
      BackStyle       =   0  'Transparent
      Caption         =   "Equipe"
      Height          =   195
      Left            =   135
      TabIndex        =   8
      Top             =   1125
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Início..:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   225
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Fim.....:"
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   2550
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmRelatObra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTipo_Click()
If cmbTipo.ListIndex = 1 Then
    lblEquipe.Visible = True
    cmbEquipe.Visible = True
    cmbEquipe.ListIndex = 0
Else
    lblEquipe.Visible = False
    cmbEquipe.Visible = False
End If
End Sub

Private Sub cmdPrint_Click()
If Not IsDate(mskDataIni.Text) Then
    MsgBox "Data de Inicio inválido", vbExclamation, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataFim.Text) Then
    MsgBox "Data de Fim inválido", vbExclamation, "atenção"
    Exit Sub
End If

If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    MsgBox "Data de Inicio tem que ser maior que data de termino", vbExclamation, "atenção"
    Exit Sub
End If

Select Case cmbTipo.ItemData(cmbTipo.ListIndex)
    Case 1
        frmReport.ShowReport2 "REGATENDIMENTO1", frmMdi.hwnd, Me.hwnd
    Case 2
        frmReport.ShowReport2 "REGATENDIMENTO3", frmMdi.hwnd, Me.hwnd
    Case 3
        frmReport.ShowReport2 "REGATENDIMENTO2", frmMdi.hwnd, Me.hwnd
    Case 4
        frmReport.ShowReport2 "REGATENDIMENTO4", frmMdi.hwnd, Me.hwnd
    Case 5
        GeraResumo
End Select

End Sub

Private Sub Form_Load()
Centraliza Me

cmbTipo.AddItem "Atendimentos por Assunto"
cmbTipo.ItemData(cmbTipo.NewIndex) = 1
cmbTipo.AddItem "Atendimentos por Equipe"
cmbTipo.ItemData(cmbTipo.NewIndex) = 2
cmbTipo.AddItem "Resumo dos Atendimentos"
cmbTipo.ItemData(cmbTipo.NewIndex) = 3
cmbTipo.AddItem "Gráfico dos Atendimentos"
cmbTipo.ItemData(cmbTipo.NewIndex) = 4
cmbTipo.AddItem "Resumo dos atendimentos 2"
cmbTipo.ItemData(cmbTipo.NewIndex) = 5

cmbTipo.ListIndex = 0
CarregaEquipe
cmbSit.ListIndex = 0
End Sub

Private Sub CarregaEquipe()
Dim Sql As String, RdoAux As rdoResultset

cmbEquipe.AddItem "(Todas as Equipes)"
cmbEquipe.ItemData(cmbEquipe.NewIndex) = 0

Sql = "SELECT CODIGO, NOME From paramobra WHERE SIGLA = 'EQ'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbEquipe.AddItem !nome
        cmbEquipe.ItemData(cmbEquipe.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub mskDataFim_GotFocus()
mskDataFim.SelStart = 0
mskDataFim.SelLength = Len(mskDataFim.Text)
End Sub

Private Sub mskDataIni_GotFocus()
mskDataIni.SelStart = 0
mskDataIni.SelLength = Len(mskDataIni.Text)
End Sub

Private Sub GeraResumo()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim itmX As ListItem, sSit As String, x As Integer
Dim z As Long
z = SendMessage(lvMain.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT * from registroatendimento where data between '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"


'if cmbSit.ListIndex
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cmbSit.ListIndex > 0 Then
            If IsDate(!dataend) Then
                If cmbSit.ListIndex = 1 Then
                    sSit = "CONCLUIDO"
                Else
                    GoTo proximo
                End If
            Else
                If IsDate(!DataCancel) Then
                    If cmbSit.ListIndex = 3 Then
                        sSit = "CANCELADO"
                    Else
                        GoTo proximo
                    End If
                Else
                    If cmbSit.ListIndex = 2 Then
                        sSit = "AGUARDANDO"
                    Else
                        GoTo proximo
                    End If
                End If
            End If
        End If
    
        Set itmX = lvMain.ListItems.Add(, , Format(!numreg, "0000") & "/" & !anoreg)
        itmX.SubItems(1) = Format(!Data, "dd/mm/yyyy")
        If IsDate(!Dataexec) Then
            itmX.SubItems(2) = Format(!Dataexec, "dd/mm/yyyy")
        Else
            itmX.SubItems(2) = ""
        End If
        itmX.SubItems(3) = SubNull(!assunto)
        Sql = "select nome from paramobra where sigla='EQ' and codigo=" & !equipe
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        itmX.SubItems(4) = SubNull(RdoAux2!nome)
        RdoAux2.Close
        If IsDate(!dataend) Then
            sSit = "CONCLUIDO"
        Else
            If IsDate(!DataCancel) Then
                sSit = "CANCELADO"
            Else
                sSit = "AGUARDANDO"
            End If
        End If
        itmX.SubItems(5) = sSit
proximo:
       .MoveNext
    Loop
   .Close
End With
Ocupado
Sql = "delete from relatorio_obra where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

For x = 1 To lvMain.ListItems.Count
    Sql = "insert relatorio_obra(usuario,ordem,dtentrada,dtexec,assunto,equipe,status) values('"
    Sql = Sql & NomeDeLogin & "','" & lvMain.ListItems(x).Text & "','" & Format(lvMain.ListItems(x).SubItems(1), "mm/dd/yyyy") & "',"
    Sql = Sql & IIf(IsDate(lvMain.ListItems(x).SubItems(2)), "'" & Format(lvMain.ListItems(x).SubItems(2), "mm/dd/yyyy") & "'", "Null") & ",'"
    Sql = Sql & Mask(Left(lvMain.ListItems(x).SubItems(3), 200)) & "','" & lvMain.ListItems(x).SubItems(4) & "','" & lvMain.ListItems(x).SubItems(5) & "')"
    cn.Execute Sql, rdExecDirect
Next
Liberado
frmReport.ShowReport2 "REGATENDIMENTO5", frmMdi.hwnd, Me.hwnd
Sql = "delete from relatorio_obra where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect


End Sub
