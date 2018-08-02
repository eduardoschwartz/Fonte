VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAtiv 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atividades para Taxa de Licença"
   ClientHeight    =   5730
   ClientLeft      =   2610
   ClientTop       =   2970
   ClientWidth     =   8040
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   3990
      TabIndex        =   20
      ToolTipText     =   "Imprimir dados da lista"
      Top             =   4020
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmAtiv.frx":0000
      PICN            =   "frmAtiv.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   315
      Left            =   6630
      TabIndex        =   15
      ToolTipText     =   "Cancelar Edição"
      Top             =   4020
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmAtiv.frx":0176
      PICN            =   "frmAtiv.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConsultar 
      Height          =   315
      Left            =   5340
      TabIndex        =   16
      ToolTipText     =   "Retorna atividade selecionada para o cadastro de empresa"
      Top             =   4020
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Selecionar"
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
      MICON           =   "frmAtiv.frx":02EC
      PICN            =   "frmAtiv.frx":0308
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEdit 
      Height          =   315
      Left            =   1470
      TabIndex        =   17
      ToolTipText     =   "Alterar atividade existente"
      Top             =   4020
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmAtiv.frx":0376
      PICN            =   "frmAtiv.frx":0392
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
      Left            =   2760
      TabIndex        =   18
      ToolTipText     =   "Excluir atividade selecionada"
      Top             =   4020
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmAtiv.frx":04EC
      PICN            =   "frmAtiv.frx":0508
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   180
      TabIndex        =   19
      ToolTipText     =   "Cadastrar nova atividade"
      Top             =   4020
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmAtiv.frx":05AA
      PICN            =   "frmAtiv.frx":05C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   5
      TabIndex        =   3
      Top             =   4620
      Width           =   1095
   End
   Begin VB.TextBox txtAliq3 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5610
      MaxLength       =   10
      TabIndex        =   7
      Top             =   5340
      Width           =   1095
   End
   Begin VB.TextBox txtAliq2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3390
      MaxLength       =   10
      TabIndex        =   6
      Top             =   5340
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   300
      TabIndex        =   4
      Top             =   4980
      Width           =   5565
   End
   Begin VB.TextBox txtAliq1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   5
      Top             =   5340
      Width           =   1095
   End
   Begin prjChameleon.chameleonButton cmdNext 
      Height          =   315
      Left            =   6240
      TabIndex        =   1
      ToolTipText     =   "Localiza próxima ocorrência"
      Top             =   60
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Localizar/Próximo"
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
      MICON           =   "frmAtiv.frx":0720
      PICN            =   "frmAtiv.frx":073C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      MaxLength       =   20
      TabIndex        =   0
      Top             =   60
      Width           =   6105
   End
   Begin MSComctlLib.ListView lvAtiv 
      Height          =   3435
      Left            =   30
      TabIndex        =   2
      Top             =   450
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6059
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1340
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição da Atividade"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Aliquota 1"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Aliquota 2"
         Object.Width           =   1762
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Aliquota 3"
         Object.Width           =   1762
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdCancelEdit 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6870
      TabIndex        =   9
      ToolTipText     =   "Cancelar Edição"
      Top             =   5220
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
      MICON           =   "frmAtiv.frx":0896
      PICN            =   "frmAtiv.frx":08B2
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
      Left            =   6870
      TabIndex        =   8
      ToolTipText     =   "Gravar os Dados"
      Top             =   4830
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmAtiv.frx":0A0C
      PICN            =   "frmAtiv.frx":0A28
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Código....:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   14
      Top             =   4650
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota 3:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   13
      Top             =   5370
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota 2:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2340
      TabIndex        =   12
      Top             =   5370
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividade.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   11
      Top             =   5010
      Width           =   1035
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota 1:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   5370
      Width           =   1035
   End
End
Attribute VB_Name = "frmAtiv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String, Evento As String
Dim NomeForm As String, nPos As Integer

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancelEdit_Click()
HabilitaBotao
Me.Height = 4920
Centraliza Me
End Sub

Private Sub cmdConsultar_Click()
Dim Achou As Boolean
Dim nArea As Variant, nQtde As Variant

If NomeForm = "frmCadMob" Then

    'Se tiver 3 valores
    If CDbl(lvAtiv.SelectedItem.SubItems(4)) > 0 Then
        ButtonText(0) = lvAtiv.SelectedItem.SubItems(2)
        ButtonText(1) = lvAtiv.SelectedItem.SubItems(3)
        ButtonText(2) = lvAtiv.SelectedItem.SubItems(4)
        'Set up the CBT hook
        hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread)
        Retval = MsgBox("Selecione a aliquota a ser utilizada.", vbInformation + vbAbortRetryIgnore, "Seleção de Aliquota")
        If Retval = vbAbort Then 'valor 1
           frmCadMob.txtValorAliq.text = FormatNumber(lvAtiv.SelectedItem.SubItems(2), 2)
           frmCadMob.lblAliq.Caption = 1
        ElseIf Retval = vbRetry Then 'valor 2
           frmCadMob.txtValorAliq.text = FormatNumber(lvAtiv.SelectedItem.SubItems(3), 2)
           frmCadMob.lblAliq.Caption = 2
        ElseIf Retval = vbIgnore Then 'valor 3
           frmCadMob.txtValorAliq.text = FormatNumber(lvAtiv.SelectedItem.SubItems(4), 2)
           frmCadMob.lblAliq.Caption = 3
        End If
    'se tiver 2 valores (pouco provavel)
    ElseIf CDbl(lvAtiv.SelectedItem.SubItems(4)) = 0 And CDbl(lvAtiv.SelectedItem.SubItems(3)) > 0 Then
        ButtonText(0) = lvAtiv.SelectedItem.SubItems(2)
        ButtonText(1) = lvAtiv.SelectedItem.SubItems(3)
        'Set up the CBT hook
        hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread)
        Retval = MsgBox("Selecione a aliquota a ser utilizada.", vbInformation + vbYesNo, "Seleção de Aliquota")
        If Retval = vbYes Then 'valor 1
           frmCadMob.txtValorAliq.text = FormatNumber(lvAtiv.SelectedItem.SubItems(2), 2)
           frmCadMob.lblAliq.Caption = 1
        ElseIf Retval = vbNo Then 'valor 2
           frmCadMob.txtValorAliq.text = FormatNumber(lvAtiv.SelectedItem.SubItems(3), 2)
           frmCadMob.lblAliq.Caption = 2
        End If
    'se tiver um ou nenhum valor
    Else
        frmCadMob.txtValorAliq.text = FormatNumber(lvAtiv.SelectedItem.SubItems(2), 2)
        frmCadMob.lblAliq.Caption = 1
    End If
    
   frmCadMob.txtAtiv.text = lvAtiv.SelectedItem.text & " - " & lvAtiv.SelectedItem.SubItems(1)
   If frmCadMob.txtAtivExt.text = "" Then
      frmCadMob.txtAtivExt.text = lvAtiv.SelectedItem.SubItems(1)
   End If
   CodEmpresa = 0
   Unload Me
ElseIf NomeForm = "frmOutraAtividade" Then
    Achou = False
    If Val(lvAtiv.SelectedItem.text) = Val(Left(frmCadMob.txtAtiv.text, 5)) Then
       MsgBox "Atividade já cadastrada.", vbExclamation, "Atenção"
       Exit Sub
    End If
    With frmOutraAtividade.grdAtiv
        For x = 1 To .Rows - 1
            If Val(lvAtiv.SelectedItem.text) = Val(.TextMatrix(x, 0)) Then
               Achou = True
               Exit For
            End If
        Next
    End With
    If Achou Then
       MsgBox "Atividade já cadastrada.", vbExclamation, "Atenção"
       Exit Sub
    End If
    nArea = InputBox("Digite a Área para esta atividade.", "Definição de Área", "1")
    If Not IsNumeric(nArea) Then nArea = 0
    If CDbl(nArea) = 0 Then nArea = 1
    nQtde = InputBox("Digite a Qtde de profissionais para esta atividade.", "Quantidade de Profissionais", "1")
    If Not IsNumeric(nQtde) Then nQtde = 0
    If CDbl(nQtde) = 0 Then nQtde = 1
    
    'Se tiver 3 valores
    If CDbl(lvAtiv.SelectedItem.SubItems(4)) > 0 Then
        ButtonText(0) = lvAtiv.SelectedItem.SubItems(2)
        ButtonText(1) = lvAtiv.SelectedItem.SubItems(3)
        ButtonText(2) = lvAtiv.SelectedItem.SubItems(4)
        'Set up the CBT hook
        hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread)
        Retval = MsgBox("Selecione a aliquota a ser utilizada.", vbInformation + vbAbortRetryIgnore, "Seleção de Aliquota")
        If Retval = vbAbort Then 'valor 1
           frmOutraAtividade.grdAtiv.AddItem lvAtiv.SelectedItem.text & Chr(9) & lvAtiv.SelectedItem.SubItems(1) & Chr(9) & FormatNumber(lvAtiv.SelectedItem.SubItems(2), 2) & Chr(9) & 1 & Chr(9) & FormatNumber(nArea, 2) & Chr(9) & nQtde
        ElseIf Retval = vbRetry Then 'valor 2
           frmOutraAtividade.grdAtiv.AddItem lvAtiv.SelectedItem.text & Chr(9) & lvAtiv.SelectedItem.SubItems(1) & Chr(9) & FormatNumber(lvAtiv.SelectedItem.SubItems(3), 2) & Chr(9) & 2 & Chr(9) & FormatNumber(nArea, 2) & Chr(9) & nQtde
        ElseIf Retval = vbIgnore Then 'valor 3
           frmOutraAtividade.grdAtiv.AddItem lvAtiv.SelectedItem.text & Chr(9) & lvAtiv.SelectedItem.SubItems(1) & Chr(9) & FormatNumber(lvAtiv.SelectedItem.SubItems(4), 2) & Chr(9) & 3 & Chr(9) & FormatNumber(nArea, 2) & Chr(9) & nQtde
        End If
    'se tiver 2 valores (pouco provavel)
    ElseIf CDbl(lvAtiv.SelectedItem.SubItems(4)) = 0 And CDbl(lvAtiv.SelectedItem.SubItems(3)) > 0 Then
        ButtonText(0) = lvAtiv.SelectedItem.SubItems(2)
        ButtonText(1) = lvAtiv.SelectedItem.SubItems(3)
        'Set up the CBT hook
        hInst = GetWindowLong(Me.hwnd, GWL_HINSTANCE)
        Thread = GetCurrentThreadId()
        hHook = SetWindowsHookEx(WH_CBT, AddressOf Manipulate, hInst, Thread)
        Retval = MsgBox("Selecione a aliquota a ser utilizada.", vbInformation + vbYesNo, "Seleção de Aliquota")
        If Retval = vbYes Then 'valor 1
           frmOutraAtividade.grdAtiv.AddItem lvAtiv.SelectedItem.text & Chr(9) & lvAtiv.SelectedItem.SubItems(1) & Chr(9) & FormatNumber(lvAtiv.SelectedItem.SubItems(2), 2) & Chr(9) & 1 & Chr(9) & FormatNumber(nArea, 2) & Chr(9) & nQtde
        ElseIf Retval = vbNo Then 'valor 2
           frmOutraAtividade.grdAtiv.AddItem lvAtiv.SelectedItem.text & Chr(9) & lvAtiv.SelectedItem.SubItems(1) & Chr(9) & FormatNumber(lvAtiv.SelectedItem.SubItems(3), 2) & Chr(9) & 2 & Chr(9) & FormatNumber(nArea, 2) & Chr(9) & nQtde
        End If
    'se tiver um ou nenhum valor
    Else
        frmOutraAtividade.grdAtiv.AddItem lvAtiv.SelectedItem.text & Chr(9) & lvAtiv.SelectedItem.SubItems(1) & Chr(9) & FormatNumber(lvAtiv.SelectedItem.SubItems(2), 2) & Chr(9) & 1 & Chr(9) & FormatNumber(nArea, 2) & Chr(9) & nQtde
    End If
    CodEmpresa = 0
    Unload Me
Else
   MsgBox "A tela de Cadastro Mobiliário não esta ativa.", vbExclamation, "Atenção"
End If

End Sub

Private Sub cmdEdit_Click()
Evento = "Alterar"
DesabilitaBotao
txtCod.Enabled = False
Me.Height = 6165
Centraliza Me
With lvAtiv
    txtCod.text = .SelectedItem.text
    txtDesc.text = .SelectedItem.SubItems(1)
    txtAliq1.text = .SelectedItem.SubItems(2)
    txtAliq2.text = .SelectedItem.SubItems(3)
    txtAliq3.text = .SelectedItem.SubItems(4)
End With

txtDesc.SetFocus
End Sub

Private Sub cmdExcluir_Click()
Dim n As Long

n = lvAtiv.SelectedItem.text

Sql = "SELECT CODIGOMOB FROM MOBILIARIO WHERE CODATIVIDADE=" & n
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
       MsgBox "Não é possível excluir esta atividade pois existem empresas cadastradas com esta atividade.", vbExclamation, "Atenção"
       Exit Sub
    End If
   .Close
End With

If MsgBox("Excluir esta atividade ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
   Sql = "DELETE FROM ATIVIDADE WHERE CODATIVIDADE=" & n
   cn.Execute Sql, rdExecDirect
   CarregaLista
End If

End Sub

Private Sub cmdGravar_Click()
Dim Achou As Boolean

If Len(txtCod.text) <> 5 Then
   MsgBox "Código inválido, 5 digitos requeridos.", vbExclamation, "Atenção"
   txtCod.SetFocus
   Exit Sub
End If

If Evento = "Novo" Then
    t = txtCod.text
    Achou = False
    With lvAtiv
        For x = 1 To .ListItems.Count
            s = UCase$(.ListItems(x).text)
            If InStr(1, s, t, vbBinaryCompare) > 0 Then
                Achou = True
                Exit For
            End If
        Next
    End With
    If Achou Then
       MsgBox "Código já cadastrado.", vbExclamation, "Atenção"
       txtCod.SetFocus
       Exit Sub
    End If
End If

If Trim$(txtDesc.text) = "" Then
   MsgBox "Digite a descrição da atividade.", vbExclamation, "Atenção"
   txtDesc.SetFocus
   Exit Sub
End If

If Trim$(txtAliq1.text) = "" Then txtAliq1.text = "0,00"
If Trim$(txtAliq2.text) = "" Then txtAliq2.text = "0,00"
If Trim$(txtAliq3.text) = "" Then txtAliq3.text = "0,00"

If (CDbl(txtAliq2.text) > 0 Or CDbl(txtAliq3.text) > 0) And CDbl(txtAliq1.text) = 0 Then
   MsgBox "Não é possível ter aliquotas 2 e/ou 3 sem aliquota 1.", vbExclamation, "Atenção"
   Exit Sub
End If

If CDbl(txtAliq3.text) > 0 And CDbl(txtAliq2.text) = 0 Then
   MsgBox "Não é possível ter aliquota 3 sem aliquota 2.", vbExclamation, "Atenção"
   Exit Sub
End If

If Evento = "Novo" Then
   Sql = "INSERT ATIVIDADE (CODATIVIDADE,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3) VALUES("
   Sql = Sql & txtCod.text & ",'" & Mask(txtDesc.text) & "'," & Virg2Ponto(txtAliq1.text) & ","
   Sql = Sql & Virg2Ponto(txtAliq2.text) & "," & Virg2Ponto(txtAliq3.text) & ")"
Else
   Sql = "UPDATE ATIVIDADE SET DESCATIVIDADE='" & Mask(txtDesc.text) & "',VALORALIQ1=" & Virg2Ponto(txtAliq1.text)
   Sql = Sql & ",VALORALIQ2=" & Virg2Ponto(txtAliq2.text) & ",VALORALIQ3=" & Virg2Ponto(txtAliq3.text)
   Sql = Sql & " WHERE CODATIVIDADE=" & txtCod.text
End If
cn.Execute Sql, rdExecDirect
CarregaLista
HabilitaBotao
Me.Height = 4920
Centraliza Me
End Sub

Private Sub cmdNext_Click()
Dim R As String
t = UCase$(Trim$(txtSearch.text))
Inicio:
nPos = nPos + 1

With lvAtiv
    For x = nPos To .ListItems.Count
        s = UCase$(.ListItems(x).SubItems(1))
        R = .ListItems(x).text
        If InStr(1, s, t, vbBinaryCompare) > 0 Then
            .ListItems(x).Selected = True
            .SetFocus
            nPos = x
            'AUTOSCROLL
            For Y = 1 To 5
               SendKeys "{" & Mid(.ListItems(x).text, Y, 1) & "}"
            Next
            Exit For
        End If
    Next
End With

If x >= lvAtiv.ListItems.Count Then
   If MsgBox("Fim da Pesquisa." & vbCrLf & "Deseja pesquisar do início ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
      nPos = 0
      GoTo Inicio
   End If
End If

End Sub

Private Sub cmdNovo_Click()
Evento = "Novo"
DesabilitaBotao
Me.Height = 6165
Centraliza Me
txtCod.Enabled = True
txtCod.text = ""
txtDesc.text = ""
txtAliq1.text = ""
txtAliq2.text = ""
txtAliq3.text = ""
txtCod.SetFocus
End Sub

Private Sub DesabilitaBotao()
cmdNovo.Enabled = False
cmdEdit.Enabled = False
cmdExcluir.Enabled = False
cmdConsultar.Enabled = False
cmdCancel.Enabled = False
End Sub

Private Sub HabilitaBotao()
cmdNovo.Enabled = True
cmdEdit.Enabled = True
cmdExcluir.Enabled = True
cmdConsultar.Enabled = True
cmdCancel.Enabled = True
End Sub

Private Sub cmdPrint_Click()
'EXIBE RELATORIO
frmReport.ShowReport "ALIQATIVIDADETL", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub Form_Activate()
If NomeForm = "frmCadMob" Or NomeForm = "frmOutraAtividade" Then
   cmdConsultar.Enabled = True
Else
   cmdConsultar.Enabled = False
End If
End Sub

Private Sub Form_Load()

Me.Height = 4920

Ocupado

CarregaLista
Centraliza Me

Liberado
End Sub

Private Sub CarregaLista()
Dim itmX As ListItem
Dim z As Long
z = SendMessage(lvAtiv.hwnd, LVM_DELETEALLITEMS, 0, 0)

Sql = "SELECT CODATIVIDADE,DESCATIVIDADE,VALORALIQ1,VALORALIQ2,VALORALIQ3 FROM ATIVIDADE ORDER BY DESCATIVIDADE "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       Set itmX = lvAtiv.ListItems.Add(, "K" & Format(!CODATIVIDADE, "00000"), Format(!CODATIVIDADE, "00000"))
       itmX.SubItems(1) = !DESCATIVIDADE
       itmX.SubItems(2) = FormatNumber(!VALORALIQ1, 2)
       itmX.SubItems(3) = FormatNumber(!VALORALIQ2, 2)
       itmX.SubItems(4) = FormatNumber(!VALORALIQ3, 2)
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub lvAtiv_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvAtiv.SortKey = ColumnHeader.Position - 1
lvAtiv.Sorted = True
lvAtiv.SortOrder = lvwAscending

End Sub

Private Sub txtAliq1_KeyPress(KeyAscii As Integer)
Tweak txtAliq1, KeyAscii, DecimalPositive
End Sub

Private Sub txtAliq2_KeyPress(KeyAscii As Integer)
Tweak txtAliq2, KeyAscii, DecimalPositive
End Sub

Private Sub txtAliq3_KeyPress(KeyAscii As Integer)
Tweak txtAliq3, KeyAscii, DecimalPositive
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)

If KeyAscii <> 32 Then

Tweak txtDesc, KeyAscii, AllLettersAllCaps
End If
End Sub

Private Sub txtSearch_Change()
nPos = 0
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdNext_Click
End If
End Sub
