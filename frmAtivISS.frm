VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmAtivISS 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atividades para ISS"
   ClientHeight    =   5715
   ClientLeft      =   4575
   ClientTop       =   3420
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5715
   ScaleWidth      =   8295
   Begin esMaskEdit.esMaskedEdit mskData 
      Height          =   285
      Left            =   2010
      TabIndex        =   4
      Top             =   5370
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmAtivISS.frx":0000
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
      Left            =   3990
      TabIndex        =   19
      ToolTipText     =   "Imprimir dados da lista"
      Top             =   4050
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
      MICON           =   "frmAtivISS.frx":001C
      PICN            =   "frmAtivISS.frx":0038
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
      Left            =   6600
      TabIndex        =   14
      ToolTipText     =   "Cancelar Edição"
      Top             =   4050
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
      MICON           =   "frmAtivISS.frx":0192
      PICN            =   "frmAtivISS.frx":01AE
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
      Left            =   5310
      TabIndex        =   15
      ToolTipText     =   "Retorna atividade selecionada para o cadastro de empresa"
      Top             =   4050
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
      MICON           =   "frmAtivISS.frx":0308
      PICN            =   "frmAtivISS.frx":0324
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
      Left            =   1440
      TabIndex        =   16
      ToolTipText     =   "Alterar atividade existente"
      Top             =   4050
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
      MICON           =   "frmAtivISS.frx":0392
      PICN            =   "frmAtivISS.frx":03AE
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
      Left            =   2730
      TabIndex        =   17
      ToolTipText     =   "Excluir atividade selecionada"
      Top             =   4050
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
      MICON           =   "frmAtivISS.frx":0508
      PICN            =   "frmAtivISS.frx":0524
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
      Left            =   150
      TabIndex        =   18
      ToolTipText     =   "Cadastrar nova atividade"
      Top             =   4050
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
      MICON           =   "frmAtivISS.frx":05C6
      PICN            =   "frmAtivISS.frx":05E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtAliq 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5475
      MaxLength       =   7
      TabIndex        =   2
      Top             =   4635
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmAtivISS.frx":073C
      Left            =   2865
      List            =   "frmAtivISS.frx":074C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4635
      Width           =   1500
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   30
      MaxLength       =   20
      TabIndex        =   8
      Top             =   90
      Width           =   6105
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1005
      MaxLength       =   300
      TabIndex        =   3
      Top             =   5010
      Width           =   5565
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1020
      MaxLength       =   5
      TabIndex        =   0
      Top             =   4620
      Width           =   1095
   End
   Begin prjChameleon.chameleonButton cmdNext 
      Height          =   315
      Left            =   6240
      TabIndex        =   7
      ToolTipText     =   "Localiza próxima ocorrência"
      Top             =   90
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
      MICON           =   "frmAtivISS.frx":076A
      PICN            =   "frmAtivISS.frx":0786
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ListView lvAtiv 
      Height          =   3435
      Left            =   30
      TabIndex        =   9
      Top             =   480
      Width           =   8235
      _ExtentX        =   14526
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Tipo"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição da Atividade"
         Object.Width           =   8203
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Aliquota"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Data"
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Seq"
         Object.Width           =   0
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdCancelEdit 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6750
      TabIndex        =   6
      ToolTipText     =   "Cancelar Edição"
      Top             =   5220
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmAtivISS.frx":08E0
      PICN            =   "frmAtivISS.frx":08FC
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
      Left            =   6750
      TabIndex        =   5
      ToolTipText     =   "Gravar os Dados"
      Top             =   4830
      Width           =   1215
      _ExtentX        =   2143
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
      MICON           =   "frmAtivISS.frx":0A56
      PICN            =   "frmAtivISS.frx":0A72
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
      Caption         =   "Seq....:"
      Height          =   255
      Index           =   4
      Left            =   3510
      TabIndex        =   22
      Top             =   5400
      Width           =   570
   End
   Begin VB.Label lblSeq 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   4110
      TabIndex        =   21
      Top             =   5400
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota válida a partir de:"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   5400
      Width           =   1860
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aliquota..:"
      Height          =   255
      Index           =   1
      Left            =   4650
      TabIndex        =   13
      Top             =   4680
      Width           =   780
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo..:"
      Height          =   255
      Index           =   0
      Left            =   2295
      TabIndex        =   12
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividade.:"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   5040
      Width           =   870
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Código....:"
      Height          =   255
      Index           =   3
      Left            =   90
      TabIndex        =   10
      Top             =   4680
      Width           =   870
   End
End
Attribute VB_Name = "frmAtivISS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String, Evento As String
Dim NomeForm As String, nPos As Integer
Dim AreaISS As Double, TipoISS As Double, Cancelado As Boolean

Public Property Let bCancel(bCancelado As Boolean)
    Cancelado = bCancelado
End Property

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Public Property Let nArea(nAreaTotal As Double)
    AreaISS = nAreaTotal
End Property

Public Property Let nTipo(nTipoIss As Double)
    TipoISS = nTipoIss
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

Dim x As Integer, Achou As Boolean
Dim nQtde As Variant, sTipoIss As String

If lvAtiv.SelectedItem.Text = "" Then
    MsgBox "Selecione uma Atividade.", vbExclamation, "Atenção"
    Exit Sub
End If

If NomeForm = "frmCadMob" Then
    Achou = False
    With frmCadMob.grdAtiv
       For x = 1 To .Rows - 1
           If Val(Left$(.TextMatrix(x, 0), 4)) = Val(lvAtiv.SelectedItem.Text) Then
              Achou = True
              Exit For
           End If
       Next
    End With
    If Not Achou Then
       nQtde = InputBox("Digite a Qtde de profissionais para esta atividade.", "Quantidade de Profissionais", "1")
       If Not IsNumeric(nQtde) Then nQtde = 0
       If CDbl(nQtde) = 0 Then nQtde = 1
       If TipoISS <> 12 Then
          If TipoISS = 13 Then
             sTipoIss = "V"
          Else
             sTipoIss = "F"
          End If
          frmCadMob.grdAtiv.AddItem sTipoIss & Chr(9) & Format(lvAtiv.SelectedItem.Text, "0000") & " - " & lvAtiv.SelectedItem.SubItems(2) & Chr(9) & nQtde & Chr(9) & lvAtiv.SelectedItem.SubItems(3)
       Else
          sTipoIss = "E"
          nValor = InputBox("Digite o Valor Estimado para esta atividade.", "Valor Estimado", "1")
          If Not IsNumeric(nValor) Then nValor = 0
          If CDbl(nValor) = 0 Then nValor = 1
          frmCadMob.grdAtiv.AddItem sTipoIss & Chr(9) & Format(lvAtiv.SelectedItem.Text, "0000") & " - " & lvAtiv.SelectedItem.SubItems(2) & Chr(9) & nQtde & Chr(9) & FormatNumber(nValor, 2)
       End If
       CodEmpresa = 0
       Unload Me
    Else
       MsgBox "Esta atividade já foi atribuida à empresa.", vbExclamation, "Atenção"
    End If
    
End If

End Sub

Private Sub cmdEdit_Click()
'Exit Sub
Evento = "Alterar"
DesabilitaBotao
txtCod.Enabled = False
Me.Height = 6165
Centraliza Me
With lvAtiv
    txtCod.Text = .SelectedItem.Text
    Select Case Left$(.SelectedItem.SubItems(1), 1)
        Case "F"
            cmbTipo.Text = "Fixo"
        Case "E"
            cmbTipo.Text = "Estimado"
        Case "V"
            cmbTipo.Text = "Variável"
    End Select
    txtDesc.Text = .SelectedItem.SubItems(2)
    txtAliq.Text = .SelectedItem.SubItems(3)
    mskData.Text = .SelectedItem.SubItems(4)
    lblSeq.Caption = .SelectedItem.SubItems(5)
End With

txtDesc.SetFocus
End Sub

Private Sub cmdExcluir_Click()
Dim n As Long, nTipo As Integer
'Exit Sub
n = lvAtiv.SelectedItem.Text


Select Case Left$(lvAtiv.SelectedItem.SubItems(1), 1)
    Case "F"
        nTipo = 11
    Case "E"
        nTipo = 12
    Case "V"
        nTipo = 13
End Select

Sql = "SELECT CODMOBILIARIO FROM MOBILIARIOATIVIDADEISS WHERE CODATIVIDADE=" & n
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
       MsgBox "Não é possível excluir esta atividade pois existem empresas cadastradas com esta atividade.", vbExclamation, "Atenção"
       Exit Sub
    End If
   .Close
End With

If MsgBox("Excluir esta atividade ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
   Sql = "DELETE FROM TABELAISS WHERE TIPOISS=" & nTipo & " AND CODIGOATIV=" & n
   cn.Execute Sql, rdExecDirect
   Sql = "DELETE FROM ATIVIDADEISS WHERE CODATIVIDADE=" & n
   cn.Execute Sql, rdExecDirect
   CarregaLista
End If

End Sub

Private Sub cmdGravar_Click()
Dim t As Long, nCodOld As Long, nSeq As Integer

If Evento = "Novo" Then
    Sql = "SELECT MAX(CODATIVIDADE) AS MAXIMO FROM ATIVIDADEISS"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        t = !maximo + 1
        txtCod.Text = Format(t, "00000")
       .Close
    End With
'    Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM TABELAISS WHERE TIPOISS=" & cmbTipo.ItemData(cmbTipo.ListIndex)
'    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'    With RdoAux
'        nCodOld = !MAXIMO + 1
'       .Close
'    End With
End If

If CDbl(txtAliq.Text) = 0 Then
   MsgBox "Digite o valor da alíquota.", vbExclamation, "Atenção"
   txtAliq.SetFocus
   Exit Sub
End If

If cmbTipo.ListIndex = -1 Then
   MsgBox "Selecione o tipo de ISS.", vbExclamation, "Atenção"
   cmbTipo.SetFocus
   Exit Sub
End If

If Trim$(txtDesc.Text) = "" Then
   MsgBox "Digite a descrição da atividade.", vbExclamation, "Atenção"
   txtDesc.SetFocus
   Exit Sub
End If

If Not IsDate(mskData.Text) Then
    MsgBox "Data inválida.", vbCritical, "Atenção"
    Exit Sub
End If

If Evento = "Novo" Then
   Sql = "INSERT ATIVIDADEISS (CODATIVIDADE,DESCATIVIDADE) VALUES("
   Sql = Sql & t & ",'" & Mask(txtDesc.Text) & "')"
   cn.Execute Sql, rdExecDirect
   Sql = "INSERT TABELAISS (TIPOISS,CODIGOATIV,SEQ,DESCATIV,ALIQUOTA,DATA) VALUES("
   Sql = Sql & cmbTipo.ItemData(cmbTipo.ListIndex) & "," & t & "," & nSeq & ",'"
   Sql = Sql & Mask(txtDesc.Text) & "'," & Virg2Ponto(txtAliq.Text) & ",'" & Format(mskData.Text, "mm/dd/yyyy") & "')"
   cn.Execute Sql, rdExecDirect
Else
   Sql = "UPDATE ATIVIDADEISS SET DESCATIVIDADE='" & Mask(txtDesc.Text) & "' "
   Sql = Sql & "WHERE CODATIVIDADE=" & txtCod.Text
   cn.Execute Sql, rdExecDirect
   Sql = "UPDATE TABELAISS SET DESCATIV='" & Mask(txtDesc.Text) & "', ALIQUOTA=" & Virg2Ponto(txtAliq.Text) & ",DATA='" & Format(mskData.Text, "mm/dd/yyyy") & "'"
   Sql = Sql & " WHERE TIPOISS=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND CODIGOATIV=" & txtCod.Text & " AND SEQ=" & Val(lblSeq.Caption)
   cn.Execute Sql, rdExecDirect
End If

CarregaLista
HabilitaBotao
Me.Height = 4920
Centraliza Me
End Sub


Private Sub cmdNext_Click()
t = UCase$(Trim$(txtSearch.Text))
Inicio:
nPos = nPos + 1

With lvAtiv
    For x = nPos To .ListItems.Count
        s = UCase$(.ListItems(x).SubItems(2))
        If InStr(1, s, t, vbBinaryCompare) > 0 Then
            .ListItems(x).Selected = True
            .SetFocus
            .ListItems(x).EnsureVisible
            nPos = x
            'AUTOSCROLL
 '           For Y = 1 To 5
'               SendKeys "{" & Mid(.ListItems(x).Text, Y, 1) & "}"
  '          Next
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
'Exit Sub
Evento = "Novo"
DesabilitaBotao
Me.Height = 6165
Centraliza Me
txtCod.Enabled = False
txtCod.Text = ""
txtDesc.Text = ""
LimpaMascara mskData
txtDesc.SetFocus
End Sub

Private Sub DesabilitaBotao()
cmdNovo.Enabled = False
cmdEdit.Enabled = False
cmdExcluir.Enabled = False
cmdConsultar.Enabled = False
cmdCancel.Enabled = False
lvAtiv.Enabled = False
End Sub

Private Sub HabilitaBotao()
cmdNovo.Enabled = True
cmdEdit.Enabled = True
cmdExcluir.Enabled = True
cmdConsultar.Enabled = True
cmdCancel.Enabled = True
lvAtiv.Enabled = True
End Sub

Private Sub cmdPrint_Click()
'EXIBE RELATORIO
frmReport.ShowReport "ALIQATIVIDADEISS", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub Form_Activate()
If NomeForm = "frmCadMob" Then
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

Sql = "SELECT DISTINCT ATIVIDADEISS.CODATIVIDADE,ATIVIDADEISS.DESCATIVIDADE,TABELAISS.TIPOISS "
Sql = Sql & "FROM ATIVIDADEISS INNER JOIN TABELAISS ON ATIVIDADEISS.CODATIVIDADE = TABELAISS.CODIGOATIV "
If TipoISS > 10 Then
    Sql = Sql & "WHERE TIPOISS=" & TipoISS
End If
Sql = Sql & " ORDER BY ATIVIDADEISS.DESCATIVIDADE "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvAtiv.ListItems.Add(, "K" & CStr(!TipoISS) & Format(!codatividade, "00000"), Format(!codatividade, "00000"))
       If !TipoISS = 11 Then
          itmX.SubItems(1) = "Fixo"
       ElseIf !TipoISS = 12 Then
          itmX.SubItems(1) = "Est."
       ElseIf !TipoISS = 13 Then
          itmX.SubItems(1) = "Var."
       End If
       itmX.SubItems(2) = !descatividade
       'itmX.SubItems(3) = FormatNumber(!Aliquota, 3)
       'If !codatividade = 513 Then MsgBox "teste"
       
       itmX.SubItems(3) = RetornaAliquotaISS(!codatividade, Format(Now, "dd/mm/yyyy"))
'       itmX.SubItems(4) = Format(!Data, "dd/mm/yyyy")
'       itmX.SubItems(5) = !Seq
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

Private Sub txtAliq_KeyPress(KeyAscii As Integer)

Tweak txtAliq, KeyAscii, DecimalPositive
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
Tweak txtCod, KeyAscii, IntegerPositive
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

