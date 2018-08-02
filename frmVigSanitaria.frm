VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmVigSanitaria 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta a Taxa de Vigilância Sanitária"
   ClientHeight    =   5985
   ClientLeft      =   615
   ClientTop       =   2310
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   7320
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   2700
      TabIndex        =   14
      ToolTipText     =   "Imprimir dados da lista"
      Top             =   4290
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmVigSanitaria.frx":0000
      PICN            =   "frmVigSanitaria.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancelEdit 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6180
      TabIndex        =   6
      ToolTipText     =   "Cancelar Edição"
      Top             =   5610
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
      MICON           =   "frmVigSanitaria.frx":0176
      PICN            =   "frmVigSanitaria.frx":0192
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
      Left            =   5100
      TabIndex        =   5
      ToolTipText     =   "Gravar os Dados"
      Top             =   5610
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
      MICON           =   "frmVigSanitaria.frx":02EC
      PICN            =   "frmVigSanitaria.frx":0308
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
      Left            =   1140
      TabIndex        =   4
      Top             =   5610
      Width           =   1095
   End
   Begin VB.TextBox txtSub 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   3
      Top             =   5250
      Width           =   6075
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1140
      TabIndex        =   2
      Top             =   4890
      Width           =   6075
   End
   Begin MSComctlLib.ImageList imgIcon 
      Left            =   4650
      Top             =   1470
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
            Picture         =   "frmVigSanitaria.frx":06AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVigSanitaria.frx":0809
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvVig 
      Height          =   4170
      Left            =   -30
      TabIndex        =   0
      Top             =   0
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   7355
      _Version        =   393217
      Indentation     =   794
      LabelEdit       =   1
      Style           =   3
      HotTracking     =   -1  'True
      ImageList       =   "imgIcon"
      Appearance      =   1
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Height          =   315
      Left            =   4950
      TabIndex        =   11
      ToolTipText     =   "Fechar a Tela"
      Top             =   4290
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Fechar"
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
      MICON           =   "frmVigSanitaria.frx":10E5
      PICN            =   "frmVigSanitaria.frx":1101
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
      Left            =   6030
      TabIndex        =   12
      ToolTipText     =   "Retorna Taxa Selecionada"
      Top             =   4290
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Selecionar"
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
      MICON           =   "frmVigSanitaria.frx":125B
      PICN            =   "frmVigSanitaria.frx":1277
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
      Left            =   3870
      TabIndex        =   13
      ToolTipText     =   "Editar a Tabela"
      Top             =   4290
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmVigSanitaria.frx":12E5
      PICN            =   "frmVigSanitaria.frx":1301
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
      Caption         =   "Aliquota:"
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
      TabIndex        =   10
      Top             =   5640
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SubItem:"
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
      TabIndex        =   9
      Top             =   5280
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item.....:"
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
      TabIndex        =   8
      Top             =   4920
      Width           =   1035
   End
   Begin VB.Label lblAliq 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1950
      TabIndex        =   7
      Top             =   4380
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Aliquota (Ufir):"
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
      TabIndex        =   1
      Top             =   4380
      Width           =   1755
   End
   Begin VB.Menu mnuEditor 
      Caption         =   "&Editor"
      Visible         =   0   'False
      Begin VB.Menu mnuNew 
         Caption         =   "&Novo Item"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Alterar Item"
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Novo &SubItem"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "&Excluir Item"
      End
   End
End
Attribute VB_Name = "frmVigSanitaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String, NomeForm As String

Public Property Let sForm(sNomeForm As String)
    NomeForm = sNomeForm
End Property

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancelEdit_Click()
EndEdit
End Sub

Private Sub cmdConsultar_Click()
Dim x As Integer, Achou As Boolean
Dim nQtde As Variant

If tvVig.SelectedItem.Children > 0 Then
    MsgBox "Selecione um dos SubItens da categoria selecionada.", vbExclamation, "Atenção"
    Exit Sub
End If

If NomeForm = "frmCadMob" Then
    If Len(tvVig.SelectedItem.Key) > 4 Then
       Achou = False
       With frmCadMob.grdVS
          For x = 1 To .Rows - 1
              If .TextMatrix(x, 0) = Mid$(tvVig.SelectedItem.Key, 2, 3) And .TextMatrix(x, 1) = Right$(tvVig.SelectedItem.Key, 3) Then
                 Achou = True
                 Exit For
              End If
          Next
       End With
       If Not Achou Then
          nQtde = InputBox("Digite a Qtde  para esta atividade.", "Quantidade", "1")
          If Not IsNumeric(nQtde) Then nQtde = 0
          If CDbl(nQtde) = 0 Then nQtde = 1
          frmCadMob.grdVS.AddItem Mid$(tvVig.SelectedItem.Key, 2, 3) & Chr(9) & Right$(tvVig.SelectedItem.Key, 3) & Chr(9) & tvVig.SelectedItem.Parent.text & " - " & tvVig.SelectedItem.text & Chr(9) & nQtde & Chr(9) & FormatNumber(tvVig.SelectedItem.Tag, 2)
          CodEmpresa = 0
       Else
          MsgBox "Esta atividade já foi atribuida à empresa.", vbExclamation, "Atenção"
       End If
    Else
       Achou = False
       With frmCadMob.grdVS
          For x = 1 To .Rows - 1
              If .TextMatrix(x, 0) = Mid$(tvVig.SelectedItem.Key, 2, 3) And .TextMatrix(x, 1) = "000" Then
                 Achou = True
                 Exit For
              End If
          Next
       End With
       If Not Achou Then
          nQtde = InputBox("Digite a Qtde  para esta atividade.", "Quantidade", "1")
          If Not IsNumeric(nQtde) Then nQtde = 0
          If CDbl(nQtde) = 0 Then nQtde = 1
          frmCadMob.grdVS.AddItem Mid$(tvVig.SelectedItem.Key, 2, 3) & Chr(9) & "000" & Chr(9) & tvVig.SelectedItem.text & Chr(9) & nQtde & Chr(9) & FormatNumber(tvVig.SelectedItem.Tag, 2)
          CodEmpresa = 0
       Else
          MsgBox "Esta atividade já foi atribuida à empresa.", vbExclamation, "Atenção"
       End If
    End If
    Unload Me
End If

End Sub

Private Sub cmdEdit_Click()

PopupMenu mnuEditor

End Sub

Private Sub cmdGravar_Click()
Dim nLastCod As Integer, nLastSubCod As Integer

Select Case Evento
    Case "NI"
        If Trim$(txtItem.text) = "" Then
           MsgBox "Digite a Descrição do Novo Item.", vbExclamation, "Atenção"
           Exit Sub
        End If
        Sql = "SELECT MAX(CODVIGSANIT) AS COD FROM VIGSANITARIA"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nLastCod = RdoAux!COD + 1
        nLastSubCod = 1
        Set NodX = tvVig.Nodes.Add(, , "C" & Format(nLastCod, "000"), txtItem.text, 2)
        tvVig.Nodes("C" & Format(nLastCod, "000")).Tag = FormatNumber(txtAliq, 2)
        Sql = "INSERT VIGSANITARIA (CODVIGSANIT,SUBCODVIGSANIT,DESCVIGSANITARIA,VALORALIQ) VALUES("
        Sql = Sql & nLastCod & "," & 0 & ",'" & Mask(txtItem.text) & "'," & Virg2Ponto(txtAliq.text) & ")"
        cn.Execute Sql, rdExecDirect
    Case "NS"
        If Trim$(txtSub.text) = "" Then
           MsgBox "Digite a Descrição do Novo SubItem.", vbExclamation, "Atenção"
           Exit Sub
        End If
        If Val(txtAliq.text) = 0 Then
           MsgBox "Digite o Valor da Aliquota.", vbExclamation, "Atenção"
           Exit Sub
        End If
        nLastCod = Val(Right$(tvVig.SelectedItem.Key, 3))
        Sql = "SELECT MAX(SUBCODVIGSANIT) AS SUBCOD FROM VIGSANITARIA WHERE CODVIGSANIT=" & nLastCod
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!SUBCOD) Then
            nLastSubCod = 1
        Else
            nLastSubCod = RdoAux!SUBCOD + 1
        End If
        Set NodX = tvVig.Nodes.Add("C" & Format(nLastCod, "000"), tvwChild, "C" & Format(nLastCod, "000") & "S" & Format(nLastSubCod, "000"), txtSub.text, 1)
        tvVig.Nodes("C" & Format(nLastCod, "000") & "S" & Format(nLastSubCod, "000")).ForeColor = vbBlue
        tvVig.Nodes("C" & Format(nLastCod, "000") & "S" & Format(nLastSubCod, "000")).Tag = FormatNumber(txtAliq, 2)
        Sql = "INSERT VIGSANITARIA (CODVIGSANIT,SUBCODVIGSANIT,DESCVIGSANITARIA,VALORALIQ) VALUES("
        Sql = Sql & nLastCod & "," & nLastSubCod & ",'" & Mask(txtSub.text) & "'," & Virg2Ponto(txtAliq.text) & ")"
        cn.Execute Sql, rdExecDirect
        tvVig.Nodes("C" & Format(nLastCod, "000") & "S" & Format(nLastSubCod, "000")).EnsureVisible
    Case "EI"
        If Trim$(txtItem.text) = "" Then
           MsgBox "Digite a Descrição do Item.", vbExclamation, "Atenção"
           Exit Sub
        End If
        If Val(txtAliq.text) = 0 Then
           MsgBox "Digite o Valor da Aliquota.", vbExclamation, "Atenção"
           Exit Sub
        End If
        tvVig.Nodes(tvVig.SelectedItem.Index).text = txtItem.text
        tvVig.Nodes(tvVig.SelectedItem.Index).Tag = FormatNumber(txtAliq, 2)
        
        nLastCod = Val(Right$(tvVig.SelectedItem.Key, 3))
        nLastSubCod = 0
        
        Sql = "UPDATE VIGSANITARIA SET DESCVIGSANITARIA='" & txtItem.text & "', "
        Sql = Sql & "VALORALIQ=" & Virg2Ponto(txtAliq.text) & " WHERE CODVIGSANIT=" & nLastCod & " AND SUBCODVIGSANIT=0"
        cn.Execute Sql, rdExecDirect
    Case "ES"
        If Trim$(txtSub.text) = "" Then
           MsgBox "Digite a Descrição do SubItem.", vbExclamation, "Atenção"
           Exit Sub
        End If
        If Val(txtAliq.text) = 0 Then
           MsgBox "Digite o Valor da Aliquota.", vbExclamation, "Atenção"
           Exit Sub
        End If
        
        nLastCod = Val(Mid$(tvVig.SelectedItem.Key, 2, 3))
        nLastSubCod = Val(Right$(tvVig.SelectedItem.Key, 3))
        
        tvVig.Nodes(tvVig.SelectedItem.Index).text = txtSub.text
        tvVig.Nodes(tvVig.SelectedItem.Index).Tag = FormatNumber(txtAliq, 2)
        
        Sql = "UPDATE VIGSANITARIA SET DESCVIGSANITARIA='" & txtSub.text & "', "
        Sql = Sql & "VALORALIQ=" & Virg2Ponto(txtAliq.text) & " WHERE CODVIGSANIT=" & nLastCod & " AND SUBCODVIGSANIT=" & nLastSubCod
        cn.Execute Sql, rdExecDirect
End Select

EndEdit
End Sub


Private Sub cmdPrint_Click()
'EXIBE RELATORIO
frmReport.ShowReport "ALIQATIVIDADEVS", frmMdi.hwnd, Me.hwnd

End Sub

Private Sub Form_Load()
Ocupado
Me.Height = 5190
Centraliza Me
If NomeForm = "" Then
   cmdConsultar.Enabled = False
Else
   cmdConsultar.Enabled = True
End If
CarregaLista
End Sub

Private Sub CarregaLista()

Sql = "SELECT CODVIGSANIT,SUBCODVIGSANIT,DESCVIGSANITARIA,VALORALIQ FROM VIGSANITARIA "
Sql = Sql & "ORDER BY CODVIGSANIT,SUBCODVIGSANIT"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)


With RdoAux
    Do Until .EOF
       If !SUBCODVIGSANIT = 0 Then
           Set NodX = tvVig.Nodes.Add(, , "C" & Format(!CODVIGSANIT, "000"), !DESCVIGSANITARIA, 2)
           tvVig.Nodes("C" & Format(!CODVIGSANIT, "000")).Tag = FormatNumber(!VALORALIQ, 2)
       Else
           Set NodX = tvVig.Nodes.Add("C" & Format(!CODVIGSANIT, "000"), tvwChild, "C" & Format(!CODVIGSANIT, "000") & "S" & Format(!SUBCODVIGSANIT, "000"), !DESCVIGSANITARIA, 1)
           tvVig.Nodes("C" & Format(!CODVIGSANIT, "000") & "S" & Format(!SUBCODVIGSANIT, "000")).ForeColor = vbBlue
           tvVig.Nodes("C" & Format(!CODVIGSANIT, "000") & "S" & Format(!SUBCODVIGSANIT, "000")).Tag = FormatNumber(!VALORALIQ, 2)
       End If
      .MoveNext
    Loop
   .Close
End With

'Geral
With tvVig
    For x = 1 To .Nodes.Count
       .Nodes(x).EnsureVisible
    Next
    .Nodes(1).Selected = True
End With
Liberado
End Sub

Private Sub Form_Unload(Cancel As Integer)
NomeForm = ""
End Sub

Private Sub mnuDel_Click()
Dim nLastCod As Integer, nLastSubCod As Integer

If tvVig.SelectedItem.ForeColor <> vbBlue And tvVig.SelectedItem.Children > 0 Then
    MsgBox "Exclua todos os SubItens antes de excluir o Item principal.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Excluir o Item: " & tvVig.SelectedItem.text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    If Len(tvVig.SelectedItem.Key) = 4 Then
        nLastCod = Val(Right$(tvVig.SelectedItem.Key, 3))
        nLastSubCod = 0
    Else
        nLastCod = Val(Mid$(tvVig.SelectedItem.Key, 2, 3))
        nLastSubCod = Val(Right$(tvVig.SelectedItem.Key, 3))
    End If
    Sql = "DELETE FROM VIGSANITARIA WHERE CODVIGSANIT=" & nLastCod & " AND SUBCODVIGSANIT=" & nLastSubCod
    cn.Execute Sql, rdExecDirect
    tvVig.Nodes.Remove (tvVig.SelectedItem.Index)
    End If

End Sub

Private Sub mnuEdit_Click()

IniEdit
If tvVig.SelectedItem.ForeColor = vbBlue Then
    'subitem
    Evento = "ES"
    txtItem.text = tvVig.SelectedItem.Parent.text
    txtAliq.text = tvVig.SelectedItem.Tag
    txtSub.text = tvVig.SelectedItem.text
    txtSub.Enabled = True
    txtSub.BackColor = Branco
    txtItem.Locked = True
    txtItem.BackColor = Kde
    txtSub.SetFocus
Else
    'item
    Evento = "EI"
    txtItem.text = tvVig.SelectedItem.text
    txtAliq.text = tvVig.SelectedItem.Tag
    txtSub.text = ""
    txtSub.Enabled = False
    txtSub.BackColor = Kde
    txtItem.Locked = False
    txtItem.BackColor = Branco
    txtItem.SetFocus
End If

End Sub

Private Sub mnuNew_Click()

Evento = "NI"
IniEdit
txtItem.text = ""
txtItem.Locked = False
txtItem.BackColor = Branco
txtAliq.text = 0
txtSub.text = ""
txtSub.Enabled = False
txtSub.BackColor = Kde
txtItem.SetFocus

End Sub

Private Sub IniEdit()

Me.Height = 6390
Me.Top = Me.Top - 450
cmdEdit.Enabled = False
cmdConsultar.Enabled = False
cmdCancel.Enabled = False
tvVig.Enabled = False
End Sub

Private Sub EndEdit()

Me.Height = 5190
Me.Top = Me.Top + 450
cmdEdit.Enabled = True
cmdConsultar.Enabled = True
cmdCancel.Enabled = True
tvVig.Enabled = True
End Sub

Private Sub mnuSub_Click()

If tvVig.SelectedItem.ForeColor = vbBlue Then
   MsgBox "Selecione um item principal para inserir um subitem.", vbExclamation, "Atenção"
   Exit Sub
End If

Evento = "NS"
IniEdit
txtItem.text = tvVig.SelectedItem.text
txtItem.Locked = True
txtItem.BackColor = Kde
txtAliq.text = 0
txtSub.text = ""
txtSub.Enabled = True
txtSub.BackColor = Branco
txtSub.SetFocus

End Sub

Private Sub tvVig_NodeClick(ByVal Node As MSComctlLib.Node)
lblAliq.Caption = Node.Tag
End Sub

Private Sub txtAliq_GotFocus()
txtAliq.SelStart = 0
txtAliq.SelLength = Len(txtAliq.text)
End Sub

Private Sub txtAliq_KeyPress(KeyAscii As Integer)
Tweak txtAliq, KeyAscii, DecimalPositive
End Sub

