VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmIssRetido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empresas com ISS retido na fonte"
   ClientHeight    =   4350
   ClientLeft      =   6210
   ClientTop       =   4845
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4350
   ScaleWidth      =   6705
   Begin MSComctlLib.ListView lvTmp 
      Height          =   3855
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   6800
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Razão Social"
         Object.Width           =   9173
      EndProperty
   End
   Begin prjChameleon.chameleonButton cmdDel 
      Height          =   315
      Left            =   4245
      TabIndex        =   2
      ToolTipText     =   "Remover empresa"
      Top             =   3960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Remover"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmIssRetido.frx":0000
      PICN            =   "frmIssRetido.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5460
      TabIndex        =   3
      ToolTipText     =   "Sair da Tela"
      Top             =   3960
      Width           =   1155
      _ExtentX        =   2037
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmIssRetido.frx":0176
      PICN            =   "frmIssRetido.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAdd 
      Height          =   315
      Left            =   3030
      TabIndex        =   1
      ToolTipText     =   "Adicionar empresa"
      Top             =   3960
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Adicionar"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmIssRetido.frx":0200
      PICN            =   "frmIssRetido.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "frmIssRetido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String

Private Sub cmdAdd_Click()
'Sql = "SELECT distinct cidadao.codcidadao,Cidadao.nomecidadao From Cidadao INNER JOIN debitotributo ON (cidadao.codcidadao = debitotributo.codreduzido) "
'Sql = Sql & "WHERE codreduzido >= 500000 and debitotributo.anoexercicio = 2007 AND debitotributo.CodTributo = 502 order by codcidadao"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'With RdoAux
'    Do Until .EOF
'        Sql = "INSERT ISSRETIDO (CODREDUZIDO) VALUES(" & !CodCidadao & ")"
'        cn.Execute Sql, rdExecDirect
'       .MoveNext
'    Loop
'   .Close
'End With

Dim z As Variant, x As Integer

z = InputBox("Digite o Código da empresa.", "Inclusão de empresas com retenção de ISS")
If z = "" Then Exit Sub
If Val(z) < 500000 Or Val(z) > 600000 Then
    MsgBox "Código inválido.", vbCritical, "Atenção"
Else
    For x = 1 To lvTmp.ListItems.Count
        If Val(lvTmp.ListItems(x).text) = Val(z) Then
            MsgBox "Código já existe na lista.", vbCritical, "Atenção"
            Exit Sub
        End If
    Next

    Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & Val(z)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Desesja adicionar a empresa " & Val(z) & " - " & !NOMECIDADAO & " à lista?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                Sql = "INSERT ISSRETIDO (CODREDUZIDO) VALUES(" & Val(z) & ")"
                cn.Execute Sql, rdExecDirect
                CarregaLista
            End If
        Else
            MsgBox "Código não cadastrado.", vbCritical, "Atenção"
        End If
       .Close
    End With
End If

End Sub

Private Sub cmdDel_Click()

Dim nCodReduz As Long
If Val(lvTmp.ListItems(lvTmp.SelectedItem.Index).text) > 0 Then
    nCodReduz = Val(lvTmp.ListItems(lvTmp.SelectedItem.Index).text)
    If MsgBox("Desesja remover a empresa " & nCodReduz & " - " & lvTmp.ListItems(lvTmp.SelectedItem.Index).SubItems(1) & " da lista?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        Sql = "DELETE FROM ISSRETIDO WHERE CODREDUZIDO=" & Val(z)
        cn.Execute Sql, rdExecDirect
        CarregaLista
    End If
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
CarregaLista
End Sub

Private Sub CarregaLista()
Dim itmX As ListItem
Dim z As Long
z = SendMessage(lvTmp.hwnd, LVM_DELETEALLITEMS, 0, 0)

Ocupado

Sql = "SELECT DISTINCT  issretido.codreduzido,Cidadao.nomecidadao From Cidadao INNER JOIN issretido ON (cidadao.codcidadao = issretido.codreduzido) ORDER BY NOMECIDADAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvTmp.ListItems.Add(, , !CODREDUZIDO)
       itmX.SubItems(1) = !NOMECIDADAO
      .MoveNext
    Loop
   .Close
End With

Liberado
End Sub

Private Sub lvTmp_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvTmp.SortKey = ColumnHeader.Position - 1
lvTmp.Sorted = True
lvTmp.SortOrder = lvwAscending

End Sub
