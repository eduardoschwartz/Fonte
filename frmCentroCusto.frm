VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCentroCusto 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centro de Custos"
   ClientHeight    =   7725
   ClientLeft      =   3000
   ClientTop       =   2115
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   5700
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   4530
      TabIndex        =   19
      ToolTipText     =   "Sair da Tela"
      Top             =   7380
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
      MICON           =   "frmCentroCusto.frx":0000
      PICN            =   "frmCentroCusto.frx":001C
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
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4530
      TabIndex        =   14
      ToolTipText     =   "Cancelar Edição"
      Top             =   7380
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
      MICON           =   "frmCentroCusto.frx":008A
      PICN            =   "frmCentroCusto.frx":00A6
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
      Left            =   90
      TabIndex        =   15
      ToolTipText     =   "Novo Registro"
      Top             =   7380
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmCentroCusto.frx":0200
      PICN            =   "frmCentroCusto.frx":021C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1140
      TabIndex        =   16
      ToolTipText     =   "Editar Registro"
      Top             =   7380
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
      MICON           =   "frmCentroCusto.frx":0376
      PICN            =   "frmCentroCusto.frx":0392
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
      Left            =   2190
      TabIndex        =   17
      ToolTipText     =   "Excluir Registro"
      Top             =   7380
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MICON           =   "frmCentroCusto.frx":04EC
      PICN            =   "frmCentroCusto.frx":0508
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
      Left            =   3480
      TabIndex        =   18
      ToolTipText     =   "Gravar os Dados"
      Top             =   7380
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
      MICON           =   "frmCentroCusto.frx":05AA
      PICN            =   "frmCentroCusto.frx":05C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ImageList ImCC 
      Left            =   2580
      Top             =   2250
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentroCusto.frx":096B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentroCusto.frx":0AC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCentroCusto.frx":0BAF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1545
      Left            =   30
      TabIndex        =   8
      Top             =   5730
      Width           =   5655
      Begin VB.CheckBox chkAtivo 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Ativo"
         Height          =   255
         Left            =   4560
         TabIndex        =   7
         Top             =   1170
         Width           =   1005
      End
      Begin VB.TextBox txtPertence 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   4515
      End
      Begin VB.TextBox txtFone 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1515
      End
      Begin VB.TextBox txtSigla 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3390
         MaxLength       =   6
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   510
         Width           =   1005
      End
      Begin VB.CheckBox chkPref 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Prefeitura"
         Height          =   255
         Left            =   4590
         TabIndex        =   4
         Top             =   510
         Width           =   1005
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         MaxLength       =   50
         TabIndex        =   5
         Top             =   840
         Width           =   4515
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1020
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   510
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pertence....:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone....:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   1170
         Width           =   885
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sigla......:"
         Height          =   195
         Index           =   1
         Left            =   2670
         TabIndex        =   11
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição..:"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   10
         Top             =   870
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código.......:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   555
         Width           =   915
      End
   End
   Begin MSComctlLib.TreeView tvCC 
      Height          =   5670
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   10001
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   794
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ImCC"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmCentroCusto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean
Dim NodX As Object

Private Sub cmdAlterar_Click()
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    sOldDesc = UCase$(txtDesc.Text)
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Eventos "INICIAR"
    Evento = ""
    tvCC.SetFocus
    On Error Resume Next
    SendKeys "{DOWN}": SendKeys "{UP}"
End Sub

Private Sub cmdExcluir_Click()
If Val(txtCod.Text) = 0 Then Exit Sub
    
If tvCC.SelectedItem.Children > 0 Then
    MsgBox "Exclua os locais dependentes primeiro.", vbExclamation, "Atenção"
    Exit Sub
End If
    
If MsgBox("Excluir o Local " & txtDesc.Text & " ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
   Sql = "DELETE FROM CENTROCUSTO WHERE CODIGO=" & Val(txtCod.Text)
   cn.Execute Sql, rdExecDirect
   Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & UCase$(txtDesc.Text)
   Limpa
   tvCC.Nodes.Remove (tvCC.SelectedItem.Index)
End If

End Sub

Private Sub cmdGravar_Click()
    If UCase$(txtDesc.Text) = "" Then
       MsgBox "Favor digitar a descrição do local.", vbExclamation, "Atenção"
       txtDesc.SetFocus
       Exit Sub
    End If
    If UCase$(txtSigla.Text) = "" Then
       MsgBox "Favor digitar a sigla do local.", vbExclamation, "Atenção"
       txtDesc.SetFocus
       Exit Sub
    End If
    Grava
    Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
    txtPertence.Text = Right$(tvCC.SelectedItem.Key, 3) & "-" & tvCC.SelectedItem.Text
End Sub

Private Sub cmdSair_Click()
   Unload Me
End Sub

Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()

Centraliza Me
sRet = RetEventUserForm(Me.Name)
'chkAtivo.Value = 1
CarregaLista

Eventos "INICIAR"

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmCentroCusto
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Locked = True
       End If
   Next
   tvCC.Enabled = True
'   chkAtivo.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmCentroCusto
       If TypeOf Ct Is TextBox Then
          If Ct.Name <> "txtPertence" Then
             Ct.BackColor = Branco
             Ct.Locked = False
          End If
       End If
   Next
   tvCC.Enabled = False
  ' chkAtivo.Enabled = False
   txtCod.BackColor = Kde
   txtCod.Locked = True
   txtSigla.SetFocus
End If

FormHagana

End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDesc.Text = ""
txtSigla.Text = ""
chkPref.value = 0
txtFone.Text = ""
End Sub

Private Sub CarregaLista()
Dim nVinculo As Integer, nCodigo As Integer, bAchou As Boolean, Y As Integer, bContinua As Boolean
Dim aKeys() As String
ReDim aKeys(0)
Inicio:
For x = 1 To tvCC.Nodes.Count
    tvCC.Nodes.Remove (x)
    GoTo Inicio
Next

Set NodX = tvCC.Nodes.Add(, , "CC", "CENTRO DE CUSTOS", 3)

Sql = "Select CODIGO,DESCRICAO,VINCULO FROM CENTROCUSTO "
Sql = Sql & " WHERE  VINCULO=0 ORDER BY DESCRICAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

With RdoAux
    Do Until .EOF
       Set NodX = tvCC.Nodes.Add("CC", tvwChild, "C" & Format(!Codigo, "000"), !descricao, 2)
       ReDim Preserve aKeys(UBound(aKeys) + 1)
       aKeys(UBound(aKeys)) = "C" & Format(!Codigo, "000")
      .MoveNext
    Loop
   .Close
End With


Inicio2:
bContinua = False
For x = 1 To tvCC.Nodes.Count
    nVinculo = Val(Right$(tvCC.Nodes(x).Key, 3))
    Sql = "Select CODIGO,DESCRICAO,VINCULO FROM CENTROCUSTO "
    Sql = Sql & " WHERE  VINCULO=" & nVinculo & " AND VINCULO>0  ORDER BY DESCRICAO"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        
        Do Until .EOF
            
            
            bAchou = False
            For Y = 1 To UBound(aKeys)
                If aKeys(Y) = "C" & Format(!Codigo, "000") Then
                    bAchou = True
                    Exit For
                End If
            Next
            If Not bAchou Then
                Set NodX = tvCC.Nodes.Add(tvCC.Nodes(x).Key, tvwChild, "C" & Format(!Codigo, "000"), !descricao, 1)
                ReDim Preserve aKeys(UBound(aKeys) + 1)
                aKeys(UBound(aKeys)) = "C" & Format(!Codigo, "000")
                bContinua = True
            End If
           .MoveNext
        Loop
       .Close
    End With
proximo:
Next

If bContinua Then GoTo Inicio2


'Sql = "Select CODIGO,DESCRICAO,VINCULO FROM CENTROCUSTO "
'Sql = Sql & " WHERE  VINCULO>0 AND VINCULO<CODIGO  ORDER BY DESCRICAO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'On Error Resume Next
'With RdoAux
'    Do Until .EOF
'       Set NodX = tvCC.Nodes.Add("C" & Format(!VINCULO, "000"), tvwChild, "C" & Format(!Codigo, "000"), !DESCRICAO, 1)
'      .MoveNext
'    Loop
'   .Close
'End With

'Sql = "Select CODIGO,DESCRICAO,VINCULO FROM CENTROCUSTO "
'Sql = Sql & " WHERE  VINCULO>0 AND VINCULO>CODIGO ORDER BY DESCRICAO"
'Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

'On Error Resume Next
'With RdoAux
'    Do Until .EOF
'       Set NodX = tvCC.Nodes.Add("C" & Format(!VINCULO, "000"), tvwChild, "C" & Format(!Codigo, "000"), !DESCRICAO, 1)
'      .MoveNext
'    Loop
'   .Close
' End With

For x = 1 To tvCC.Nodes.Count
   tvCC.Nodes(x).EnsureVisible
Next
tvCC.Nodes(1).EnsureVisible
On Error Resume Next
SendKeys "{DOWN}": SendKeys "{UP}"

End Sub


Private Sub Grava()
Dim nCodNovo As Integer, nParent As Integer

nParent = Val(Left$(txtPertence.Text, 3))

If Evento = "Novo" Then
    Sql = "SELECT MAX(CODIGO) AS MAXIMO FROM CENTROCUSTO WHERE CODIGO < 900"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        nCodNovo = !maximo + 1
       .Close
    End With
    
    If nParent = 0 Then
        Set NodX = tvCC.Nodes.Add("CC", tvwChild, "C" & Format(nCodNovo, "000"), txtDesc.Text, 2)
    Else
        Set NodX = tvCC.Nodes.Add("C" & Format(nParent, "000"), tvwChild, "C" & Format(nCodNovo, "000"), txtDesc.Text, 1)
    End If

    Sql = "INSERT CENTROCUSTO (CODIGO,DESCRICAO,ATIVO,PREFEITURA,SIGLA,"
    Sql = Sql & "TELEFONE,VINCULO) VALUES(" & nCodNovo & ",'" & UCase$(Mask(txtDesc.Text)) & "',"
    Sql = Sql & chkAtivo.value & "," & chkPref.value & ",'" & Mask(txtSigla.Text) & "','"
    Sql = Sql & Mask(txtFone.Text) & "'," & nParent & ")"
Else
    tvCC.SelectedItem.Text = txtDesc.Text
    Sql = "UPDATE CENTROCUSTO SET DESCRICAO='" & UCase$(Mask(txtDesc.Text)) & "',PREFEITURA=" & chkPref.value & ",ATIVO=" & chkAtivo.value & ","
    Sql = Sql & "SIGLA='" & Mask(txtSigla.Text) & "',TELEFONE='" & Mask(txtFone.Text) & "' WHERE CODIGO=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect
      
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If cmdNovo.Visible = True Then
      cmdNovo_Click
   Else
      cmdGravar_Click
   End If
End If
End Sub


Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub tvCC_NodeClick(ByVal Node As MSComctlLib.Node)
Dim nParent As Integer, sParent As String

Limpa
On Error Resume Next
Err.Number = 0
nParent = Val(Right$(Node.Parent.Key, 3))
sParent = Node.Parent.Text
If Err.Number > 0 Then
   nParent = 0
   sParent = "CENTRO DE CUSTOS"
End If

txtPertence.Text = Format(nParent, "000") & "-" & sParent
Sql = "SELECT * FROM CENTROCUSTO WHERE CODIGO=" & Val(Right$(Node.Key, 3))
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        txtCod.Text = !Codigo
        txtDesc.Text = !descricao
        txtSigla.Text = SubNull(!Sigla)
        txtFone.Text = SubNull(!telefone)
        txtClass.Text = SubNull(!CLASSIFICACAO)
        chkPref.value = IIf(!PREFEITURA, 1, 0)
        chkAtivo.value = IIf(!Ativo, 1, 0)
    End If
End With

End Sub
